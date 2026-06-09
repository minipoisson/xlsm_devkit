Attribute VB_Name = "xlsm_devkit"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
Private Declare PtrSafe Function GetLocaleInfoA Lib "kernel32" ( _
    ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long
#Else
Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" ( _
    ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long
#End If

' Core module. No dependencies on optional feature modules.
' Entry points callable from feature modules:
'   CallInitDevMode  - create DEV_ copy of this workbook and inject devkit_* modules from src/
'   CallSaveAsRelease    - strip devkit modules and save production copy (call from DEV_ workbook)
'   ExportAllModulesFormsSheetMaps - export all VBA modules, forms, and sheet maps to src/ and sheet/?
'   ImportAllModulesFormsSheetMaps - import all VBA modules, forms, and sheet maps from src/ and sheet/?
'   CallExportAllComponents  - export all VBA components and forms to src/
'   CallImportAllComponents  - import all VBA components and forms from src/
'   CallExportAllSheetMapsToMD - export all sheet maps to sheet/ as Markdown files
'   CallImportAllSheetMapsFromMD - import all sheet maps from sheet/ Markdown files
'   ExportSheetToMDFile(ws as Worksheet, filePath as String) - export a single sheet map to a specified Markdown file

Private Const MODULE_NAME As String = "xlsm_devkit"
' Set True to skip optional devkit feature modules on import (recommended for users).
' Set False when developing devkit optional modules themselves.
Private Const SKIP_DEVKIT_MODULES As Boolean = True
Private Const IMPORT_DIAGNOSTICS_ENABLED As Boolean = False
Private Const STRING_BUILDER_CHUNK_SIZE As Long = 1000

' NOTE: This module itself is NOT imported by ImportAllComponents()
' because a module cannot delete or overwrite itself.

Public g_LangCode  As String
Public g_LangCache As Object
Public g_EnCache   As Object
Private m_LangCodeDetected As String
Private devkitImportLogPath As String

Private Type DevkitStringBuilder
    Parts() As String
    Count As Long
    Capacity As Long
End Type

Public Sub CallInitDevMode()
    InitDevMode
End Sub

Public Sub CallSaveAsRelease()
    SaveAsRelease
End Sub

Public Sub ExportAllModulesFormsSheetMaps()
    If MsgBox(t("msg.export_all_confirm", "Export all modules, forms, and sheet maps to src/ and sheet/?"), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ExportAllComponents True
    ExportAllSheetMapsToMD True
End Sub

Public Sub ImportAllModulesFormsSheetMaps()
    If MsgBox(t("msg.import_all_confirm", "Import all modules, forms, and sheet maps from src/ and sheet/?"), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    Dim sheetFolderPath As String
    Dim fso As Object
    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sheetFolderPath) Then
        If AbortIfProtectedImportSheets(sheetFolderPath, fso) Then Exit Sub
    End If
    ImportAllComponents True
    ImportAllSheetMapsFromMD True
End Sub

Public Sub CallExportAllComponents(Optional skipConfirm As Boolean = False)
    ExportAllComponents skipConfirm
End Sub

Public Sub CallImportAllComponents()
    ImportAllComponents
End Sub

Public Sub CallExportAllSheetMapsToMD()
    ExportAllSheetMapsToMD
End Sub

Public Sub CallImportAllSheetMapsFromMD()
    ImportAllSheetMapsFromMD
End Sub

Private Function CheckVBProjectAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox t("msg.trust_vba_required", "Please enable 'Trust access to the VBA project object model' in Excel macro settings.")
        Exit Function
    End If
    On Error GoTo 0
    CheckVBProjectAccess = True
End Function

Sub ExportAllComponents(Optional skipConfirm As Boolean = False)
    If Not CheckVBProjectAccess() Then Exit Sub
    If Not skipConfirm Then
        If MsgBox(t("msg.export_components_confirm", "Export all modules and forms to src/?"), vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim dirPath As String: dirPath = ThisWorkbook.Path & "\src"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dirPath) Then fso.CreateFolder dirPath

    Dim comp As Object, expPath As String, frxPath As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If SKIP_DEVKIT_MODULES And (LCase(Left(comp.Name, 7)) = "devkit_" _
        Or LCase(Left(comp.Name, 11)) = "xlsm_devkit") Then GoTo lblNext
        
        If comp.Type = 3 Then
            expPath = dirPath & "\" & comp.Name & ".frm"
            frxPath = dirPath & "\" & comp.Name & ".frx"
            If fso.FileExists(frxPath) Then fso.DeleteFile frxPath, True
        Else
            expPath = dirPath & "\" & comp.Name & ".bas"
        End If
        comp.Export expPath
        ConvertEncoding expPath, GetSystemAnsiCharset(), "UTF-8", (comp.Type = 3)
lblNext:
    Next comp
    MsgBox t("msg.export_complete", "Export complete.")
End Sub

Sub ImportAllComponents(Optional skipConfirm As Boolean = False)
    If Not CheckVBProjectAccess() Then Exit Sub
    If Not skipConfirm Then
        If MsgBox(t("msg.import_components_confirm", "Import all modules and forms from src/?"), vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String: srcPath = ThisWorkbook.Path & "\src"
    If Not fso.FolderExists(srcPath) Then
        MsgBox Fmt(t("msg.source_not_found", "Source folder not found: {0}"), srcPath), vbExclamation
        Exit Sub
    End If

    Dim successCount As Long, failCount As Long, skipCount As Long
    ' Move capture is locked when either the registry state exists (Phase 1->2)
    ' or devkit_frmMoveWait is still loaded (e.g. Import called from devkit_frmInstruction while
    ' devkit_frmMoveWait is still on screen ? replacing devkit_Move.bas would reset the VBA runtime
    ' and crash Excel because devkit_Move.bas is on the active call stack).
    Dim isMoveLocked As Boolean
    isMoveLocked = IsFormCurrentlyLoaded("devkit_frmMoveWait") Or _
                   (Len(GetSetting("xlsm_devkit", "MoveCapture", "State", "")) > 0)

    Dim srcFile As Object, ext As String, compName As String
    For Each srcFile In fso.GetFolder(srcPath).Files
        ext = LCase(fso.GetExtensionName(srcFile.Name))
        compName = Left(srcFile.Name, Len(srcFile.Name) - Len(ext) - 1)
        If ext = "bas" Then
            If LCase(compName) = LCase(MODULE_NAME) Then GoTo lblNext              ' xlsm_devkit (self)
            If SKIP_DEVKIT_MODULES And LCase(Left(compName, 7)) = "devkit_" Then GoTo lblNext  ' optional devkit modules
            If LCase(compName) = "devkit_move" And isMoveLocked Then         ' devkit_Move on call stack
                skipCount = skipCount + 1: GoTo lblNext
            End If
            If ComponentExists(compName) Then
                If ReplaceExistingComponentCodeFromFile(compName, srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            Else
                If ImportNewComponentFromFile(srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            End If
        ElseIf ext = "frm" Then
            If SKIP_DEVKIT_MODULES And LCase(Left(compName, 7)) = "devkit_" Then GoTo lblNext  ' optional devkit forms
            If IsFormCurrentlyLoaded(compName) Then                  ' skip any loaded form
                skipCount = skipCount + 1: GoTo lblNext
            End If
            If ImportForm(compName, srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
        End If
lblNext:
    Next srcFile

    Dim msg As String
    msg = t("msg.import_complete", "Import complete.") & vbLf & "Success: " & successCount & vbLf & "Failed: " & failCount
    If skipCount > 0 Then msg = msg & vbLf & "Skipped (in use): " & skipCount
    MsgBox msg, IIf(failCount = 0, vbInformation, vbExclamation)
End Sub

Private Sub ConvertEncoding(filePath As String, fromCharset As String, toCharset As String, _
                            Optional normalizeStartUpPos As Boolean = False)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Open
    st.Type = 2
    st.Charset = fromCharset
    st.LoadFromFile filePath
    Dim content As String
    content = st.ReadText
    st.Close

    If normalizeStartUpPos Then content = NormalizeStartUpPosition(content)

    If LCase(toCharset) = "utf-8" Then
        SaveAsUTF8 filePath, content
    Else
        st.Open
        st.Type = 2
        st.Charset = toCharset
        st.WriteText content
        st.SaveToFile filePath, 2
        st.Close
    End If
End Sub

Private Function ReplaceExistingComponentCodeFromFile(modName As String, filePath As String) As Boolean
    Dim tempName As String
    tempName = BuildUniqueTempModuleName(modName)
    
    Dim tempPath As String
    tempPath = ThisWorkbook.Path & "\src\" & tempName & ".bas"
    
    Dim tempComp As Object
    Dim targetComp As Object
    Dim codeText As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler

    ' Update existing code via a temporary module to avoid removing the target component.
    fso.CopyFile filePath, tempPath
    RenameModuleInFile tempPath, modName, tempName

    ' VBComponents.Import expects system ANSI encoding.
    ConvertEncoding tempPath, "UTF-8", GetSystemAnsiCharset()
    ThisWorkbook.VBProject.VBComponents.Import tempPath
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True

    Set tempComp = ThisWorkbook.VBProject.VBComponents(tempName)
    If tempComp.CodeModule.CountOfLines > 0 Then
        codeText = tempComp.CodeModule.lines(1, tempComp.CodeModule.CountOfLines)
        codeText = StripAttributeLines(codeText)
    End If

    Set targetComp = ThisWorkbook.VBProject.VBComponents(modName)
    With targetComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(codeText) > 0 Then .AddFromString codeText
    End With

    ThisWorkbook.VBProject.VBComponents.Remove tempComp
    ReplaceExistingComponentCodeFromFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not tempComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove tempComp
    End If
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    ReplaceExistingComponentCodeFromFile = False
End Function

Private Sub ImportDocumentModule(modName As String, filePath As String)
    ReplaceExistingComponentCodeFromFile modName, filePath
End Sub

Private Function ImportNewComponentFromFile(filePath As String) As Boolean
    Dim backupPath As String
    backupPath = filePath & "_"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler

    fso.CopyFile filePath, backupPath
    ConvertEncoding filePath, "UTF-8", GetSystemAnsiCharset()
    ThisWorkbook.VBProject.VBComponents.Import filePath
    fso.DeleteFile filePath, True
    fso.MoveFile backupPath, filePath
    ImportNewComponentFromFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If fso.FileExists(backupPath) Then
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        fso.MoveFile backupPath, filePath
    End If
    ImportNewComponentFromFile = False
End Function

Private Function IsFormCurrentlyLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If LCase(frm.Name) = LCase(formName) Then IsFormCurrentlyLoaded = True: Exit Function
    Next frm
End Function

Private Function ComponentExists(modName As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(modName)
    ComponentExists = Not comp Is Nothing
    On Error GoTo 0
End Function


Private Function BuildUniqueTempModuleName(baseName As String) As String
    Dim candidate As String
    candidate = baseName & "_tmp_"
    
    If Not ComponentExists(candidate) Then
        BuildUniqueTempModuleName = candidate
        Exit Function
    End If
    
    Dim idx As Long
    idx = 1
    Do
        candidate = baseName & "_tmp_" & CStr(idx)
        If Not ComponentExists(candidate) Then
            BuildUniqueTempModuleName = candidate
            Exit Function
        End If
        idx = idx + 1
    Loop
End Function

Private Function StripAttributeLines(ByVal codeText As String) As String
    Dim normalized As String
    normalized = Replace(codeText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim lines() As String
    lines = Split(normalized, vbLf)

    Dim outLines() As String
    ReDim outLines(UBound(lines))
    Dim outCount As Long

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Left$(LCase$(Trim$(lines(i))), 13) <> "attribute vb_" Then
            outLines(outCount) = lines(i)
            outCount = outCount + 1
        End If
    Next i

    If outCount = 0 Then Exit Function
    ReDim Preserve outLines(outCount - 1)
    StripAttributeLines = Join(outLines, vbCrLf)
End Function

Private Sub RenameModuleInFile(filePath As String, oldName As String, newName As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    
    ' Read as UTF-8 (handles BOM if present)
    st.Open: st.Type = 2: st.Charset = "UTF-8"
    st.LoadFromFile filePath
    Dim content As String
    content = st.ReadText
    st.Close
    
    content = Replace(content, "VB_Name = """ & oldName & """", "VB_Name = """ & newName & """")
    
    ' Write back as UTF-8
    st.Open: st.Type = 2: st.Charset = "UTF-8"
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
End Sub

Private Function ImportForm(formName As String, frmPath As String) As Boolean
    If ComponentExists(formName) Then
        ImportForm = ReplaceExistingFormCodeFromFile(formName, frmPath)
    Else
        ImportForm = ImportNewComponentFromFile(frmPath)
    End If
End Function

Private Function ReplaceExistingFormCodeFromFile(formName As String, frmPath As String) As Boolean
    Dim tempName As String
    tempName = BuildUniqueTempModuleName(formName)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim srcDir As String
    srcDir = fso.GetParentFolderName(frmPath)

    Dim tempFrmPath As String
    tempFrmPath = srcDir & "\" & tempName & ".frm"
    Dim frxPath As String
    frxPath = srcDir & "\" & formName & ".frx"
    Dim tempFrxPath As String
    tempFrxPath = srcDir & "\" & tempName & ".frx"

    Dim tempComp As Object
    Dim codeText As String

    On Error GoTo ErrHandler

    fso.CopyFile frmPath, tempFrmPath
    RenameModuleInFile tempFrmPath, formName, tempName
    RenameFormBeginInFile tempFrmPath, formName, tempName
    If fso.FileExists(frxPath) Then fso.CopyFile frxPath, tempFrxPath

    ConvertEncoding tempFrmPath, "UTF-8", GetSystemAnsiCharset()
    ThisWorkbook.VBProject.VBComponents.Import tempFrmPath

    If fso.FileExists(tempFrmPath) Then fso.DeleteFile tempFrmPath, True
    If fso.FileExists(tempFrxPath) Then fso.DeleteFile tempFrxPath, True

    Set tempComp = ThisWorkbook.VBProject.VBComponents(tempName)
    If tempComp.CodeModule.CountOfLines > 0 Then
        codeText = tempComp.CodeModule.lines(1, tempComp.CodeModule.CountOfLines)
        codeText = StripAttributeLines(codeText)
    End If

    With ThisWorkbook.VBProject.VBComponents(formName).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(codeText) > 0 Then .AddFromString codeText
    End With

    ThisWorkbook.VBProject.VBComponents.Remove tempComp
    ReplaceExistingFormCodeFromFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not tempComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove tempComp
    End If
    If fso.FileExists(tempFrmPath) Then fso.DeleteFile tempFrmPath, True
    If fso.FileExists(tempFrxPath) Then fso.DeleteFile tempFrxPath, True
    ReplaceExistingFormCodeFromFile = False
End Function

Private Sub RenameFormBeginInFile(filePath As String, oldName As String, newName As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Open: st.Type = 2: st.Charset = "UTF-8"
    st.LoadFromFile filePath
    Dim content As String
    content = st.ReadText
    st.Close

    ' The Begin block header line ends "} FormName " (VBE always appends a trailing space).
    content = Replace(content, "} " & oldName & " ", "} " & newName & " ")

    st.Open: st.Type = 2: st.Charset = "UTF-8"
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
End Sub

Sub ExportAllSheetMapsToMD(Optional skipConfirm As Boolean = False)
    If Not skipConfirm Then
        If MsgBox(t("msg.export_sheets_confirm", "Export all sheet maps to sheet/?"), vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim oldStatusBar As Variant
    Dim oldDisplayStatusBar As Boolean
    oldStatusBar = Application.StatusBar
    oldDisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Dim ws As Worksheet
    Dim sheetFolderPath As String
    Dim fileName As String
    Dim mdContent As String
    Dim sheetIndex As Long
    Dim sheetCount As Long

    On Error GoTo lblErr
    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sheetFolderPath) Then
        fso.CreateFolder sheetFolderPath
    End If

    sheetCount = ThisWorkbook.Worksheets.Count
    For Each ws In ThisWorkbook.Worksheets
        sheetIndex = sheetIndex + 1
        Application.StatusBar = Fmt(t("status.export_sheet_map", "Exporting sheet map {0}/{1}: {2}"), _
                                    sheetIndex, sheetCount, ws.Name)
        mdContent = GenerateSheetMapMarkdown(ws)
        
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        
        SaveAsUTF8 fileName, mdContent
    Next
    
    MsgBox Fmt(t("msg.sheet_maps_exported", "All sheet maps exported. Saved to: {0}"), sheetFolderPath), vbInformation
    GoTo lblFin

lblErr:
    MsgBox Fmt(t("msg.sheet_map_export_error", _
                  "Error while exporting sheet maps: {0}"), Err.Description), vbCritical

lblFin:
    On Error Resume Next
    Application.StatusBar = oldStatusBar
    Application.DisplayStatusBar = oldDisplayStatusBar
    On Error GoTo 0
End Sub

Sub ImportAllSheetMapsFromMD(Optional skipConfirm As Boolean = False)
    If Not skipConfirm Then
        If MsgBox(t("msg.import_sheets_confirm", "Import all sheet maps from sheet/?"), vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim preScreenUpdate As Boolean
    preScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Dim preCalcMode As XlCalculation
    preCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Dim preEvents As Boolean
    preEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Dim sheetFolderPath As String
    Dim fileName As String
    Dim mdContent As String

    On Error GoTo lblErr
    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sheetFolderPath) Then
        MsgBox Fmt(t("msg.sheet_not_found", "Sheet folder not found: {0}"), sheetFolderPath), vbExclamation
        GoTo lblFin
    End If
    StartImportDiagnosticLog sheetFolderPath
    LogImportDiagnostic "ImportAllSheetMapsFromMD start workbook=" & ThisWorkbook.FullName
    If AbortIfProtectedImportSheets(sheetFolderPath, fso, "ImportAllSheetMapsFromMD") Then
        GoTo lblFin
    End If

    For Each ws In ThisWorkbook.Worksheets
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        If fso.FileExists(fileName) Then
            LogImportDiagnostic "Sheet import start codeName=" & ws.codeName & _
                                " name=" & ws.Name & _
                                " file=" & fileName & _
                                " bytes=" & fso.GetFile(fileName).Size
            mdContent = ReadUTF8(fileName)
            ApplySheetMapMarkdown ws, mdContent
            LogImportDiagnostic "Sheet import done codeName=" & ws.codeName & _
                                " name=" & ws.Name
        Else
            LogImportDiagnostic "Sheet import skipped missing file codeName=" & ws.codeName & _
                                " name=" & ws.Name & _
                                " file=" & fileName
        End If
    Next ws

    LogImportDiagnostic "ImportAllSheetMapsFromMD done"
    MsgBox t("msg.sheet_maps_imported", "All sheet maps imported."), vbInformation
    GoTo lblFin
lblErr:
    LogImportDiagnostic "ImportAllSheetMapsFromMD ERROR " & ErrText()
    MsgBox t("msg.sheet_map_import_error", _
    "Error importing sheet maps. Some sheets may be partially updated.") _
    & vbCrLf & Err.Description, vbExclamation
lblFin:
    Application.EnableEvents = preEvents
    Application.Calculation = preCalcMode
    Application.ScreenUpdating = preScreenUpdate
End Sub

Private Sub StartImportDiagnosticLog(sheetFolderPath As String)
    If Not IMPORT_DIAGNOSTICS_ENABLED Then Exit Sub

    devkitImportLogPath = ThisWorkbook.Path & "\devkit_import_diagnostics.log"

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(devkitImportLogPath, True)
    ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & _
                 vbTab & "Import diagnostics start sheetFolder=" & sheetFolderPath
    ts.Close
    On Error GoTo 0
End Sub

Private Sub LogImportDiagnostic(message As String)
    If Not IMPORT_DIAGNOSTICS_ENABLED Then Exit Sub
    If devkitImportLogPath = "" Then Exit Sub

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(devkitImportLogPath, 8, True)
    ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & message
    ts.Close
    On Error GoTo 0
End Sub

Private Function ErrText() As String
    ErrText = "Err " & Err.Number & ": " & Err.Description
End Function

Private Function ShortLogText(text As String, Optional maxLen As Long = 240) As String
    Dim s As String
    s = Replace(text, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, Chr(11), "\v")
    If Len(s) > maxLen Then s = Left(s, maxLen) & "..."
    ShortLogText = s
End Function

Private Function GetProtectedImportSheetList(sheetFolderPath As String, fso As Object) As String
    Dim ws As Worksheet
    Dim fileName As String
    Dim result As String

    result = ""
    For Each ws In ThisWorkbook.Worksheets
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        If fso.FileExists(fileName) Then
            If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
                If result <> "" Then result = result & vbCrLf
                result = result & "- " & ws.Name & " (" & ws.codeName & ")"
            End If
        End If
    Next ws

    GetProtectedImportSheetList = result
End Function

Private Function AbortIfProtectedImportSheets(sheetFolderPath As String, fso As Object, _
                                              Optional logPrefix As String = "") As Boolean
    Dim protectedSheets As String

    protectedSheets = GetProtectedImportSheetList(sheetFolderPath, fso)
    If protectedSheets <> "" Then
        If logPrefix <> "" Then
            LogImportDiagnostic logPrefix & " canceled protectedSheets=" & _
                                Replace(protectedSheets, vbCrLf, " | ")
        End If
        MsgBox t("msg.import_sheets_protected", _
        "Import canceled. Unprotect these sheets before importing:") & _
        vbCrLf & vbCrLf & protectedSheets, vbExclamation
        AbortIfProtectedImportSheets = True
    End If
End Function

Private Sub StringBuilderInit(ByRef sb As DevkitStringBuilder)
    sb.Capacity = STRING_BUILDER_CHUNK_SIZE
    sb.Count = 0
    ReDim sb.Parts(0 To sb.Capacity - 1)
End Sub

Private Sub StringBuilderAppend(ByRef sb As DevkitStringBuilder, ByVal text As String)
    If sb.Capacity = 0 Then StringBuilderInit sb

    If sb.Count >= sb.Capacity Then
        sb.Capacity = sb.Capacity + STRING_BUILDER_CHUNK_SIZE
        ReDim Preserve sb.Parts(0 To sb.Capacity - 1)
    End If

    sb.Parts(sb.Count) = text
    sb.Count = sb.Count + 1
End Sub

Private Function StringBuilderToString(ByRef sb As DevkitStringBuilder) As String
    If sb.Count = 0 Then
        StringBuilderToString = ""
        Exit Function
    End If

    If sb.Count < sb.Capacity Then
        ReDim Preserve sb.Parts(0 To sb.Count - 1)
        sb.Capacity = sb.Count
    End If

    StringBuilderToString = Join(sb.Parts, "")
End Function

Private Function GenerateSheetMapMarkdown(ws As Worksheet) As String
    Dim rng As Range
    Dim mapText As DevkitStringBuilder
    Dim role As String
    Dim cellName As String
    Dim styleParts As String
    Dim mergeMarker As String
    Dim cellValue As Variant
    Dim cellValueText As String
    Dim cellFormulaText As String
    Dim shouldExportCell As Boolean

    StringBuilderInit mapText
    StringBuilderAppend mapText, "# Sheet Configuration" & vbCrLf
    StringBuilderAppend mapText, "- VBA CodeName: " & ws.codeName & vbCrLf
    StringBuilderAppend mapText, "- Excel UI Name: " & ws.Name & vbCrLf
    Dim hiddenRowStr As String
    Dim hiddenColStr As String
    hiddenRowStr = CollectHiddenRowRanges(ws)
    hiddenColStr = CollectHiddenColRanges(ws)
    If Len(hiddenRowStr) > 0 Then
        StringBuilderAppend mapText, "- Hidden Rows: " & hiddenRowStr & vbCrLf
    End If
    If Len(hiddenColStr) > 0 Then
        StringBuilderAppend mapText, "- Hidden Columns: " & hiddenColStr & vbCrLf
    End If
    StringBuilderAppend mapText, vbCrLf

    StringBuilderAppend mapText, "| Address | Name | Value / Label | Formula | Style |" & vbCrLf
    StringBuilderAppend mapText, "| :--- | :--- | :--- | :--- | :--- |" & vbCrLf
    
    For Each rng In ws.UsedRange
        cellValue = rng.Value
        If IsError(cellValue) Then
            cellValueText = rng.Text
            shouldExportCell = (cellValueText <> "" _
            Or rng.HasFormula Or rng.Interior.ColorIndex <> xlNone)
        Else
            cellValueText = CStr(cellValue)
            shouldExportCell = (cellValueText <> "" _
            Or rng.HasFormula Or rng.Interior.ColorIndex <> xlNone)
        End If

        If rng.MergeCells And Not (rng.Row = rng.MergeArea.Row _
        And rng.Column = rng.MergeArea.Column) Then
            ' Slave cell in a merged range
            cellName = ""
            On Error Resume Next
            cellName = rng.Name.Name
            If cellName = "" Then cellName = "-"
            On Error GoTo 0

            If rng.Row = rng.MergeArea.Row Then
                mergeMarker = "!merged_left"
            ElseIf rng.Column = rng.MergeArea.Column Then
                mergeMarker = "!merged_up"
            Else
                mergeMarker = "!merged_ul"
            End If

            styleParts = BuildCellStyle(rng)

            StringBuilderAppend mapText, "| " & rng.Address(False, False) & _
                                 " | " & cellName & _
                                 " | " & mergeMarker & _
                                 " | -" & _
                                 " | " & IIf(styleParts = "", "-", styleParts) & " |" & vbCrLf
        ElseIf shouldExportCell Then
            ' Normal cell (or master cell of a merged range)
            cellName = ""
            On Error Resume Next
            cellName = rng.Name.Name
            If cellName = "" Then cellName = "-"
            On Error GoTo 0

            styleParts = BuildCellStyle(rng)
            role = IIf(styleParts = "", "-", styleParts)

            If rng.HasFormula Then
                cellFormulaText = "`" & EscapeCellValue(CStr(rng.Formula)) & "`"
            Else
                cellFormulaText = "-"
            End If

            StringBuilderAppend mapText, "| " & rng.Address(False, False) & _
                                 " | " & cellName & _
                                 " | " & EscapeCellValue(cellValueText) & _
                                 " | " & cellFormulaText & _
                                 " | " & role & " |" & vbCrLf
        End If
    Next
    
    Dim shp As Object
    Dim shpLabel As String
    Dim shpOnAction As String
    Dim shpStyle As String
    Dim shpFillRGB As Long
    Dim shpFontRGB As Long
    Dim shpFSize As Variant
    Dim shpFormula As String
    Dim shpFml As String
    Dim shapeRows As DevkitStringBuilder

    StringBuilderInit shapeRows

    For Each shp In ws.Shapes
        On Error GoTo lblFinShp

        shpLabel = ""
        shpOnAction = ""
        shpStyle = ""
        shpFormula = "-"
        shpFml = ""

        On Error Resume Next
        shpLabel = shp.TextFrame.Characters.text
        shpOnAction = shp.OnAction
        On Error GoTo 0

        On Error Resume Next
        shpFillRGB = shp.Fill.ForeColor.RGB
        If Err.Number = 0 Then shpStyle = "BG:" & ColorToHex(shpFillRGB)
        Err.Clear
        shpFontRGB = shp.TextFrame.Characters.Font.Color
        If Err.Number = 0 And shpFontRGB <> 0 Then
            If shpStyle <> "" Then shpStyle = shpStyle & "; "
            shpStyle = shpStyle & "FG:" & ColorToHex(shpFontRGB)
        End If
        Err.Clear
        shpFSize = shp.TextFrame.Characters.Font.Size
        If Err.Number = 0 And Not IsNull(shpFSize) Then
            If shpFSize <> Application.StandardFontSize Then
                If shpStyle <> "" Then shpStyle = shpStyle & "; "
                shpStyle = shpStyle & "FontSize:" & shpFSize
            End If
        End If

        Err.Clear
        shpFml = Trim(shp.DrawingObject.Formula)
        If Err.Number = 0 And shpFml <> "" Then
            If Left(shpFml, 1) <> "=" Then shpFml = "=" & shpFml
            shpFormula = "`" & EscapeCellValue(shpFml) & "`"
        End If
        On Error GoTo 0

        If shpLabel <> "" Or shpFormula <> "-" Or shpOnAction <> "" Then
            StringBuilderAppend shapeRows, "| " & shp.TopLeftCell.Address(False, False) & _
                                           " | " & EscapeCellValue(shp.Name) & _
                                           " | " & IIf(shpLabel <> "", IIf(shpLabel = "-", "\-", EscapeCellValue(shpLabel)), "-") & _
                                           " | " & shpFormula & _
                                           " | " & IIf(shpOnAction <> "", IIf(shpOnAction = "-", "\-", EscapeCellValue(shpOnAction)), "-") & _
                                           " | " & IIf(shpStyle <> "", shpStyle, "-") & " |" & vbCrLf
        End If
lblFinShp:
        Err.Clear
    Next shp

    If shapeRows.Count > 0 Then
        StringBuilderAppend mapText, vbCrLf & "## Shapes" & vbCrLf & vbCrLf
        StringBuilderAppend mapText, "| Address | Name | Label | Formula | OnAction | Style |" & vbCrLf
        StringBuilderAppend mapText, "| :--- | :--- | :--- | :--- | :--- | :--- |" & vbCrLf
        StringBuilderAppend mapText, StringBuilderToString(shapeRows)
    End If

    GenerateSheetMapMarkdown = StringBuilderToString(mapText)
End Function

Private Sub ApplySheetMapMarkdown(ws As Worksheet, mdContent As String)
    Dim normContent As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim inCellTable As Boolean
    Dim cols() As String
    Dim addr As String
    Dim cName As String
    Dim cValue As String
    Dim cFormula As String
    Dim cStyle As String
    Dim rng As Range
    Dim fml As String
    Dim unescapedValue As String
    Dim refAddr As String
    Dim mAddr As String
    Dim slaveAddrs() As String
    Dim masterAddrs() As String
    Dim slaveCount As Long
    Dim hiddenRowsStr As String
    Dim hiddenColsStr As String
    Dim parsedRows As Long
    Dim malformedRows As Long
    Dim invalidAddressRows As Long
    Dim formulaApplied As Long
    Dim valueApplied As Long
    Dim blankOrStyleRows As Long
    Dim mergeRows As Long
    hiddenRowsStr = ""
    hiddenColsStr = ""

    ' Pre-clear cell data, formatting, merges, and validation
    On Error Resume Next
    Err.Clear
    ws.Cells.UnMerge
    If Err.Number <> 0 Then LogImportDiagnostic "WARN sheet=" & ws.codeName & " unmerge failed " & ErrText()
    Err.Clear
    ws.Cells.ClearContents
    If Err.Number <> 0 Then LogImportDiagnostic "ERROR sheet=" & ws.codeName & " ClearContents failed " & ErrText()
    Err.Clear
    ws.Cells.ClearFormats
    If Err.Number <> 0 Then LogImportDiagnostic "ERROR sheet=" & ws.codeName & " ClearFormats failed " & ErrText()
    Err.Clear
    ws.Cells.Validation.Delete
    If Err.Number <> 0 Then LogImportDiagnostic "WARN sheet=" & ws.codeName & " Validation.Delete failed " & ErrText()
    On Error GoTo 0

    ' Normalize line endings and split
    normContent = Replace(mdContent, vbCrLf, vbLf)
    normContent = Replace(normContent, vbCr, vbLf)
    lines = Split(normContent, vbLf)
    LogImportDiagnostic "ApplySheetMapMarkdown start sheet=" & ws.codeName & _
                        " name=" & ws.Name & _
                        " lineCount=" & (UBound(lines) + 1)

    inCellTable = False
    slaveCount = 0
    ReDim slaveAddrs(UBound(lines))
    ReDim masterAddrs(UBound(lines))

    For i = 0 To UBound(lines)
        line = Trim(lines(i))

        If line = "## Shapes" Then
            LogImportDiagnostic "Sheet cell table ended at Shapes sheet=" & ws.codeName & _
                                " mdLine=" & (i + 1)
            Exit For
        End If

        If Left(line, 10) = "| Address " Then
            inCellTable = True
        ElseIf Left(line, 15) = "- Hidden Rows: " Then
            hiddenRowsStr = Trim(Mid(line, 16))
        ElseIf Left(line, 18) = "- Hidden Columns: " Then
            hiddenColsStr = Trim(Mid(line, 19))
        ElseIf inCellTable And Left(line, 1) = "|" And Left(line, 3) <> "| :" Then
            cols = ParseMDTableRow(line)
            If UBound(cols) >= 4 Then
                addr = Trim(cols(0))
                cName = Trim(cols(1))
                cValue = cols(2)
                cFormula = cols(3)
                cStyle = Trim(cols(4))
                parsedRows = parsedRows + 1

                Set rng = Nothing
                On Error Resume Next
                Err.Clear
                Set rng = ws.Range(addr)
                If Err.Number <> 0 Then
                    LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                        " mdLine=" & (i + 1) & _
                                        " invalid address=" & ShortLogText(addr) & _
                                        " " & ErrText()
                    invalidAddressRows = invalidAddressRows + 1
                    Err.Clear
                End If
                On Error GoTo 0

                If Not rng Is Nothing Then
                    If cValue = "!merged_left" Or cValue = "!merged_up" Or cValue = "!merged_ul" Then
                        ' Slave cell: apply style and name, record for Pass 2
                        ApplyCellStyle rng, cStyle
                        If cName <> "-" Then ApplyCellName ws, rng, cName
                        mergeRows = mergeRows + 1

                        If cValue = "!merged_left" Then
                            refAddr = ws.Cells(rng.Row, rng.Column - 1).Address(False, False)
                        Else
                            ' !merged_up or !merged_ul: follow upward to find master
                            refAddr = ws.Cells(rng.Row - 1, rng.Column).Address(False, False)
                        End If
                        mAddr = ResolveMasterAddr(refAddr, slaveAddrs, masterAddrs, slaveCount)

                        slaveAddrs(slaveCount) = addr
                        masterAddrs(slaveCount) = mAddr
                        slaveCount = slaveCount + 1
                    Else
                        ' Normal or master cell
                        If cFormula <> "-" And cFormula <> "" Then
                            fml = cFormula
                            If Left(fml, 1) = "`" And Right(fml, 1) = "`" Then
                                fml = Mid(fml, 2, Len(fml) - 2)
                            End If
                            fml = UnescapeCellValue(fml)
                            On Error Resume Next
                            Err.Clear
                            rng.Formula = fml
                            If Err.Number <> 0 Then
                                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                                    " mdLine=" & (i + 1) & _
                                                    " addr=" & addr & _
                                                    " set formula failed " & ErrText() & _
                                                    " formula=" & ShortLogText(fml)
                                Err.Clear
                            Else
                                formulaApplied = formulaApplied + 1
                            End If
                            On Error GoTo 0
                        ElseIf cValue <> "-" And cValue <> "" Then
                            On Error Resume Next
                            Err.Clear
                            unescapedValue = UnescapeCellValue(cValue)
                            rng.Value = unescapedValue
                            If Err.Number <> 0 Then
                                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                                    " mdLine=" & (i + 1) & _
                                                    " addr=" & addr & _
                                                    " set value failed " & ErrText() & _
                                                    " value=" & ShortLogText(unescapedValue)
                                Err.Clear
                            Else
                                valueApplied = valueApplied + 1
                            End If
                            On Error GoTo 0
                        Else
                            blankOrStyleRows = blankOrStyleRows + 1
                        End If
                        ApplyCellStyle rng, cStyle
                        If cName <> "-" Then ApplyCellName ws, rng, cName
                    End If
                ElseIf Err.Number = 0 Then
                    LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                        " mdLine=" & (i + 1) & _
                                        " address resolved to Nothing addr=" & ShortLogText(addr)
                    invalidAddressRows = invalidAddressRows + 1
                End If
            Else
                malformedRows = malformedRows + 1
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " mdLine=" & (i + 1) & _
                                    " malformed row cols=" & (UBound(cols) + 1) & _
                                    " line=" & ShortLogText(line)
            End If
        End If
    Next i

    ' Pass 2: reconstruct merged ranges
    If slaveCount > 0 Then ReconstructMerges ws, slaveAddrs, masterAddrs, slaveCount

    If Len(hiddenRowsStr) > 0 Then ApplyHiddenRows ws, hiddenRowsStr
    If Len(hiddenColsStr) > 0 Then ApplyHiddenCols ws, hiddenColsStr
    ApplyShapeMapMarkdown ws, lines
    LogImportDiagnostic "ApplySheetMapMarkdown summary sheet=" & ws.codeName & _
                        " parsed=" & parsedRows & _
                        " formulas=" & formulaApplied & _
                        " values=" & valueApplied & _
                        " blankOrStyle=" & blankOrStyleRows & _
                        " mergeRows=" & mergeRows & _
                        " malformed=" & malformedRows & _
                        " invalidAddress=" & invalidAddressRows & _
                        " hiddenRows='" & hiddenRowsStr & "'" & _
                        " hiddenCols='" & hiddenColsStr & "'"
End Sub

Private Sub ApplyShapeMapMarkdown(ws As Worksheet, lines() As String)
    Dim i As Long
    Dim line As String
    Dim inShapeTable As Boolean
    Dim cols() As String
    Dim addr As String
    Dim shpName As String
    Dim shpLabel As String
    Dim shpFormula As String
    Dim shpOnAction As String
    Dim shpStyle As String
    Dim shp As Object
    Dim anchor As Range
    Dim updatedCount As Long
    Dim createdCount As Long
    Dim malformedCount As Long
    Dim invalidAddressCount As Long

    inShapeTable = False

    For i = 0 To UBound(lines)
        line = Trim(lines(i))

        If line = "## Shapes" Then
            inShapeTable = True
        ElseIf inShapeTable And Left(line, 1) = "|" And Left(line, 3) <> "| :" Then
            cols = ParseMDTableRow(line)
            If UBound(cols) >= 5 Then
                addr = Trim(cols(0))
                If addr = "Address" Then GoTo lblNextShapeLine

                shpName = UnescapeCellValue(cols(1))
                shpLabel = cols(2)
                shpFormula = cols(3)
                shpOnAction = cols(4)
                shpStyle = Trim(cols(5))

                Set anchor = Nothing
                On Error Resume Next
                Err.Clear
                Set anchor = ws.Range(addr)
                If Err.Number <> 0 Then
                    LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                        " mdLine=" & (i + 1) & _
                                        " shape invalid anchor=" & ShortLogText(addr) & _
                                        " " & ErrText()
                    invalidAddressCount = invalidAddressCount + 1
                    Err.Clear
                End If
                On Error GoTo 0
                If anchor Is Nothing Then GoTo lblNextShapeLine

                Set shp = FindShapeByName(ws, shpName)
                If shp Is Nothing Then
                    Set shp = CreateImportedShape(ws, anchor, shpName)
                    If shp Is Nothing Then
                        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                            " mdLine=" & (i + 1) & _
                                            " shape create failed name=" & ShortLogText(shpName) & _
                                            " anchor=" & addr
                        GoTo lblNextShapeLine
                    End If
                    createdCount = createdCount + 1
                Else
                    updatedCount = updatedCount + 1
                End If

                ApplyShapeFields ws, shp, i + 1, shpLabel, shpFormula, shpOnAction, shpStyle
            Else
                malformedCount = malformedCount + 1
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " mdLine=" & (i + 1) & _
                                    " malformed shape row cols=" & (UBound(cols) + 1) & _
                                    " line=" & ShortLogText(line)
            End If
        End If
lblNextShapeLine:
    Next i

    LogImportDiagnostic "ApplyShapeMapMarkdown summary sheet=" & ws.codeName & _
                        " updated=" & updatedCount & _
                        " created=" & createdCount & _
                        " malformed=" & malformedCount & _
                        " invalidAddress=" & invalidAddressCount
End Sub

Private Function FindShapeByName(ws As Worksheet, shapeName As String) As Object
    On Error Resume Next
    Set FindShapeByName = ws.Shapes(shapeName)
    If Err.Number <> 0 Then
        Err.Clear
        Set FindShapeByName = Nothing
    End If
    On Error GoTo 0
End Function

Private Function CreateImportedShape(ws As Worksheet, anchor As Range, shapeName As String) As Object
    On Error Resume Next
    Dim shp As Object
    Set shp = ws.Shapes.AddShape(GuessShapeType(shapeName), anchor.Left, anchor.Top, 120, 30)
    If Err.Number <> 0 Then
        Err.Clear
        Set shp = Nothing
    End If
    If Not shp Is Nothing Then
        Err.Clear
        shp.Name = shapeName
        Err.Clear
    End If
    Set CreateImportedShape = shp
    On Error GoTo 0
End Function

Private Function GuessShapeType(shapeName As String) As Long
    Dim lowerName As String
    Dim calloutPrefix As String

    lowerName = LCase(shapeName)
    calloutPrefix = "rounded rectangular callout"

    If Left(lowerName, Len(calloutPrefix)) = calloutPrefix Then
        GuessShapeType = msoShapeRoundedRectangularCallout
    ElseIf Left(lowerName, 6) = "bevel " Then
        GuessShapeType = msoShapeBevel
    ElseIf lowerName = "bevel" Then
        GuessShapeType = msoShapeBevel
    Else
        GuessShapeType = msoShapeRectangle
    End If
End Function

Private Sub ApplyShapeFields(ws As Worksheet, shp As Object, mdLine As Long, _
                             shpLabel As String, shpFormula As String, _
                             shpOnAction As String, shpStyle As String)
    Dim fml As String

    On Error Resume Next
    Err.Clear
    If shpLabel = "-" Then
        shp.TextFrame.Characters.text = ""
    Else
        shp.TextFrame.Characters.text = UnescapeCellValue(shpLabel)
    End If
    If Err.Number <> 0 Then
        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                            " mdLine=" & mdLine & _
                            " shape label failed name=" & ShortLogText(shp.Name) & _
                            " " & ErrText()
        Err.Clear
    End If

    If shpFormula <> "-" And shpFormula <> "" Then
        fml = shpFormula
        If Left(fml, 1) = "`" And Right(fml, 1) = "`" Then
            fml = Mid(fml, 2, Len(fml) - 2)
        End If
        fml = UnescapeCellValue(fml)
        Err.Clear
        shp.DrawingObject.Formula = fml
        If Err.Number <> 0 Then
            LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                " mdLine=" & mdLine & _
                                " shape formula failed name=" & ShortLogText(shp.Name) & _
                                " " & ErrText() & _
                                " formula=" & ShortLogText(fml)
            Err.Clear
        End If
    End If

    Err.Clear
    If shpOnAction = "-" Then
        shp.OnAction = ""
    Else
        shp.OnAction = UnescapeCellValue(shpOnAction)
    End If
    If Err.Number <> 0 Then
        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                            " mdLine=" & mdLine & _
                            " shape OnAction failed name=" & ShortLogText(shp.Name) & _
                            " " & ErrText()
        Err.Clear
    End If
    On Error GoTo 0

    ApplyShapeStyle ws, shp, mdLine, shpStyle
End Sub

Private Sub ApplyShapeStyle(ws As Worksheet, shp As Object, mdLine As Long, styleStr As String)
    If styleStr = "-" Or styleStr = "" Then Exit Sub

    Dim parts() As String
    Dim p As String
    Dim tokenName As String
    Dim tokenValue As String
    Dim sep As Long
    Dim k As Long
    parts = ParseStyleTokens(styleStr)

    For k = 0 To UBound(parts)
        p = Trim(parts(k))
        If p = "" Then GoTo lblNextShapeStyle
        sep = InStr(p, ":")
        If sep > 0 Then
            tokenName = UCase(Trim(Left(p, sep - 1)))
            tokenValue = UnescapeStyleValue(Mid(p, sep + 1))
        Else
            tokenName = UCase(p)
            tokenValue = ""
        End If

        If tokenName = "BG" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            shp.Fill.ForeColor.RGB = HexToRGB(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " mdLine=" & mdLine & _
                                    " shape BG failed name=" & ShortLogText(shp.Name) & _
                                    " " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "FG" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            shp.TextFrame.Characters.Font.Color = HexToRGB(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " mdLine=" & mdLine & _
                                    " shape FG failed name=" & ShortLogText(shp.Name) & _
                                    " " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "FONTSIZE" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            shp.TextFrame.Characters.Font.Size = CDbl(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " mdLine=" & mdLine & _
                                    " shape FontSize failed name=" & ShortLogText(shp.Name) & _
                                    " " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf p <> "" Then
            LogImportDiagnostic "WARN sheet=" & ws.codeName & _
                                " mdLine=" & mdLine & _
                                " shape unknown style token name=" & ShortLogText(shp.Name) & _
                                " token=" & ShortLogText(p)
        End If
lblNextShapeStyle:
    Next k
End Sub

Private Function EscapeCellValue(cellValue As String) As String
    Dim v As String
    v = cellValue
    v = Replace(v, "\", "\\")
    v = Replace(v, vbCrLf, "\n")
    v = Replace(v, vbCr, "\n")
    v = Replace(v, vbLf, "\n")
    v = Replace(v, Chr(11), "\v")
    v = Replace(v, "|", "\|")
    EscapeCellValue = v
End Function

Private Function ColorToHex(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And 255
    g = (colorVal \ 256) And 255
    b = (colorVal \ 65536) And 255
    ColorToHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

Private Function BuildCellStyle(rng As Range) As String
    Dim styleParts As String
    Dim valFormula As String
    Dim valType As Long
    Dim numberFormat As String

    styleParts = ""

    On Error Resume Next
    If rng.Interior.ColorIndex <> xlNone Then
        AppendStyleKey styleParts, "BG", ColorToHex(rng.Interior.Color)
    End If
    If rng.Font.ColorIndex <> xlColorIndexAutomatic And rng.Font.Color <> 0 Then
        AppendStyleKey styleParts, "FG", ColorToHex(rng.Font.Color)
    End If
    If rng.Font.Size <> Application.StandardFontSize Then
        AppendStyleKey styleParts, "FontSize", CStr(rng.Font.Size)
    End If
    If rng.Font.Bold = True Then AppendStyleFlag styleParts, "Bold"
    If rng.Font.Italic = True Then AppendStyleFlag styleParts, "Italic"
    If rng.Font.Strikethrough = True Then AppendStyleFlag styleParts, "Strike"
    If rng.WrapText = True Then AppendStyleFlag styleParts, "Wrap"
    If rng.Locked = False Then AppendStyleFlag styleParts, "Unlocked"

    Err.Clear
    numberFormat = CStr(rng.NumberFormat)
    If Err.Number = 0 Then
        If numberFormat <> "" And numberFormat <> "General" Then
            AppendStyleKey styleParts, "NumFmt", numberFormat
        End If
    End If

    Err.Clear
    valFormula = ""
    valType = rng.Validation.Type
    If Err.Number = 0 And valType = xlValidateList Then
        valFormula = rng.Validation.Formula1
    End If
    If valFormula <> "" Then AppendStyleKey styleParts, "List", valFormula
    On Error GoTo 0

    BuildCellStyle = styleParts
End Function

Private Sub AppendStyleFlag(ByRef styleParts As String, flagName As String)
    If styleParts <> "" Then styleParts = styleParts & "; "
    styleParts = styleParts & flagName
End Sub

Private Sub AppendStyleKey(ByRef styleParts As String, keyName As String, valueText As String)
    If styleParts <> "" Then styleParts = styleParts & "; "
    styleParts = styleParts & keyName & ":" & EscapeStyleValue(valueText)
End Sub

Private Function NormalizeStartUpPosition(content As String) As String
    Dim labels(0 To 3) As String
    labels(0) = "0 - Manual"
    labels(1) = "1 - CenterOwner"
    labels(2) = "2 - CenterScreen"
    labels(3) = "3 - Windows Default"

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "StartUpPosition\s*=\s*(\d)\s+'[^\r\n]*"
    re.Global = True

    Dim ms As Object
    Set ms = re.Execute(content)

    Dim i As Long
    For i = ms.count - 1 To 0 Step -1
        Dim m As Object: Set m = ms(i)
        Dim v As Long: v = CLng(m.SubMatches(0))
        Dim canonical As String
        canonical = "StartUpPosition =   " & v & "  '" & labels(v)
        content = Left(content, m.FirstIndex) & canonical & _
                  Mid(content, m.FirstIndex + m.Length + 1)
    Next i

    NormalizeStartUpPosition = content
End Function

Private Sub SaveAsUTF8(filePath As String, content As String)
    ' Normalize trailing blank lines: strip extra CRLFs appended by Excel's VBA exporter,
    ' then ensure exactly one CRLF at end of file.
    Do While Len(content) >= 2 And Right(content, 2) = vbCrLf
        content = Left(content, Len(content) - 2)
    Loop
    Do While Len(content) >= 1 And (Right(content, 1) = vbCr Or Right(content, 1) = vbLf)
        content = Left(content, Len(content) - 1)
    Loop
    content = content & vbCrLf

    ' ADODB.Stream always prepends a 3-byte UTF-8 BOM; strip it via binary copy.
    Dim stText As Object
    Set stText = CreateObject("ADODB.Stream")
    stText.Type = 2
    stText.Charset = "UTF-8"
    stText.Open
    stText.WriteText content
    stText.Position = 0
    stText.Type = 1
    stText.Position = 3    ' skip BOM (EF BB BF)

    Dim stBin As Object
    Set stBin = CreateObject("ADODB.Stream")
    stBin.Type = 1
    stBin.Open
    stBin.Write stText.Read
    stBin.SaveToFile filePath, 2
    stBin.Close
    stText.Close
End Sub

' Returns the ADODB.Stream charset name matching the system ANSI code page.
' VBComponents.Export / Import always use the system ANSI code page,
' so this is used when converting between UTF-8 (on disk) and ANSI (for VBE).
Private Function ReadUTF8(filePath As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "UTF-8"
    st.Open
    st.LoadFromFile filePath
    ReadUTF8 = st.ReadText
    st.Close
End Function

Private Function ParseMDTableRow(line As String) As String()
    ' "| A1 | name | value | formula | style |" -> Array("A1","name","value","formula","style")
    Dim s As String
    s = line
    If Left(s, 1) = "|" Then s = Mid(s, 2)
    If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
    Dim parts() As String
    Dim token As String
    Dim i As Long
    Dim ch As String
    Dim count As Long

    ReDim parts(0)
    token = ""
    count = 0

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "|" And Not IsEscapedMDDelimiter(s, i) Then
            parts(count) = TrimMDTableField(token)
            count = count + 1
            ReDim Preserve parts(count)
            token = ""
        Else
            token = token & ch
        End If
    Next i
    parts(count) = TrimMDTableField(token)
    ParseMDTableRow = parts
End Function

Private Function IsEscapedMDDelimiter(s As String, pos As Long) As Boolean
    Dim slashCount As Long
    Dim i As Long
    slashCount = 0

    i = pos - 1
    Do While i >= 1
        If Mid(s, i, 1) <> "\" Then Exit Do
        slashCount = slashCount + 1
        i = i - 1
    Loop

    IsEscapedMDDelimiter = ((slashCount Mod 2) = 1)
End Function

Private Function TrimMDTableField(v As String) As String
    If Left(v, 1) = " " Then v = Mid(v, 2)
    If Right(v, 1) = " " Then v = Left(v, Len(v) - 1)
    TrimMDTableField = v
End Function

Private Function UnescapeCellValue(v As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String

    result = ""
    i = 1

    Do While i <= Len(v)
        ch = Mid(v, i, 1)
        If ch = "\" And i < Len(v) Then
            nextCh = Mid(v, i + 1, 1)
            Select Case nextCh
                Case "\"
                    result = result & "\"
                Case "n"
                    result = result & vbLf
                Case "v"
                    result = result & Chr(11)
                Case "|"
                    result = result & "|"
                Case "-"
                    result = result & "-"
                Case Else
                    result = result & ch & nextCh
            End Select
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    UnescapeCellValue = result
End Function

Private Function EscapeStyleValue(styleValue As String) As String
    Dim v As String
    v = EscapeCellValue(styleValue)
    v = Replace(v, ";", "\;")
    EscapeStyleValue = v
End Function

Private Function UnescapeStyleValue(v As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String

    result = ""
    i = 1

    Do While i <= Len(v)
        ch = Mid(v, i, 1)
        If ch = "\" And i < Len(v) Then
            nextCh = Mid(v, i + 1, 1)
            Select Case nextCh
                Case "\"
                    result = result & "\"
                Case "n"
                    result = result & vbLf
                Case "v"
                    result = result & Chr(11)
                Case "|"
                    result = result & "|"
                Case "-"
                    result = result & "-"
                Case ";"
                    result = result & ";"
                Case Else
                    result = result & ch & nextCh
            End Select
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    UnescapeStyleValue = result
End Function

Private Function ParseStyleTokens(styleStr As String) As String()
    Dim parts() As String
    Dim token As String
    Dim i As Long
    Dim ch As String
    Dim count As Long

    ReDim parts(0)
    token = ""
    count = 0

    For i = 1 To Len(styleStr)
        ch = Mid(styleStr, i, 1)
        If ch = ";" And Not IsEscapedMDDelimiter(styleStr, i) Then
            parts(count) = Trim(token)
            count = count + 1
            ReDim Preserve parts(count)
            token = ""
            If i < Len(styleStr) Then
                If Mid(styleStr, i + 1, 1) = " " Then i = i + 1
            End If
        Else
            token = token & ch
        End If
    Next i

    parts(count) = Trim(token)
    ParseStyleTokens = parts
End Function

Private Function StyleFlagEnabled(tokenValue As String, hasValue As Boolean) As Boolean
    Dim v As String

    If Not hasValue Then
        StyleFlagEnabled = True
        Exit Function
    End If

    v = LCase(Trim(UnescapeStyleValue(tokenValue)))
    StyleFlagEnabled = (v = "true" Or v = "1" Or v = "yes" Or v = "on")
End Function

Private Function TextToHorizontalAlignment(text As String) As Variant
    Select Case LCase(Trim(text))
        Case "general"
            TextToHorizontalAlignment = xlGeneral
        Case "left"
            TextToHorizontalAlignment = xlLeft
        Case "center"
            TextToHorizontalAlignment = xlCenter
        Case "right"
            TextToHorizontalAlignment = xlRight
        Case "fill"
            TextToHorizontalAlignment = xlFill
        Case "justify"
            TextToHorizontalAlignment = xlJustify
        Case "centeracrossselection"
            TextToHorizontalAlignment = xlCenterAcrossSelection
        Case "distributed"
            TextToHorizontalAlignment = xlDistributed
        Case Else
            TextToHorizontalAlignment = CLng(text)
    End Select
End Function

Private Function TextToVerticalAlignment(text As String) As Variant
    Select Case LCase(Trim(text))
        Case "top"
            TextToVerticalAlignment = xlTop
        Case "center"
            TextToVerticalAlignment = xlCenter
        Case "bottom"
            TextToVerticalAlignment = xlBottom
        Case "justify"
            TextToVerticalAlignment = xlJustify
        Case "distributed"
            TextToVerticalAlignment = xlDistributed
        Case Else
            TextToVerticalAlignment = CLng(text)
    End Select
End Function

Private Function HexToRGB(hexStr As String) As Long
    If Left(hexStr, 1) = "#" Then hexStr = Mid(hexStr, 2)
    HexToRGB = RGB(CLng("&H" & Left(hexStr, 2)), CLng("&H" & Mid(hexStr, 3, 2)), CLng("&H" & Right(hexStr, 2)))
End Function

Private Sub ApplyCellStyle(rng As Range, styleStr As String)
    If styleStr = "-" Or styleStr = "" Then Exit Sub
    Dim parts() As String
    Dim p As String
    Dim tokenName As String
    Dim tokenValue As String
    Dim tokenValueRaw As String
    Dim sep As Long
    Dim k As Long
    parts = ParseStyleTokens(styleStr)

    For k = 0 To UBound(parts)
        p = Trim(parts(k))
        If p = "" Then GoTo lblNextCellStyle
        sep = InStr(p, ":")
        If sep > 0 Then
            tokenName = UCase(Trim(Left(p, sep - 1)))
            tokenValueRaw = Mid(p, sep + 1)
            tokenValue = UnescapeStyleValue(tokenValueRaw)
        Else
            tokenName = UCase(p)
            tokenValueRaw = ""
            tokenValue = ""
        End If

        If tokenName = "BG" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.Interior.Color = HexToRGB(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply BG failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "FG" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.Font.Color = HexToRGB(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply FG failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "FONTSIZE" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.Font.Size = CDbl(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply FontSize failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "NUMFMT" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.NumberFormat = tokenValue
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply NumFmt failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "BOLD" Then
            If StyleFlagEnabled(tokenValueRaw, sep > 0) Then rng.Font.Bold = True
        ElseIf tokenName = "ITALIC" Then
            If StyleFlagEnabled(tokenValueRaw, sep > 0) Then rng.Font.Italic = True
        ElseIf tokenName = "STRIKE" Then
            If StyleFlagEnabled(tokenValueRaw, sep > 0) Then rng.Font.Strikethrough = True
        ElseIf tokenName = "WRAP" Then
            If StyleFlagEnabled(tokenValueRaw, sep > 0) Then rng.WrapText = True
        ElseIf tokenName = "UNLOCKED" Then
            If StyleFlagEnabled(tokenValueRaw, sep > 0) Then rng.Locked = False
        ElseIf tokenName = "HALIGN" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.HorizontalAlignment = TextToHorizontalAlignment(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply HAlign failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "VALIGN" And sep > 0 Then
            On Error Resume Next
            Err.Clear
            rng.VerticalAlignment = TextToVerticalAlignment(tokenValue)
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " apply VAlign failed " & ErrText() & _
                                    " style=" & ShortLogText(styleStr)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf tokenName = "LIST" And sep > 0 Then
            Dim listFml As String
            listFml = tokenValue
            On Error Resume Next
            Err.Clear
            rng.Validation.Delete
            If Err.Number <> 0 Then
                LogImportDiagnostic "WARN sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " delete validation failed " & ErrText()
                Err.Clear
            End If
            rng.Validation.Add Type:=xlValidateList, Formula1:=listFml
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & rng.Worksheet.codeName & _
                                    " addr=" & rng.Address(False, False) & _
                                    " add validation failed " & ErrText() & _
                                    " formula=" & ShortLogText(listFml)
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf p <> "" Then
            LogImportDiagnostic "WARN sheet=" & rng.Worksheet.codeName & _
                                " addr=" & rng.Address(False, False) & _
                                " unknown style token=" & ShortLogText(p) & _
                                " style=" & ShortLogText(styleStr)
        End If
lblNextCellStyle:
    Next k
End Sub

Private Sub ApplyCellName(ws As Worksheet, rng As Range, nameStr As String)
    On Error Resume Next
    Dim excl As Long
    excl = InStr(nameStr, "!")
    If excl > 0 Then
        ' Sheet-scoped name (e.g. "Sheet1!myRange") -> add as local name on this sheet
        Err.Clear
        ws.Names.Add Mid(nameStr, excl + 1), rng
    Else
        ' Workbook-scoped
        Err.Clear
        ThisWorkbook.Names.Add nameStr, rng
    End If
    If Err.Number <> 0 Then
        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                            " addr=" & rng.Address(False, False) & _
                            " apply name failed " & ErrText() & _
                            " name=" & ShortLogText(nameStr)
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function ResolveMasterAddr(refAddr As String, slaveAddrs() As String, masterAddrs() As String, slaveCount As Long) As String
    ' If refAddr is itself a slave, return its recorded master; otherwise refAddr is the master
    Dim j As Long
    For j = 0 To slaveCount - 1
        If slaveAddrs(j) = refAddr Then
            ResolveMasterAddr = masterAddrs(j)
            Exit Function
        End If
    Next j
    ResolveMasterAddr = refAddr
End Function

Private Sub ReconstructMerges(ws As Worksheet, slaveAddrs() As String, masterAddrs() As String, slaveCount As Long)
    Dim processed() As Boolean
    ReDim processed(slaveCount - 1)
    Dim i As Long, j As Long
    Dim mAddr As String
    Dim masterRng As Range
    Dim slvRng As Range
    Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long

    For i = 0 To slaveCount - 1
        If Not processed(i) Then
            mAddr = masterAddrs(i)
            On Error Resume Next
            Err.Clear
            Set masterRng = ws.Range(mAddr)
            If Err.Number <> 0 Or masterRng Is Nothing Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " reconstruct merge master failed master=" & ShortLogText(mAddr) & _
                                    " " & ErrText()
                Err.Clear
                On Error GoTo 0
                GoTo lblNextMerge
            End If
            On Error GoTo 0
            minRow = masterRng.Row
            maxRow = masterRng.Row
            minCol = masterRng.Column
            maxCol = masterRng.Column

            For j = 0 To slaveCount - 1
                If masterAddrs(j) = mAddr Then
                    On Error Resume Next
                    Err.Clear
                    Set slvRng = ws.Range(slaveAddrs(j))
                    If Err.Number <> 0 Or slvRng Is Nothing Then
                        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                            " reconstruct merge slave failed slave=" & ShortLogText(slaveAddrs(j)) & _
                                            " master=" & ShortLogText(mAddr) & _
                                            " " & ErrText()
                        Err.Clear
                        On Error GoTo 0
                        GoTo lblNextSlave
                    End If
                    On Error GoTo 0
                    If slvRng.Row < minRow Then minRow = slvRng.Row
                    If slvRng.Row > maxRow Then maxRow = slvRng.Row
                    If slvRng.Column < minCol Then minCol = slvRng.Column
                    If slvRng.Column > maxCol Then maxCol = slvRng.Column
                    processed(j) = True
                End If
lblNextSlave:
            Next j

            On Error Resume Next
            Err.Clear
            ws.Range(ws.Cells(minRow, minCol), ws.Cells(maxRow, maxCol)).Merge
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " merge failed range=" & _
                                    ws.Range(ws.Cells(minRow, minCol), ws.Cells(maxRow, maxCol)).Address(False, False) & _
                                    " " & ErrText()
                Err.Clear
            End If
            On Error GoTo 0
        End If
lblNextMerge:
    Next i
End Sub

Public Sub ExportSheetToMDFile(ws As Worksheet, filePath As String)
    SaveAsUTF8 filePath, GenerateSheetMapMarkdown(ws)
End Sub

Public Function ColLetterToNum(col As String) As Long
    Dim n As Long
    Dim i As Long
    For i = 1 To Len(col)
        n = n * 26 + (Asc(UCase(Mid(col, i, 1))) - 64)
    Next i
    ColLetterToNum = n
End Function

Public Function ColNumToLetter(colNum As Long) As String
    Dim s As String
    Dim n As Long
    n = colNum
    Do While n > 0
        Dim r As Long
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColNumToLetter = s
End Function

Private Function LongArrayToRangeStr(nums() As Long, count As Long) As String
    LongArrayToRangeStr = NumArrayToRangeStr(nums, count, False)
End Function

Private Function ColNumArrayToLetterRangeStr(nums() As Long, count As Long) As String
    ColNumArrayToLetterRangeStr = NumArrayToRangeStr(nums, count, True)
End Function

Private Function NumArrayToRangeStr(nums() As Long, count As Long, asLetter As Boolean) As String
    If count = 0 Then Exit Function
    Dim result As String
    Dim rangeStart As Long: rangeStart = nums(0)
    Dim rangeEnd   As Long: rangeEnd   = nums(0)
    Dim i As Long
    For i = 1 To count - 1
        If nums(i) = rangeEnd + 1 Then
            rangeEnd = nums(i)
        Else
            If Len(result) > 0 Then result = result & ", "
            result = result & FormatNumRange(rangeStart, rangeEnd, asLetter)
            rangeStart = nums(i)
            rangeEnd   = nums(i)
        End If
    Next i
    If Len(result) > 0 Then result = result & ", "
    NumArrayToRangeStr = result & FormatNumRange(rangeStart, rangeEnd, asLetter)
End Function

Private Function FormatNumRange(s As Long, e As Long, asLetter As Boolean) As String
    If asLetter Then
        FormatNumRange = IIf(s = e, ColNumToLetter(s), ColNumToLetter(s) & "-" & ColNumToLetter(e))
    Else
        FormatNumRange = IIf(s = e, CStr(s), CStr(s) & "-" & CStr(e))
    End If
End Function

Private Function CollectHiddenRowRanges(ws As Worksheet) As String
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ur.Row + ur.Rows.count - 1

    Dim nums() As Long
    Dim count As Long
    ReDim nums(lastRow)

    Dim rr As Long
    For rr = 1 To lastRow
        If ws.Rows(rr).Hidden Then
            nums(count) = rr
            count = count + 1
        End If
    Next rr

    CollectHiddenRowRanges = LongArrayToRangeStr(nums, count)
End Function

Private Function CollectHiddenColRanges(ws As Worksheet) As String
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Function

    Dim lastCol As Long
    lastCol = ur.Column + ur.Columns.count - 1

    Dim nums() As Long
    Dim count As Long
    ReDim nums(lastCol)

    Dim c As Long
    For c = 1 To lastCol
        If ws.Columns(c).Hidden Then
            nums(count) = c
            count = count + 1
        End If
    Next c

    CollectHiddenColRanges = ColNumArrayToLetterRangeStr(nums, count)
End Function

Private Sub ParseRangeStrToNums(rangeStr As String, ByRef nums() As Long, ByRef count As Long)
    count = 0
    ReDim nums(0)
    Dim parts() As String
    parts = Split(rangeStr, ",")
    Dim p As Variant
    For Each p In parts
        Dim token As String
        token = Trim(CStr(p))
        If Len(token) = 0 Then GoTo NextPart
        Dim dashPos As Long
        dashPos = InStr(token, "-")
        If dashPos > 0 Then
            Dim fromN As Long
            Dim toN As Long
            fromN = CLng(Left(token, dashPos - 1))
            toN = CLng(Mid(token, dashPos + 1))
            Dim n As Long
            For n = fromN To toN
                ReDim Preserve nums(count)
                nums(count) = n
                count = count + 1
            Next n
        Else
            ReDim Preserve nums(count)
            nums(count) = CLng(token)
            count = count + 1
        End If
NextPart:
    Next p
End Sub

Private Function ColLetterRangeStrToNumRangeStr(colRangeStr As String) As String
    Dim parts() As String
    parts = Split(colRangeStr, ",")
    Dim result As String
    Dim p As Variant
    For Each p In parts
        Dim token As String
        token = Trim(CStr(p))
        If Len(token) = 0 Then GoTo NextPart
        If Len(result) > 0 Then result = result & ", "
        Dim dashPos As Long
        dashPos = InStr(token, "-")
        If dashPos > 0 Then
            result = result & _
                CStr(ColLetterToNum(Trim(Left(token, dashPos - 1)))) & _
                "-" & _
                CStr(ColLetterToNum(Trim(Mid(token, dashPos + 1))))
        Else
            result = result & CStr(ColLetterToNum(token))
        End If
NextPart:
    Next p
    ColLetterRangeStrToNumRangeStr = result
End Function

Private Sub ApplyHiddenRows(ws As Worksheet, rangeStr As String)
    Dim nums() As Long
    Dim count As Long
    ParseRangeStrToNums rangeStr, nums, count
    If count = 0 Then Exit Sub

    Dim urLastRow As Long
    On Error Resume Next
    Err.Clear
    urLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    If Err.Number <> 0 Then
        LogImportDiagnostic "WARN sheet=" & ws.codeName & " read UsedRange for hidden rows failed " & ErrText()
        Err.Clear
    End If
    On Error GoTo 0

    Dim maxRow As Long
    maxRow = urLastRow
    Dim i As Long
    For i = 0 To count - 1
        If nums(i) > maxRow Then maxRow = nums(i)
    Next i

    On Error Resume Next
    Err.Clear
    ws.Rows("1:" & maxRow).Hidden = False
    If Err.Number <> 0 Then
        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                            " unhide rows failed maxRow=" & maxRow & _
                            " " & ErrText()
        Err.Clear
    End If
    On Error GoTo 0
    For i = 0 To count - 1
        If nums(i) >= 1 And nums(i) <= 1048576 Then
            On Error Resume Next
            Err.Clear
            ws.Rows(nums(i)).Hidden = True
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " hide row failed row=" & nums(i) & _
                                    " " & ErrText()
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub ApplyHiddenCols(ws As Worksheet, rangeStr As String)
    Dim numRangeStr As String
    numRangeStr = ColLetterRangeStrToNumRangeStr(rangeStr)

    Dim nums() As Long
    Dim count As Long
    ParseRangeStrToNums numRangeStr, nums, count
    If count = 0 Then Exit Sub

    Dim urLastCol As Long
    On Error Resume Next
    Err.Clear
    urLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.count - 1
    If Err.Number <> 0 Then
        LogImportDiagnostic "WARN sheet=" & ws.codeName & " read UsedRange for hidden cols failed " & ErrText()
        Err.Clear
    End If
    On Error GoTo 0

    Dim maxCol As Long
    maxCol = urLastCol
    Dim i As Long
    For i = 0 To count - 1
        If nums(i) > maxCol Then maxCol = nums(i)
    Next i

    On Error Resume Next
    Err.Clear
    ws.Columns("A:" & ColNumToLetter(maxCol)).Hidden = False
    If Err.Number <> 0 Then
        LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                            " unhide cols failed maxCol=" & maxCol & _
                            " " & ErrText()
        Err.Clear
    End If
    On Error GoTo 0
    For i = 0 To count - 1
        If nums(i) >= 1 And nums(i) <= 16384 Then
            On Error Resume Next
            Err.Clear
            ws.Columns(nums(i)).Hidden = True
            If Err.Number <> 0 Then
                LogImportDiagnostic "ERROR sheet=" & ws.codeName & _
                                    " hide col failed col=" & nums(i) & _
                                    " " & ErrText()
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' ============================================================
' Internationalization — T(), Fmt(), SetLang(), GetLangCode()
' ============================================================

' Returns localized text for key, falling back through: target lang -> en.ini -> defaultVal -> key.
' INI files live in {ThisWorkbook.Path}\lang\<code>.ini (BOM-less UTF-8).
' Use Fmt() separately for {0}/{1}/... placeholder substitution.
Public Function t(key As String, Optional defaultVal As String = "") As String
    Dim code As String: code = GetLangCode()
    If g_LangCode <> code Or g_LangCache Is Nothing Then
        g_LangCode = code
        Set g_LangCache = ParseINI(ThisWorkbook.Path & "\lang\" & code & ".ini")
        If code <> "en" Then
            Set g_EnCache = ParseINI(ThisWorkbook.Path & "\lang\en.ini")
        Else
            Set g_EnCache = Nothing
        End If
    End If
    If Not g_LangCache Is Nothing Then
        If g_LangCache.Exists(key) Then t = g_LangCache(key): Exit Function
    End If
    If Not g_EnCache Is Nothing Then
        If g_EnCache.Exists(key) Then t = g_EnCache(key): Exit Function
    End If
    If Len(defaultVal) > 0 Then t = defaultVal: Exit Function
    t = key
End Function

' Replaces {0}, {1}, ... in text with the supplied args.
Public Function Fmt(text As String, ParamArray args() As Variant) As String
    Dim result As String: result = text
    Dim i As Long
    For i = 0 To UBound(args)
        result = Replace(result, "{" & i & "}", CStr(args(i)))
    Next i
    Fmt = result
End Function

' Saves the language code to Registry and resets the cache.
' Pass an empty string to revert to system-language auto-detection.
Public Sub SetLang(langCode As String)
    SaveSetting "xlsm_devkit", "Language", "Code", langCode
    
    g_LangCode = ""
    Set g_LangCache = Nothing
    Set g_EnCache = Nothing
End Sub

' Returns the active language code (Registry > OS detection > "en").
Public Function GetLangCode() As String
    Dim code As String
    code = GetSetting("xlsm_devkit", "Language", "Code", "")
    If Len(code) > 0 Then GetLangCode = code: Exit Function
    If m_LangCodeDetected = "" Then
        m_LangCodeDetected = DetectSystemLang()
    End If
    GetLangCode = m_LangCodeDetected
End Function

' Returns a value from the [meta] section of the current language INI.
Public Function GetLangMeta(metaKey As String) As String
    GetLangMeta = t("meta." & metaKey, "")
End Function

' Parses a BOM-less UTF-8 INI file into a Dictionary keyed as "section.key".
' Returns Nothing if the file does not exist.
' INI values may use \n as a newline escape.
Public Function ParseINI(filePath As String) As Object
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function

    Dim content As String: content = ReadUTF8(filePath)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim lines() As String
    lines = Split(Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf), vbLf)

    Dim currentSection As String
    Dim i As Long, eq As Long, cb As Long
    Dim ln As String, k As String, v As String
    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) = 0 Then GoTo NextINILine
        If Left(ln, 1) = ";" Or Left(ln, 1) = "#" Then GoTo NextINILine
        If Left(ln, 1) = "[" Then
            cb = InStr(ln, "]")
            If cb > 2 Then currentSection = Mid(ln, 2, cb - 2)
        Else
            eq = InStr(ln, "=")
            If eq > 1 And Len(currentSection) > 0 Then
                k = currentSection & "." & Trim(Left(ln, eq - 1))
                v = Replace(Mid(ln, eq + 1), "\n", vbLf)
                dict(k) = v
            End If
        End If
NextINILine:
    Next i
    Set ParseINI = dict
End Function

' Detects the system UI language and checks for a matching lang INI file.
' Priority: lang\<lang>-<region>.ini  ->  lang\<lang>.ini  ->  "en"
Private Function DetectSystemLang() As String
    Const LOCALE_USER_DEFAULT     As Long = &H400
    Const LOCALE_SISO639LANGNAME  As Long = &H59
    Const LOCALE_SISO3166CTRYNAME As Long = &H5A
    Dim bufLang As String: bufLang = Space(10)
    Dim bufCtry As String: bufCtry = Space(10)
    Dim nLang As Long, nCtry As Long
    nLang = GetLocaleInfoA(LOCALE_USER_DEFAULT, LOCALE_SISO639LANGNAME, bufLang, Len(bufLang))
    nCtry = GetLocaleInfoA(LOCALE_USER_DEFAULT, LOCALE_SISO3166CTRYNAME, bufCtry, Len(bufCtry))
    If nLang > 1 Then
        Dim langCode As String: langCode = LCase(Left(bufLang, nLang - 1))
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        Dim langDir As String: langDir = ThisWorkbook.Path & "\lang\"
        If nCtry > 1 Then
            Dim combined As String
            combined = langCode & "-" & LCase(Left(bufCtry, nCtry - 1))
            If fso.FileExists(langDir & combined & ".ini") Then
                DetectSystemLang = combined: Exit Function
            End If
        End If
        If fso.FileExists(langDir & langCode & ".ini") Then
            DetectSystemLang = langCode: Exit Function
        End If
    End If
    DetectSystemLang = "en"
End Function

Private Function GetSystemAnsiCharset() As String
    Dim cp As Long
    cp = GetACP()
    Select Case cp
        Case 932:  GetSystemAnsiCharset = "Shift_JIS"
        Case 936:  GetSystemAnsiCharset = "GB2312"
        Case 949:  GetSystemAnsiCharset = "ks_c_5601-1987"
        Case 950:  GetSystemAnsiCharset = "Big5"
        Case 1250: GetSystemAnsiCharset = "Windows-1250"
        Case 1251: GetSystemAnsiCharset = "Windows-1251"
        Case 1252: GetSystemAnsiCharset = "Windows-1252"
        Case 1253: GetSystemAnsiCharset = "Windows-1253"
        Case 1254: GetSystemAnsiCharset = "Windows-1254"
        Case 1255: GetSystemAnsiCharset = "Windows-1255"
        Case 1256: GetSystemAnsiCharset = "Windows-1256"
        Case 1257: GetSystemAnsiCharset = "Windows-1257"
        Case 1258: GetSystemAnsiCharset = "Windows-1258"
        Case 874:  GetSystemAnsiCharset = "Windows-874"
        Case Else: GetSystemAnsiCharset = "Windows-" & cp
    End Select
End Function


' ── Dev/Release lifecycle ────────────────────────────────────────────────────
Private Sub InitDevMode()
    If UCase(Left(ThisWorkbook.Name, 4)) = "DEV_" Then
        MsgBox t("msg.init_already_dev", "This workbook already has the DEV_ prefix. Call InitDevMode() from a production workbook."), vbExclamation
        Exit Sub
    End If
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox t("msg.init_unsaved", "Please save this workbook before calling InitDevMode()."), vbExclamation
        Exit Sub
    End If
    If Not CheckVBProjectAccess() Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim devPath As String
    devPath = ThisWorkbook.Path & "\DEV_" & ThisWorkbook.Name

    If fso.FileExists(devPath) Then
        If MsgBox(Fmt(t("msg.init_overwrite_confirm", "'{0}' already exists. Overwrite?"), devPath), _
                  vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    Else
        If MsgBox(Fmt(t("msg.init_confirm", "Create DEV_ copy as '{0}'?" & vbLf & _
                        "All devkit_* files found in src/ will be imported."), devPath), _
                  vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim devWb As Workbook
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo ErrHandler
    ThisWorkbook.SaveCopyAs devPath

    Application.EnableEvents = False
    Set devWb = Workbooks.Open(devPath, UpdateLinks:=0)
    Application.EnableEvents = prevEvents

    Dim srcDir As String
    srcDir = ThisWorkbook.Path & "\src"
    If fso.FolderExists(srcDir) Then
        Dim srcFolder As Object
        Set srcFolder = fso.GetFolder(srcDir)
        Dim f As Object, fExt As String, fBase As String
        For Each f In srcFolder.Files
            fExt  = LCase(fso.GetExtensionName(f.Path))
            fBase = fso.GetBaseName(f.Path)
            If LCase(Left(fBase, 7)) = "devkit_" And (fExt = "bas" Or fExt = "frm") Then
                If Not ImportComponentIntoProject(devWb.VBProject, f.Path) Then
                    MsgBox "Failed to import component: " & f.Name, vbExclamation
                End If

            End If
        Next f
    End If

    devWb.Save
    devWb.Close SaveChanges:=False

    MsgBox Fmt(t("msg.init_complete", "DEV_ copy created: {0}" & vbLf & vbLf & _
                 "Close this workbook WITHOUT saving (Ctrl+W -> Don't Save) to keep the production file clean."), _
               devPath), vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = prevEvents
    Dim initErrDesc As String: initErrDesc = Err.Description
    On Error Resume Next
    If Not devWb Is Nothing Then devWb.Close SaveChanges:=False
    MsgBox Fmt(t("msg.init_failed", "Failed to create DEV_ copy: {0}"), initErrDesc), vbCritical
End Sub


Private Sub SaveAsRelease()
    If UCase(Left(ThisWorkbook.Name, 4)) <> "DEV_" Then
        MsgBox t("msg.release_no_dev_prefix", "This workbook does not have the DEV_ prefix. SaveAsRelease() must be called from a DEV_ workbook."), vbExclamation
        Exit Sub
    End If
    If Not CheckVBProjectAccess() Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim releasePath As String
    releasePath = ThisWorkbook.Path & "\" & Mid(ThisWorkbook.Name, 5)

    If fso.FileExists(releasePath) Then
        If MsgBox(Fmt(t("msg.release_overwrite_confirm", "'{0}' already exists. Overwrite?"), releasePath), _
                  vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    Else
        If MsgBox(Fmt(t("msg.release_confirm", "Save release copy as '{0}'?" & vbLf & _
                        "All devkit modules will be removed from the copy."), releasePath), _
                  vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim relWb As Workbook
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo ErrHandler
    ThisWorkbook.SaveCopyAs releasePath

    Application.EnableEvents = False
    Set relWb = Workbooks.Open(releasePath, UpdateLinks:=0)
    Application.EnableEvents = prevEvents

    ' Collect devkit component names first (avoid modifying collection while iterating)
    Dim comp As Object
    Dim toRemove() As String
    ReDim toRemove(relWb.VBProject.VBComponents.Count)
    Dim removeCount As Long
    removeCount = 0
    For Each comp In relWb.VBProject.VBComponents
        If IsDevkitComponent(comp.Name) And comp.Type <> 100 Then
            toRemove(removeCount) = comp.Name
            removeCount = removeCount + 1
        End If
    Next comp

    Dim i As Long
    For i = 0 To removeCount - 1
        On Error Resume Next
        relWb.VBProject.VBComponents.Remove relWb.VBProject.VBComponents(toRemove(i))
        On Error GoTo ErrHandler
    Next i

    relWb.Save
    relWb.Close SaveChanges:=False

    MsgBox Fmt(t("msg.release_complete", "Release copy saved: {0}"), releasePath), vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = prevEvents
    Dim relErrDesc As String: relErrDesc = Err.Description
    On Error Resume Next
    If Not relWb Is Nothing Then relWb.Close SaveChanges:=False
    MsgBox Fmt(t("msg.release_failed", "Failed to create release copy: {0}"), relErrDesc), vbCritical
End Sub


Private Function ImportComponentIntoProject(vbProj As Object, utf8FilePath As String) As Boolean
    Dim backupPath As String
    backupPath = utf8FilePath & "_"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler

    Dim baseName As String
    baseName = fso.GetBaseName(utf8FilePath)

    fso.CopyFile utf8FilePath, backupPath
    ConvertEncoding utf8FilePath, "UTF-8", GetSystemAnsiCharset()

    On Error Resume Next
    Dim existComp As Object
    Set existComp = vbProj.VBComponents(baseName)
    If Not existComp Is Nothing Then
        If existComp.Type <> 100 Then vbProj.VBComponents.Remove existComp
    End If
    Set existComp = Nothing
    On Error GoTo ErrHandler

    vbProj.VBComponents.Import utf8FilePath

    fso.DeleteFile utf8FilePath, True
    fso.MoveFile backupPath, utf8FilePath
    ImportComponentIntoProject = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If fso.FileExists(backupPath) Then
        If fso.FileExists(utf8FilePath) Then fso.DeleteFile utf8FilePath, True
        fso.MoveFile backupPath, utf8FilePath
    End If
    ImportComponentIntoProject = False
End Function


Private Function IsDevkitComponent(compName As String) As Boolean
    Dim lowerName As String
    lowerName = LCase(compName)
    IsDevkitComponent = (lowerName = LCase(MODULE_NAME)) Or (Left(lowerName, 7) = "devkit_")
End Function
