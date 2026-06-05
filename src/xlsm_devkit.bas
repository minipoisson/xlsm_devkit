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

Public g_LangCode  As String
Public g_LangCache As Object
Public g_EnCache   As Object
Private m_LangCodeDetected As String

' NOTE: This module itself is NOT imported by ImportAllComponents()
' because a module cannot delete or overwrite itself.

Public Sub ExportAllModulesFormsSheetMaps()
    If MsgBox(t("msg.export_all_confirm", "Export all modules, forms, and sheet maps to src/ and sheet/?"), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ExportAllComponents True
    ExportAllSheetMapsToMD True
End Sub

Public Sub ImportAllModulesFormsSheetMaps()
    If MsgBox(t("msg.import_all_confirm", "Import all modules, forms, and sheet maps from src/ and sheet/?"), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
        If SKIP_DEVKIT_MODULES And (Left(comp.Name, 7) = "devkit_" _
        Or Left(comp.Name, 11) = "xlsm_devkit") Then GoTo lblNext
        
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
            If compName = MODULE_NAME Then GoTo lblNext              ' xlsm_devkit (self)
            If SKIP_DEVKIT_MODULES And Left(compName, 7) = "devkit_" Then GoTo lblNext  ' optional devkit modules
            If compName = "devkit_Move" And isMoveLocked Then         ' devkit_Move on call stack
                skipCount = skipCount + 1: GoTo lblNext
            End If
            If ComponentExists(compName) Then
                If ReplaceExistingComponentCodeFromFile(compName, srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            Else
                If ImportNewComponentFromFile(srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            End If
        ElseIf ext = "frm" Then
            If SKIP_DEVKIT_MODULES And Left(compName, 7) = "devkit_" Then GoTo lblNext  ' optional devkit forms
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
        If frm.Name = formName Then IsFormCurrentlyLoaded = True: Exit Function
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

    Dim i As Long
    Dim output As String
    Dim lineText As String
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        If Left$(LCase$(Trim$(lineText)), 13) <> "attribute vb_" Then
            If Len(output) > 0 Then output = output & vbCrLf
            output = output & lineText
        End If
    Next i

    StripAttributeLines = output
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

    Dim ws As Worksheet
    Dim sheetFolderPath As String
    Dim fileName As String
    Dim mdContent As String

    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sheetFolderPath) Then
        fso.CreateFolder sheetFolderPath
    End If

    For Each ws In ThisWorkbook.Worksheets
        mdContent = GenerateSheetMapMarkdown(ws)
        
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        
        SaveAsUTF8 fileName, mdContent
    Next
    
    MsgBox Fmt(t("msg.sheet_maps_exported", "All sheet maps exported. Saved to: {0}"), sheetFolderPath), vbInformation
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

    For Each ws In ThisWorkbook.Worksheets
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        If fso.FileExists(fileName) Then
            mdContent = ReadUTF8(fileName)
            ApplySheetMapMarkdown ws, mdContent
        End If
    Next ws

    MsgBox t("msg.sheet_maps_imported", "All sheet maps imported."), vbInformation
    GoTo lblFin
lblErr:
    MsgBox t("msg.sheet_map_import_error", _
    "Error importing sheet maps. Some sheets may be partially updated.") _
    & vbCrLf & Err.Description, vbExclamation
lblFin:
    Application.EnableEvents = preEvents
    Application.Calculation = preCalcMode
    Application.ScreenUpdating = preScreenUpdate
End Sub

Private Function GenerateSheetMapMarkdown(ws As Worksheet) As String
    Dim rng As Range
    Dim mapText As String
    Dim role As String
    Dim cellName As String
    Dim styleParts As String
    Dim mergeMarker As String
    
    mapText = "# Sheet Configuration" & vbCrLf
    mapText = mapText & "- VBA CodeName: " & ws.codeName & vbCrLf
    mapText = mapText & "- Excel UI Name: " & ws.Name & vbCrLf
    Dim hiddenRowStr As String
    Dim hiddenColStr As String
    hiddenRowStr = CollectHiddenRowRanges(ws)
    hiddenColStr = CollectHiddenColRanges(ws)
    If Len(hiddenRowStr) > 0 Then
        mapText = mapText & "- Hidden Rows: " & hiddenRowStr & vbCrLf
    End If
    If Len(hiddenColStr) > 0 Then
        mapText = mapText & "- Hidden Columns: " & hiddenColStr & vbCrLf
    End If
    mapText = mapText & vbCrLf
    
    mapText = mapText & "| Address | Name | Value / Label | Formula | Style |" & vbCrLf
    mapText = mapText & "| :--- | :--- | :--- | :--- | :--- |" & vbCrLf
    
    For Each rng In ws.UsedRange
        If rng.MergeCells And Not (rng.Row = rng.MergeArea.Row And rng.Column = rng.MergeArea.Column) Then
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

            styleParts = ""
            If rng.Interior.ColorIndex <> xlNone Then
                styleParts = "BG:" & ColorToHex(rng.Interior.Color)
            End If
            If rng.Font.ColorIndex <> xlColorIndexAutomatic And rng.Font.Color <> 0 Then
                If styleParts <> "" Then styleParts = styleParts & "; "
                styleParts = styleParts & "FG:" & ColorToHex(rng.Font.Color)
            End If
            If rng.Font.Size <> Application.StandardFontSize Then
                If styleParts <> "" Then styleParts = styleParts & "; "
                styleParts = styleParts & "FontSize:" & rng.Font.Size
            End If

            mapText = mapText & "| " & rng.Address(False, False) & _
                      " | " & cellName & _
                      " | " & mergeMarker & _
                      " | -" & _
                      " | " & IIf(styleParts = "", "-", styleParts) & " |" & vbCrLf
        ElseIf rng.Value <> "" Or rng.HasFormula Or rng.Interior.ColorIndex <> xlNone Then
            ' Normal cell (or master cell of a merged range)
            cellName = ""
            On Error Resume Next
            cellName = rng.Name.Name
            If cellName = "" Then cellName = "-"
            On Error GoTo 0

            styleParts = ""
            If rng.Interior.ColorIndex <> xlNone Then
                styleParts = "BG:" & ColorToHex(rng.Interior.Color)
            End If
            If rng.Font.ColorIndex <> xlColorIndexAutomatic And rng.Font.Color <> 0 Then
                If styleParts <> "" Then styleParts = styleParts & "; "
                styleParts = styleParts & "FG:" & ColorToHex(rng.Font.Color)
            End If
            If rng.Font.Size <> Application.StandardFontSize Then
                If styleParts <> "" Then styleParts = styleParts & "; "
                styleParts = styleParts & "FontSize:" & rng.Font.Size
            End If
            Dim valFormula As String
            valFormula = ""
            On Error Resume Next
            Dim valType As Long
            valType = rng.Validation.Type
            If Err.Number = 0 And valType = xlValidateList Then
                valFormula = rng.Validation.Formula1
            End If
            On Error GoTo 0
            If valFormula <> "" Then
                If styleParts <> "" Then styleParts = styleParts & "; "
                styleParts = styleParts & "List:" & Replace(valFormula, "|", ChrW(&HFF5C))
            End If
            role = IIf(styleParts = "", "-", styleParts)

            mapText = mapText & "| " & rng.Address(False, False) & _
                      " | " & cellName & _
                      " | " & EscapeCellValue(rng.Value) & _
                      " | " & IIf(rng.HasFormula, "`" & rng.Formula & "`", "-") & _
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
    Dim shapeRows As String

    shapeRows = ""

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

        shpLabel = Replace(shpLabel, vbCrLf, " ")
        shpLabel = Replace(shpLabel, vbCr, " ")
        shpLabel = Replace(shpLabel, vbLf, " ")
        shpLabel = Replace(shpLabel, "|", ChrW(&HFF5C))

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
            shpFormula = "`" & shpFml & "`"
        End If
        On Error GoTo 0

        If shpLabel <> "" Or shpFormula <> "-" Or shpOnAction <> "" Then
            shapeRows = shapeRows & "| " & shp.TopLeftCell.Address(False, False) & _
                        " | " & shp.Name & _
                        " | " & IIf(shpLabel <> "", shpLabel, "-") & _
                        " | " & shpFormula & _
                        " | " & IIf(shpOnAction <> "", shpOnAction, "-") & _
                        " | " & IIf(shpStyle <> "", shpStyle, "-") & " |" & vbCrLf
        End If
lblFinShp:
        Err.Clear
    Next shp

    If shapeRows <> "" Then
        mapText = mapText & vbCrLf & "## Shapes" & vbCrLf & vbCrLf
        mapText = mapText & "| Address | Name | Label | Formula | OnAction | Style |" & vbCrLf
        mapText = mapText & "| :--- | :--- | :--- | :--- | :--- | :--- |" & vbCrLf
        mapText = mapText & shapeRows
    End If

    GenerateSheetMapMarkdown = mapText
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
    Dim refAddr As String
    Dim mAddr As String
    Dim slaveAddrs() As String
    Dim masterAddrs() As String
    Dim slaveCount As Long
    Dim hiddenRowsStr As String
    Dim hiddenColsStr As String
    hiddenRowsStr = ""
    hiddenColsStr = ""

    ' Pre-clear cell data, formatting, merges, and validation
    On Error Resume Next
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
    On Error Resume Next
    ws.Cells.Validation.Delete
    On Error GoTo 0

    ' Normalize line endings and split
    normContent = Replace(mdContent, vbCrLf, vbLf)
    normContent = Replace(normContent, vbCr, vbLf)
    lines = Split(normContent, vbLf)

    inCellTable = False
    slaveCount = 0
    ReDim slaveAddrs(UBound(lines))
    ReDim masterAddrs(UBound(lines))

    For i = 0 To UBound(lines)
        line = Trim(lines(i))

        If line = "## Shapes" Then Exit For

        If Left(line, 10) = "| Address " Then
            inCellTable = True
        ElseIf Left(line, 15) = "- Hidden Rows: " Then
            hiddenRowsStr = Trim(Mid(line, 16))
        ElseIf Left(line, 18) = "- Hidden Columns: " Then
            hiddenColsStr = Trim(Mid(line, 19))
        ElseIf inCellTable And Left(line, 1) = "|" And Left(line, 5) <> "| :-" Then
            cols = ParseMDTableRow(line)
            If UBound(cols) >= 4 Then
                addr = Trim(cols(0))
                cName = Trim(cols(1))
                cValue = Trim(cols(2))
                cFormula = Trim(cols(3))
                cStyle = Trim(cols(4))

                Set rng = Nothing
                On Error Resume Next
                Set rng = ws.Range(addr)
                On Error GoTo 0

                If Not rng Is Nothing Then
                    If cValue = "!merged_left" Or cValue = "!merged_up" Or cValue = "!merged_ul" Then
                        ' Slave cell: apply style and name, record for Pass 2
                        ApplyCellStyle rng, cStyle
                        If cName <> "-" Then ApplyCellName ws, rng, cName

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
                            On Error Resume Next
                            rng.Formula = fml
                            On Error GoTo 0
                        ElseIf cValue <> "-" And cValue <> "" Then
                            On Error Resume Next
                            rng.Value = UnescapeCellValue(cValue)
                            On Error GoTo 0
                        End If
                        ApplyCellStyle rng, cStyle
                        If cName <> "-" Then ApplyCellName ws, rng, cName
                    End If
                End If
            End If
        End If
    Next i

    ' Pass 2: reconstruct merged ranges
    If slaveCount > 0 Then ReconstructMerges ws, slaveAddrs, masterAddrs, slaveCount

    If Len(hiddenRowsStr) > 0 Then ApplyHiddenRows ws, hiddenRowsStr
    If Len(hiddenColsStr) > 0 Then ApplyHiddenCols ws, hiddenColsStr
End Sub

Private Function EscapeCellValue(cellValue As String) As String
    Dim v As String
    v = cellValue
    v = Replace(v, "\\", "\\\\")
    v = Replace(v, vbCrLf, "\n")
    v = Replace(v, vbCr, "\n")
    v = Replace(v, vbLf, "\n")
    v = Replace(v, "|", ChrW(&HFF5C))
    EscapeCellValue = v
End Function

Private Function ColorToHex(colorVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorVal And 255
    g = (colorVal \ 256) And 255
    b = (colorVal \ 65536) And 255
    ColorToHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

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
        Dim v As Integer: v = CInt(m.SubMatches(0))
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
    parts = Split(s, "|")
    Dim k As Integer
    For k = 0 To UBound(parts)
        parts(k) = Trim(parts(k))
    Next k
    ParseMDTableRow = parts
End Function

Private Function UnescapeCellValue(v As String) As String
    ' Reverse of EscapeCellValue; use placeholder to avoid \n vs \\n ambiguity
    Dim PH As String: PH = Chr(1)
    v = Replace(v, ChrW(&HFF5C), "|")
    v = Replace(v, "\\", PH)
    v = Replace(v, "\n", vbLf)
    v = Replace(v, PH, "\")
    UnescapeCellValue = v
End Function

Private Function HexToRGB(hexStr As String) As Long
    If Left(hexStr, 1) = "#" Then hexStr = Mid(hexStr, 2)
    HexToRGB = RGB(CLng("&H" & Left(hexStr, 2)), CLng("&H" & Mid(hexStr, 3, 2)), CLng("&H" & Right(hexStr, 2)))
End Function

Private Sub ApplyCellStyle(rng As Range, styleStr As String)
    If styleStr = "-" Or styleStr = "" Then Exit Sub
    Dim parts() As String
    parts = Split(styleStr, "; ")
    Dim p As String
    Dim k As Integer
    For k = 0 To UBound(parts)
        p = Trim(parts(k))
        If Left(p, 3) = "BG:" Then
            rng.Interior.Color = HexToRGB(Mid(p, 4))
        ElseIf Left(p, 3) = "FG:" Then
            rng.Font.Color = HexToRGB(Mid(p, 4))
        ElseIf Left(p, 9) = "FontSize:" Then
            rng.Font.Size = CDbl(Mid(p, 10))
        ElseIf Left(p, 5) = "List:" Then
            Dim listFml As String
            listFml = Replace(Mid(p, 6), ChrW(&HFF5C), "|")
            On Error Resume Next
            rng.Validation.Delete
            rng.Validation.Add Type:=xlValidateList, Formula1:=listFml
            On Error GoTo 0
        End If
    Next k
End Sub

Private Sub ApplyCellName(ws As Worksheet, rng As Range, nameStr As String)
    On Error Resume Next
    Dim excl As Long
    excl = InStr(nameStr, "!")
    If excl > 0 Then
        ' Sheet-scoped name (e.g. "Sheet1!myRange") -> add as local name on this sheet
        ws.Names.Add Mid(nameStr, excl + 1), rng
    Else
        ' Workbook-scoped
        ThisWorkbook.Names.Add nameStr, rng
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
            Set masterRng = ws.Range(mAddr)
            minRow = masterRng.Row
            maxRow = masterRng.Row
            minCol = masterRng.Column
            maxCol = masterRng.Column

            For j = 0 To slaveCount - 1
                If masterAddrs(j) = mAddr Then
                    Set slvRng = ws.Range(slaveAddrs(j))
                    If slvRng.Row < minRow Then minRow = slvRng.Row
                    If slvRng.Row > maxRow Then maxRow = slvRng.Row
                    If slvRng.Column < minCol Then minCol = slvRng.Column
                    If slvRng.Column > maxCol Then maxCol = slvRng.Column
                    processed(j) = True
                End If
            Next j

            ws.Range(ws.Cells(minRow, minCol), ws.Cells(maxRow, maxCol)).Merge
        End If
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
    If count = 0 Then Exit Function
    Dim result As String
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim i As Long
    rangeStart = nums(0)
    rangeEnd = nums(0)
    For i = 1 To count - 1
        If nums(i) = rangeEnd + 1 Then
            rangeEnd = nums(i)
        Else
            If Len(result) > 0 Then result = result & ", "
            result = result & IIf(rangeStart = rangeEnd, _
                CStr(rangeStart), CStr(rangeStart) & "-" & CStr(rangeEnd))
            rangeStart = nums(i)
            rangeEnd = nums(i)
        End If
    Next i
    If Len(result) > 0 Then result = result & ", "
    result = result & IIf(rangeStart = rangeEnd, _
        CStr(rangeStart), CStr(rangeStart) & "-" & CStr(rangeEnd))
    LongArrayToRangeStr = result
End Function

Private Function ColNumArrayToLetterRangeStr(nums() As Long, count As Long) As String
    If count = 0 Then Exit Function
    Dim result As String
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim i As Long
    rangeStart = nums(0)
    rangeEnd = nums(0)
    For i = 1 To count - 1
        If nums(i) = rangeEnd + 1 Then
            rangeEnd = nums(i)
        Else
            If Len(result) > 0 Then result = result & ", "
            result = result & IIf(rangeStart = rangeEnd, _
                ColNumToLetter(rangeStart), _
                ColNumToLetter(rangeStart) & "-" & ColNumToLetter(rangeEnd))
            rangeStart = nums(i)
            rangeEnd = nums(i)
        End If
    Next i
    If Len(result) > 0 Then result = result & ", "
    result = result & IIf(rangeStart = rangeEnd, _
        ColNumToLetter(rangeStart), _
        ColNumToLetter(rangeStart) & "-" & ColNumToLetter(rangeEnd))
    ColNumArrayToLetterRangeStr = result
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
    urLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    On Error GoTo 0

    Dim maxRow As Long
    maxRow = urLastRow
    Dim i As Long
    For i = 0 To count - 1
        If nums(i) > maxRow Then maxRow = nums(i)
    Next i

    ws.Rows("1:" & maxRow).Hidden = False
    For i = 0 To count - 1
        If nums(i) >= 1 And nums(i) <= 1048576 Then
            ws.Rows(nums(i)).Hidden = True
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
    urLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.count - 1
    On Error GoTo 0

    Dim maxCol As Long
    maxCol = urLastCol
    Dim i As Long
    For i = 0 To count - 1
        If nums(i) > maxCol Then maxCol = nums(i)
    Next i

    ws.Columns("A:" & ColNumToLetter(maxCol)).Hidden = False
    For i = 0 To count - 1
        If nums(i) >= 1 And nums(i) <= 16384 Then
            ws.Columns(nums(i)).Hidden = True
        End If
    Next i
End Sub

' ============================================================
' Internationalization ? T(), Fmt(), SetLang(), GetLangCode()
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
