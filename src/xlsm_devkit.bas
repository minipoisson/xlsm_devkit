Attribute VB_Name = "xlsm_devkit"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
Private Declare Function GetACP Lib "kernel32" () As Long
#End If

Private Const MODULE_NAME As String = "xlsm_devkit"

' NOTE: This module itself is NOT imported by ImportAllComponents()
' because a module cannot delete or overwrite itself.

Public Sub exportAllModulesFormsSheetMaps()
    If MsgBox("Export all modules, forms, and sheet maps to src/ and sheet/?", _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ExportAllComponents True
    ExportAllSheetMapsToMD True
End Sub

Public Sub importAllModulesFormsSheetMaps()
    If MsgBox("Import all modules, forms, and sheet maps from src/ and sheet/?", _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ImportAllComponents True
    ImportAllSheetMapsFromMD True
End Sub

Public Sub ShowInsertDeleteForm()
    frmInsertDelete.Show
End Sub

Public Sub callExportAllComponents()
    ExportAllComponents
End Sub

Public Sub callImportAllComponents()
    ImportAllComponents
End Sub

Public Sub callExportAllSheetMapsToMD()
    ExportAllSheetMapsToMD
End Sub

Public Sub callImportAllSheetMapsFromMD()
    ImportAllSheetMapsFromMD
End Sub


Private Function CheckVBProjectAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Excel macro settings."
        Exit Function
    End If
    On Error GoTo 0
    CheckVBProjectAccess = True
End Function

Sub ExportAllComponents(Optional skipConfirm As Boolean = False)
    If Not CheckVBProjectAccess() Then Exit Sub
    If Not skipConfirm Then
        If MsgBox("Export all modules and forms to src/?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim dirPath As String: dirPath = ThisWorkbook.Path & "\src"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dirPath) Then fso.CreateFolder dirPath

    Dim comp As Object, expPath As String, frxPath As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 3 Then
            expPath = dirPath & "\" & comp.Name & ".frm"
            frxPath = dirPath & "\" & comp.Name & ".frx"
            If fso.FileExists(frxPath) Then fso.DeleteFile frxPath, True
        Else
            expPath = dirPath & "\" & comp.Name & ".bas"
        End If
        comp.Export expPath
        ConvertEncoding expPath, GetSystemAnsiCharset(), "UTF-8"
    Next
    MsgBox "Export complete."
End Sub

Sub ImportAllComponents(Optional skipConfirm As Boolean = False)
    If Not CheckVBProjectAccess() Then Exit Sub
    If Not skipConfirm Then
        If MsgBox("Import all modules and forms from src/?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String: srcPath = ThisWorkbook.Path & "\src"
    If Not fso.FolderExists(srcPath) Then
        MsgBox "Source folder not found: " & srcPath, vbExclamation
        Exit Sub
    End If

    Dim successCount As Long, failCount As Long
    Dim srcFile As Object, ext As String, compName As String
    For Each srcFile In fso.GetFolder(srcPath).Files
        ext = LCase(fso.GetExtensionName(srcFile.Name))
        compName = Left(srcFile.Name, Len(srcFile.Name) - Len(ext) - 1)
        If ext = "bas" Then
            If compName = MODULE_NAME Then GoTo lblNext
            If ComponentExists(compName) Then
                If ReplaceExistingComponentCodeFromFile(compName, srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            Else
                If ImportNewComponentFromFile(srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
            End If
        ElseIf ext = "frm" Then
            If ImportForm(compName, srcFile.Path) Then successCount = successCount + 1 Else failCount = failCount + 1
        End If
lblNext:
    Next srcFile

    MsgBox "Import complete." & vbLf & _
           "Success: " & successCount & vbLf & _
           "Failed: " & failCount, _
           IIf(failCount = 0, vbInformation, vbExclamation)
End Sub

Private Sub ConvertEncoding(filePath As String, fromCharset As String, toCharset As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Open
    st.Type = 2
    st.Charset = fromCharset
    st.LoadFromFile filePath
    Dim content As String
    content = st.ReadText
    st.Close

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
        If MsgBox("Export all sheet maps to sheet/?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
    
    MsgBox "All sheet maps exported." & vbLf & "Saved to: " & sheetFolderPath, vbInformation
End Sub

Sub ImportAllSheetMapsFromMD(Optional skipConfirm As Boolean = False)
    If Not skipConfirm Then
        If MsgBox("Import all sheet maps from sheet/?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim ws As Worksheet
    Dim sheetFolderPath As String
    Dim fileName As String
    Dim mdContent As String

    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sheetFolderPath) Then
        MsgBox "Sheet folder not found: " & sheetFolderPath, vbExclamation
        Exit Sub
    End If

    For Each ws In ThisWorkbook.Worksheets
        fileName = sheetFolderPath & "\" & ws.codeName & ".md"
        If fso.FileExists(fileName) Then
            mdContent = ReadUTF8(fileName)
            ApplySheetMapMarkdown ws, mdContent
        End If
    Next ws

    MsgBox "All sheet maps imported.", vbInformation
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
    mapText = mapText & "- Excel UI Name: " & ws.Name & vbCrLf & vbCrLf
    
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
                            rng.Value = UnescapeCellValue(cValue)
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

Private Sub SaveAsUTF8(filePath As String, content As String)
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


