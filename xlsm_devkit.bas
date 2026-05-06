Attribute VB_Name = "xlsm_devkit"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
Private Declare Function GetACP Lib "kernel32" () As Long
#End If

Private Const MODULE_NAME As String = "xlsm_devkit"

' NOTE: This module itself is NOT imported by ImportAllModules()
' because a module cannot delete or overwrite itself.

Sub callExportAllModules()
    ExportAllModules
End Sub

Sub callImportAllModules()
    ImportAllModules
End Sub

Sub callExportAllSheetMapsToMD()
    ExportAllSheetMapsToMD
End Sub


Sub ExportAllModules()
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Excel macro settings."
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim dirPath As String
    dirPath = ThisWorkbook.Path & "\src"
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
    End If
    
    Dim comp As Object, expPath As String
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type <> 3 Then
            expPath = ThisWorkbook.Path & "\src\" & comp.Name & ".bas"
            If Dir(expPath) <> "" Then
                If MsgBox("Overwrite the following file?" & vbLf & expPath, vbYesNo + vbDefaultButton2) = vbYes Then
                    Kill expPath
                Else
                    GoTo lblContinue
                End If
            End If
            comp.Export expPath
            ConvertEncoding ThisWorkbook.Path & "\src\" & comp.Name & ".bas", GetSystemAnsiCharset(), "UTF-8"
        End If
lblContinue:
    Next
    MsgBox "Export complete."
End Sub

Sub ImportAllModules()
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Excel macro settings."
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim thisModule As String
    thisModule = MODULE_NAME

    Dim successCount As Long
    Dim failCount As Long
    Dim fName As String
    fName = Dir(ThisWorkbook.Path & "\src\*.bas")

    Do While fName <> ""
        Dim modName As String
        modName = Left(fName, Len(fName) - 4)

        If modName = thisModule Then
            GoTo lblContinue:
        End If
        
        If ComponentExists(modName) Then
            If ReplaceExistingComponentCodeFromFile(modName, ThisWorkbook.Path & "\src\" & fName) Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        Else
            If ImportNewComponentFromFile(ThisWorkbook.Path & "\src\" & fName) Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        End If

lblContinue:
        fName = Dir()
    Loop

    MsgBox "Import complete." & vbLf & _
           "Success: " & successCount & vbLf & _
           "Failed: " & failCount, _
           IIf(failCount = 0, vbInformation, vbExclamation)
End Sub

Private Sub ConvertEncoding(filePath As String, fromCharset As String, toCharset As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    
    st.Open
    st.Type = 2  ' text
    st.Charset = fromCharset
    st.LoadFromFile filePath
    Dim content As String
    content = st.ReadText
    st.Close
    
    st.Open
    st.Type = 2
    st.Charset = toCharset
    st.WriteText content
    st.SaveToFile filePath, 2  ' overwrite
    st.Close
End Sub

Private Function ReplaceExistingComponentCodeFromFile(modName As String, filePath As String) As Boolean
    Dim tempName As String
    tempName = BuildUniqueTempModuleName(modName)
    
    Dim tempPath As String
    tempPath = ThisWorkbook.Path & "\src\" & tempName & ".bas"
    
    Dim tempComp As Object
    Dim targetComp As Object
    Dim codeText As String
    
    On Error GoTo ErrHandler
    
    ' Update existing code via a temporary module to avoid removing the target component.
    FileCopy filePath, tempPath
    RenameModuleInFile tempPath, modName, tempName
    
    ' VBComponents.Import expects system ANSI encoding.
    ConvertEncoding tempPath, "UTF-8", GetSystemAnsiCharset()
    ThisWorkbook.VBProject.VBComponents.Import tempPath
    If FileExists(tempPath) Then Kill tempPath
    
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
    If FileExists(tempPath) Then Kill tempPath
    ReplaceExistingComponentCodeFromFile = False
End Function

Private Sub ImportDocumentModule(modName As String, filePath As String)
    ReplaceExistingComponentCodeFromFile modName, filePath
End Sub

Private Function ImportNewComponentFromFile(filePath As String) As Boolean
    Dim backupPath As String
    backupPath = filePath & "_"

    On Error GoTo ErrHandler

    FileCopy filePath, backupPath
    ConvertEncoding filePath, "UTF-8", GetSystemAnsiCharset()
    ThisWorkbook.VBProject.VBComponents.Import filePath
    Kill filePath
    Name backupPath As filePath
    ImportNewComponentFromFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If FileExists(backupPath) Then
        If FileExists(filePath) Then Kill filePath
        Name backupPath As filePath
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

Private Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
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

Sub ExportAllSheetMapsToMD()
    Dim ws As Worksheet
    Dim sheetFolderPath As String
    Dim fileName As String
    Dim mdContent As String
    
    sheetFolderPath = ThisWorkbook.Path & "\sheet"
    If Dir(sheetFolderPath, vbDirectory) = "" Then
        MkDir sheetFolderPath
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        mdContent = GenerateSheetMapMarkdown(ws)
        
        fileName = sheetFolderPath & "\" & ws.CodeName & ".md"
        
        SaveAsUTF8 fileName, mdContent
    Next
    
    MsgBox "All sheet maps exported." & vbLf & "Saved to: " & sheetFolderPath, vbInformation
End Sub

Private Function GenerateSheetMapMarkdown(ws As Worksheet) As String
    Dim rng As Range
    Dim mapText As String
    Dim role As String
    Dim cellName As String
    Dim styleParts As String
    
    mapText = "# Sheet Configuration" & vbCrLf
    mapText = mapText & "- VBA CodeName: " & ws.CodeName & vbCrLf
    mapText = mapText & "- Excel UI Name: " & ws.Name & vbCrLf & vbCrLf
    
    mapText = mapText & "| Address | Name | Value / Label | Formula | Style |" & vbCrLf
    mapText = mapText & "| :--- | :--- | :--- | :--- | :--- |" & vbCrLf
    
    For Each rng In ws.UsedRange
        If rng.value <> "" Or rng.HasFormula Or rng.Interior.ColorIndex <> xlNone Then
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
            role = IIf(styleParts = "", "Normal", styleParts)
            
            mapText = mapText & "| " & rng.Address(False, False) & _
                      " | " & cellName & _
                      " | " & EscapeCellValue(rng.value) & _
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
        shpLabel = shp.TextFrame.Characters.Text
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
                        " | " & IIf(shpStyle <> "", shpStyle, "Normal") & " |" & vbCrLf
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
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    
    st.Type = 2
    st.Charset = "UTF-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
End Sub

' Returns the ADODB.Stream charset name matching the system ANSI code page.
' VBComponents.Export / Import always use the system ANSI code page,
' so this is used when converting between UTF-8 (on disk) and ANSI (for VBE).
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

