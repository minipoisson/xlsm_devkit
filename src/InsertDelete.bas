Attribute VB_Name = "InsertDelete"
Option Explicit

Public Sub RunInsertDelete( _
    ws As Worksheet, _
    startPos As String, _
    lineCount As Long, _
    isInsert As Boolean, _
    showResultDialog As Boolean _
)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sheetFolder As String
    sheetFolder = ThisWorkbook.Path & "\sheet"
    If Not fso.FolderExists(sheetFolder) Then fso.CreateFolder sheetFolder

    Dim oldMdPath As String
    oldMdPath = sheetFolder & "\" & ws.codeName & ".old.md"
    Dim mdPath As String
    mdPath = sheetFolder & "\" & ws.codeName & ".md"

    Dim isRow As Boolean
    isRow = IsRowPosition(startPos)
    Dim rowNum As Long
    Dim colNum As Long
    Dim colLetter As String
    If isRow Then
        rowNum = CLng(startPos)
    Else
        colLetter = UCase(startPos)
        colNum = ColLetterToNum(colLetter)
    End If

    ' Step 1: .old.md existence check
    If fso.FileExists(oldMdPath) Then
        If MsgBox("sheet\" & ws.codeName & ".old.md already exists." & vbLf & _
                  "Overwrite and continue?", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If

    ' Step 2: Export current sheet to .old.md
    ExportSheetToMDFile ws, oldMdPath

    ' Step 3: Perform Excel operation
    On Error GoTo ErrHandler3
    If isInsert Then
        If isRow Then
            ws.Rows(rowNum & ":" & (rowNum + lineCount - 1)).Insert Shift:=xlDown
        Else
            ws.Columns(ColNumToLetter(colNum) & ":" & ColNumToLetter(colNum + lineCount - 1)).Insert Shift:=xlToRight
        End If
    Else
        If isRow Then
            If rowNum + lineCount - 1 > 1048576 Then
                Err.Raise vbObjectError, , "Row range exceeds the sheet limit (1,048,576)."
            End If
            ws.Rows(rowNum & ":" & (rowNum + lineCount - 1)).Delete Shift:=xlUp
        Else
            If colNum + lineCount - 1 > 16384 Then
                Err.Raise vbObjectError, , "Column range exceeds the sheet limit (16,384)."
            End If
            ws.Columns(ColNumToLetter(colNum) & ":" & ColNumToLetter(colNum + lineCount - 1)).Delete Shift:=xlToLeft
        End If
    End If
    On Error GoTo 0

    ' Step 4: Export updated sheet to .md
    On Error GoTo ErrHandler4
    ExportSheetToMDFile ws, mdPath
    On Error GoTo 0

    ' Step 5-6: Build instruction and copy to clipboard
    Dim instructionText As String
    instructionText = BuildInstructionText(ws, lineCount, isInsert, isRow, rowNum, colNum, colLetter)
    CopyToClipboard instructionText

    ' Step 7: Show frmInstruction if requested
    If showResultDialog Then
        frmInstruction.txtInstruction.text = instructionText
        frmInstruction.Show
    End If
    Exit Sub

ErrHandler3:
    Dim errMsg3 As String
    errMsg3 = Err.Description
    On Error Resume Next
    If fso.FileExists(oldMdPath) Then fso.DeleteFile oldMdPath, True
    On Error GoTo 0
    MsgBox "Operation failed: " & errMsg3, vbExclamation
    Exit Sub

ErrHandler4:
    Dim errMsg4 As String
    errMsg4 = Err.Description
    On Error GoTo 0
    MsgBox "Warning: The Excel operation succeeded but the sheet export failed." & vbLf & _
           "The .old.md file has been retained.", vbExclamation
    Exit Sub
End Sub

Private Function IsRowPosition(pos As String) As Boolean
    If Len(pos) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(pos)
        If Mid(pos, i, 1) < "0" Or Mid(pos, i, 1) > "9" Then Exit Function
    Next i
    IsRowPosition = True
End Function

Private Function ColLetterToNum(col As String) As Long
    Dim n As Long
    Dim i As Long
    For i = 1 To Len(col)
        n = n * 26 + (Asc(UCase(Mid(col, i, 1))) - 64)
    Next i
    ColLetterToNum = n
End Function

Private Function ColNumToLetter(colNum As Long) As String
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

Private Sub CopyToClipboard(text As String)
    Dim obj As Object
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.SetText text
    obj.PutInClipboard
End Sub

Private Function BuildInstructionText( _
    ws As Worksheet, _
    lineCount As Long, _
    isInsert As Boolean, _
    isRow As Boolean, _
    rowNum As Long, _
    colNum As Long, _
    colLetter As String _
) As String
    Dim sheetName As String: sheetName = ws.Name
    Dim codeName As String: codeName = ws.codeName
    Dim WARN As String: WARN = ChrW(9888) & " "
    Dim t As String

    If isInsert And isRow Then
        t = "The following operation was performed on Excel sheet """ & sheetName & """ (VBA CodeName: " & codeName & "):" & vbCrLf & vbCrLf
        t = t & "  Inserted " & lineCount & " row(s) before row " & rowNum & "." & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any cell references" & vbCrLf
        t = t & "in the VBA code that refer to this sheet and have shifted due to this operation." & vbCrLf & vbCrLf
        t = t & "Sheet map before: @sheet/" & codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after:  @sheet/" & codeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Range(""B5""), Range(""A1:C10""), etc." & vbCrLf
        t = t & "- Numeric-form: Rows(N), Cells(R, C), etc." & vbCrLf
        t = t & "- Dynamic-form: Range(""A"" & n), Cells(startRow + 1, col), etc." & vbCrLf & vbCrLf
        t = t & "Please list any references that cannot be updated automatically or require confirmation."

    ElseIf Not isInsert And isRow Then
        Dim endRow As Long: endRow = rowNum + lineCount - 1
        t = "The following operation was performed on Excel sheet """ & sheetName & """ (VBA CodeName: " & codeName & "):" & vbCrLf & vbCrLf
        t = t & "  Deleted " & lineCount & " row(s) starting at row " & rowNum & "  (rows " & rowNum & " through " & endRow & ")." & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any cell references" & vbCrLf
        t = t & "in the VBA code that refer to this sheet and have shifted due to this operation." & vbCrLf & vbCrLf
        t = t & "Sheet map before: @sheet/" & codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after:  @sheet/" & codeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Range(""B5""), Range(""A1:C10""), etc." & vbCrLf
        t = t & "- Numeric-form: Rows(N), Cells(R, C), etc." & vbCrLf
        t = t & "- Dynamic-form: Range(""A"" & n), Cells(startRow + 1, col), etc." & vbCrLf & vbCrLf
        t = t & WARN & "References that directly targeted the deleted rows (" & rowNum & " through " & endRow & ")" & vbCrLf
        t = t & "  are now invalid. Please list these as warnings."

    ElseIf isInsert And Not isRow Then
        t = "The following operation was performed on Excel sheet """ & sheetName & """ (VBA CodeName: " & codeName & "):" & vbCrLf & vbCrLf
        t = t & "  Inserted " & lineCount & " column(s) before column " & colLetter & " (column number " & colNum & ")." & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any cell references" & vbCrLf
        t = t & "in the VBA code that refer to this sheet and have shifted due to this operation." & vbCrLf & vbCrLf
        t = t & "Sheet map before: @sheet/" & codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after:  @sheet/" & codeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Range(""C5""), Range(""A1:C10""), etc." & vbCrLf
        t = t & "- Numeric-form: Columns(N), Cells(R, C), etc." & vbCrLf
        t = t & "- Dynamic-form: Range(colLetter & n), Cells(r, colNum + 1), etc." & vbCrLf & vbCrLf
        t = t & "Please list any references that cannot be updated automatically or require confirmation."

    Else
        Dim endColNum As Long: endColNum = colNum + lineCount - 1
        Dim endColLetter As String: endColLetter = ColNumToLetter(endColNum)
        t = "The following operation was performed on Excel sheet """ & sheetName & """ (VBA CodeName: " & codeName & "):" & vbCrLf & vbCrLf
        t = t & "  Deleted " & lineCount & " column(s) starting at column " & colLetter & " (column number " & colNum & ")" & vbCrLf
        t = t & "  through column " & endColLetter & " (column number " & endColNum & ")." & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any cell references" & vbCrLf
        t = t & "in the VBA code that refer to this sheet and have shifted due to this operation." & vbCrLf & vbCrLf
        t = t & "Sheet map before: @sheet/" & codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after:  @sheet/" & codeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Range(""C5""), Range(""A1:C10""), etc." & vbCrLf
        t = t & "- Numeric-form: Columns(N), Cells(R, C), etc." & vbCrLf
        t = t & "- Dynamic-form: Range(colLetter & n), Cells(r, colNum + 1), etc." & vbCrLf & vbCrLf
        t = t & WARN & "References that directly targeted the deleted columns" & vbCrLf
        t = t & "  (" & colLetter & "(col " & colNum & ") through " & endColLetter & "(col " & endColNum & ")) are now invalid." & vbCrLf
        t = t & "  Please list these as warnings."
    End If

    BuildInstructionText = t
End Function

