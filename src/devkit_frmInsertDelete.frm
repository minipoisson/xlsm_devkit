VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmInsertDelete
   Caption         =   "Insert/Delete Rows/Columns"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "devkit_frmInsertDelete.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "devkit_frmInsertDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Part of the InsertDelete feature.
' Requires: devkit_InsertDelete.bas (RunInsertDelete)

Private m_selectedSheet As Worksheet
Private m_isValidN      As Boolean
Private m_isRowMode     As Boolean
Private m_colNum        As Long
Private m_syncLock      As Boolean

Private Sub UserForm_Initialize()
    Dim ws As Object
    spnCount.Value = 1

    m_isValidN = False
    btnOK.Enabled = False
    UpdateLabels
End Sub

Private Sub cboCodeName_Change()
    If m_syncLock Then Exit Sub
    m_syncLock = True

    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.codeName = cboCodeName.text Then
            Set m_selectedSheet = ws
            Exit For
        End If
    Next ws

    m_syncLock = False
    UpdateBtnOK
End Sub

Private Sub cboSheetName_Change()
    If m_syncLock Then Exit Sub
    m_syncLock = True

    Set m_selectedSheet = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.Name = cboSheetName.text Then
            Set m_selectedSheet = ws
            Exit For
        End If
    Next ws

    m_syncLock = False
    UpdateBtnOK
End Sub

Private Sub txtN_AfterUpdate()
    Dim pos As String
    pos = Trim(txtN.text)
    m_isValidN = False
    m_isRowMode = False
    m_colNum = 0

    If Len(pos) = 0 Then
        ' nothing entered
    ElseIf IsAllDigits(pos) Then
        Dim rn As Long
        rn = CLng(pos)
        If rn >= 1 And rn <= 1048576 Then
            m_isValidN = True
            m_isRowMode = True
        End If
    ElseIf IsAllLetters(pos) Then
        Dim cn As Long
        cn = ColLetterToNum(UCase(pos))
        If cn >= 1 And cn <= 16384 Then
            m_isValidN = True
            m_isRowMode = False
            m_colNum = cn
        End If
    End If

    UpdateLabels
    UpdateBtnOK
End Sub

Private Function IsAllDigits(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsAllLetters(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid(s, i, 1) Like "[A-Za-z]") Then Exit Function
    Next i
    IsAllLetters = True
End Function

Private Sub txtCount_Change()
    Dim s As String
    s = txtCount.text
    If s Like String(Len(s), "#") And Len(s) > 0 Then
        Dim v As Long
        v = CLng(s)
        If v >= 1 And v <= 9999 Then
            spnCount.Value = v
        End If
    End If
    UpdateLabels
End Sub

Private Sub spnCount_Change()
    txtCount.text = CStr(spnCount.Value)
    UpdateLabels
End Sub

Private Sub btnOK_Click()
    Dim lineCount As Long
    Dim s As String
    s = txtCount.text
    If s Like String(Len(s), "#") And Len(s) > 0 Then
        lineCount = CLng(s)
    Else
        lineCount = 1
    End If

    Dim storedPos As String: storedPos = txtN.text
    Dim storedInsert As Boolean: storedInsert = optInsert.Value
    Dim storedSheet As Worksheet: Set storedSheet = m_selectedSheet

    Unload Me
    RunInsertDelete storedSheet, storedPos, lineCount, storedInsert, True
End Sub

Private Sub btnCancel_Click()
    Dim isDirty As Boolean
    isDirty = (cboCodeName.ListIndex >= 0) Or _
              (Len(Trim(txtN.text)) > 0) Or _
              (txtCount.text <> "1") Or _
              (optDelete.Value = True)

    If isDirty Then
        If MsgBox("Discard changes and close?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub UpdateLabels()
    Dim pos As String
    pos = Trim(txtN.text)

    If Len(pos) = 0 Then
        lblRecognized.Caption = "(enter a row number or column letter)"
    ElseIf m_isValidN And m_isRowMode Then
        lblRecognized.Caption = ChrW(8594) & " Row " & pos
    ElseIf m_isValidN And Not m_isRowMode Then
        lblRecognized.Caption = ChrW(8594) & " Column " & UCase(pos) & "  (col " & m_colNum & ")"
    Else
        lblRecognized.Caption = ChrW(8594) & " Invalid input"
    End If

    If Not m_isValidN Or Len(pos) = 0 Then
        lblAction.Caption = ""
        Exit Sub
    End If

    Dim cnt As Long
    Dim s As String
    s = txtCount.text
    If s Like String(Len(s), "#") And Len(s) > 0 Then
        cnt = CLng(s)
    Else
        cnt = 1
    End If

    Dim rowWord As String
    Dim colWord As String
    rowWord = IIf(cnt = 1, "row", "rows")
    colWord = IIf(cnt = 1, "column", "columns")

    Dim action As String
    If m_isRowMode Then
        If optInsert.Value Then
            action = "Insert " & cnt & " " & rowWord & " before row " & pos
        Else
            action = "Delete " & cnt & " " & rowWord & " at row " & pos
        End If
    Else
        If optInsert.Value Then
            action = "Insert " & cnt & " " & colWord & " before column " & UCase(pos) & " (col " & m_colNum & ")"
        Else
            action = "Delete " & cnt & " " & colWord & " starting at column " & UCase(pos) & " (col " & m_colNum & ")"
        End If
    End If
    lblAction.Caption = action
End Sub

Private Sub UpdateBtnOK()
    btnOK.Enabled = (Not m_selectedSheet Is Nothing) And m_isValidN
End Sub



