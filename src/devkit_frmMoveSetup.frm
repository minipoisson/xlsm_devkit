VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmMoveSetup 
   Caption         =   "Cell Range Move - Setup"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "devkit_frmMoveSetup.frx":0000
   StartUpPosition =   1  '1 - CenterOwner
End
Attribute VB_Name = "devkit_frmMoveSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Part of the Move feature.
' Requires: devkit_Move.bas (RunMoveCapture)

Private m_selectedSrcSheet As Worksheet
Private m_selectedDstSheet As Worksheet
Private m_syncLockSrc      As Boolean
Private m_syncLockDst      As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = t("frmMoveSetup.caption", "Cell Range Move - Setup")
    btnOK.Caption = t("frmInsertDelete.btn_ok", "OK")
    btnCancel.Caption = t("frmInsertDelete.btn_cancel", "Cancel")
    If GetLangMeta("rtl") = "true" Then Me.RightToLeft = True

    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet Then
            cboSrcCodeName.AddItem ws.codeName
            cboSrcSheetName.AddItem ws.Name
            cboDstCodeName.AddItem ws.codeName
            cboDstSheetName.AddItem ws.Name
        End If
    Next ws
    btnOK.Enabled = False
End Sub

Private Sub cboSrcCodeName_Change()
    If m_syncLockSrc Then Exit Sub
    m_syncLockSrc = True

    Dim i As Long
    For i = 0 To cboSrcCodeName.ListCount - 1
        If cboSrcCodeName.List(i) = cboSrcCodeName.text Then
            cboSrcSheetName.ListIndex = i
            Exit For
        End If
    Next i

    Set m_selectedSrcSheet = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.codeName = cboSrcCodeName.text Then
            Set m_selectedSrcSheet = ws
            Exit For
        End If
    Next ws

    ' Auto-set destination to same sheet if destination not yet chosen
    If m_selectedDstSheet Is Nothing And Not m_selectedSrcSheet Is Nothing Then
        Call SetDstSheet(m_selectedSrcSheet)
    End If

    m_syncLockSrc = False
    UpdateBtnOK
End Sub

Private Sub cboSrcSheetName_Change()
    If m_syncLockSrc Then Exit Sub
    m_syncLockSrc = True

    Dim i As Long
    For i = 0 To cboSrcSheetName.ListCount - 1
        If cboSrcSheetName.List(i) = cboSrcSheetName.text Then
            cboSrcCodeName.ListIndex = i
            Exit For
        End If
    Next i

    Set m_selectedSrcSheet = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.Name = cboSrcSheetName.text Then
            Set m_selectedSrcSheet = ws
            Exit For
        End If
    Next ws

    If m_selectedDstSheet Is Nothing And Not m_selectedSrcSheet Is Nothing Then
        Call SetDstSheet(m_selectedSrcSheet)
    End If

    m_syncLockSrc = False
    UpdateBtnOK
End Sub

Private Sub cboDstCodeName_Change()
    If m_syncLockDst Then Exit Sub
    m_syncLockDst = True

    Dim i As Long
    For i = 0 To cboDstCodeName.ListCount - 1
        If cboDstCodeName.List(i) = cboDstCodeName.text Then
            cboDstSheetName.ListIndex = i
            Exit For
        End If
    Next i

    Set m_selectedDstSheet = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.codeName = cboDstCodeName.text Then
            Set m_selectedDstSheet = ws
            Exit For
        End If
    Next ws

    m_syncLockDst = False
    UpdateBtnOK
End Sub

Private Sub cboDstSheetName_Change()
    If m_syncLockDst Then Exit Sub
    m_syncLockDst = True

    Dim i As Long
    For i = 0 To cboDstSheetName.ListCount - 1
        If cboDstSheetName.List(i) = cboDstSheetName.text Then
            cboDstCodeName.ListIndex = i
            Exit For
        End If
    Next i

    Set m_selectedDstSheet = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet And ws.Name = cboDstSheetName.text Then
            Set m_selectedDstSheet = ws
            Exit For
        End If
    Next ws

    m_syncLockDst = False
    UpdateBtnOK
End Sub

Private Sub SetDstSheet(ws As Worksheet)
    m_syncLockDst = True
    Dim i As Long
    For i = 0 To cboDstCodeName.ListCount - 1
        If cboDstCodeName.List(i) = ws.codeName Then
            cboDstCodeName.ListIndex = i
            cboDstSheetName.ListIndex = i
            Exit For
        End If
    Next i
    Set m_selectedDstSheet = ws
    m_syncLockDst = False
End Sub

Private Sub btnOK_Click()
    Dim src As Worksheet: Set src = m_selectedSrcSheet
    Dim dst As Worksheet
    If m_selectedDstSheet Is Nothing Then
        Set dst = src
    Else
        Set dst = m_selectedDstSheet
    End If
    Unload Me
    RunMoveCapture src, dst, True
End Sub

Private Sub btnCancel_Click()
    Dim isDirty As Boolean
    isDirty = (cboSrcCodeName.ListIndex >= 0) Or _
              (cboDstCodeName.ListIndex >= 0)
    If isDirty Then
        If MsgBox(t("frmMoveSetup.discard_confirm", "Discard changes and close?"), vbOKCancel) = vbCancel Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub UpdateBtnOK()
    btnOK.Enabled = Not (m_selectedSrcSheet Is Nothing)
End Sub



