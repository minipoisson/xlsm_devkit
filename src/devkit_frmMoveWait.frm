VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmMoveWait 
   Caption         =   "Cell Range Move - Recording"
   ClientHeight    =   3140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "devkit_frmMoveWait.frx":0000
   StartUpPosition =   1  '1 - CenterOwner
End
Attribute VB_Name = "devkit_frmMoveWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Part of the Move feature.
' Requires: devkit_Move.bas (ExecuteMoveCapturePhase2, CancelMoveCapture)

Private Sub UserForm_Initialize()
    Me.Caption = t("frmMoveWait.caption", "Cell Range Move - Recording")
    btnStop.Caption = t("frmMoveWait.btn_stop", "Stop")
    btnCancel.Caption = t("frmMoveWait.btn_cancel", "Cancel")
    If GetLangMeta("rtl") = "true" Then Me.RightToLeft = True

    lblInstruction.Caption = _
        t("frmMoveWait.lbl_step1", "1. Start macro recording: Developer tab -> Record Macro.") & vbLf & _
        t("frmMoveWait.lbl_step1b", "   Set 'Store macro in: This Workbook', then click OK.") & vbLf & _
        t("frmMoveWait.lbl_step2", "2. Move cells using Cut & Paste only (do not drag-and-drop).") & vbLf & _
        t("frmMoveWait.lbl_step3", "3. Stop recording: Developer tab -> Stop Recording.") & vbLf & _
        t("frmMoveWait.lbl_step4", "4. Click Stop below.")
    lblReminder.Caption = t("frmMoveWait.lbl_reminder", "[WARNING] Do not use drag-and-drop. Use Cut (Ctrl+X) then Paste (Ctrl+V or Enter).")
End Sub

Private Sub btnStop_Click()
    Me.Hide
    ExecuteMoveCapturePhase2
    Unload Me
End Sub

Private Sub btnCancel_Click()
    If MsgBox(t("frmMoveWait.stop_reminder", "Stop recording first (Developer tab -> Stop Recording) if still active.") & vbLf & vbLf & _
              t("frmMoveWait.cancel_confirm", "Cancel the cell move operation? Snapshot files will be deleted."), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    CancelMoveCapture
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub

