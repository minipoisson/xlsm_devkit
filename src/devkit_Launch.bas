Attribute VB_Name = "devkit_Launch"
Option Explicit

' Optional feature: show launcher
' Requires the following files in the same VBA project:
'   devkit_frmLauncher.frm/frx - launcher dialog (shown by ShowLauncherForm)
' Entry point: ShowLauncherForm

' Opens the language-selection launcher form.
Public Sub ShowLauncherForm()
    devkit_frmLauncher.Show
End Sub

' Called via Application.OnTime from devkit_frmLauncher after the modal launcher closes.
Public Sub RunShowMoveSetupForm()
    On Error Resume Next
    Application.Run "ShowMoveSetupForm"
    Dim e As Long: e = Err.Number
    On Error GoTo 0
    If e = 1004 Then
        MsgBox t("msg.optional_module_missing", _
                 "This feature requires an optional module that is not currently loaded."), vbInformation
    ElseIf e <> 0 Then
        Err.Raise e
    End If
End Sub

Public Sub RunShowInsertDeleteForm()
    On Error Resume Next
    Application.Run "ShowInsertDeleteForm"
    Dim e As Long: e = Err.Number
    On Error GoTo 0
    If e = 1004 Then
        MsgBox t("msg.optional_module_missing", _
                 "This feature requires an optional module that is not currently loaded."), vbInformation
    ElseIf e <> 0 Then
        Err.Raise e
    End If
End Sub
