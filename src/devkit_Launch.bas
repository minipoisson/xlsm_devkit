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

