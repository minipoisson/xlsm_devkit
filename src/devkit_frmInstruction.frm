VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmInstruction 
   Caption         =   "Update VBA References"
   ClientHeight    =   6640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "devkit_frmInstruction.frx":0000
   StartUpPosition =   1  '1 - CenterOwner
End
Attribute VB_Name = "devkit_frmInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Shared result/import dialog used by InsertDelete and Move features.
' Requires: xlsm_devkit.bas (ImportAllComponents)

Private Sub UserForm_Initialize()
    Me.Caption = t("frmInstruction.caption", "Update VBA References")
    btnCopy.Caption = t("frmInstruction.btn_copy", "Copy")
    btnImport.Caption = t("frmInstruction.btn_import", "Import")
    btnClose.Caption = t("frmInstruction.btn_close", "Close")
    If GetLangMeta("rtl") = "true" Then Me.RightToLeft = True
End Sub

Private Sub btnCopy_Click()
    Dim obj As MSForms.DataObject
    Set obj = New MSForms.DataObject
    obj.SetText txtInstruction.text
    obj.PutInClipboard
End Sub

Private Sub btnImport_Click()
    If MsgBox(t("frmInstruction.import_confirm", _
                "Import all .bas files from src/ into this workbook?" & vbLf & _
                "Make sure all changes are saved in VS Code first."), _
              vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrHandler
    ImportAllComponents True
    If MsgBox(t("frmInstruction.import_done", "Code has been imported."), vbOKOnly + vbInformation) = vbOK Then
        Unload Me
    End If
    Exit Sub

ErrHandler:
    MsgBox Fmt(t("frmInstruction.import_failed", "Import failed: {0}"), Err.Description), vbExclamation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
