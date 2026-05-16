VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInstruction 
   Caption         =   "Update VBA References"
   ClientHeight    =   6640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmInstruction.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCopy_Click()
    Dim obj As MSForms.DataObject
    Set obj = New MSForms.DataObject
    obj.SetText txtInstruction.text
    obj.PutInClipboard
End Sub

Private Sub btnImport_Click()
    If MsgBox("Import all .bas files from src/ into this workbook?" & vbLf & _
              "Make sure all changes are saved in VS Code first.", _
              vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrHandler
    ImportAllComponents True
    If MsgBox("Code has been imported.", vbOKOnly + vbInformation) = vbOK Then
        Unload Me
    End If
    Exit Sub

ErrHandler:
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

