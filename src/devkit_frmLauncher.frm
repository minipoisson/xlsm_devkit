VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmLauncher
   Caption         =   "Language Settings"
   ClientHeight    =   1440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4340
   StartUpPosition =   1  '1 - CenterOwner
End
Attribute VB_Name = "devkit_frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Language-selection launcher form.
' All controls are created programmatically to avoid requiring a .frx binary.
' Requires: xlsm_devkit.bas (T, Fmt, SetLang, GetLangCode, ParseINI)

Private WithEvents m_btnOK     As MSForms.CommandButton
Private WithEvents m_btnCancel As MSForms.CommandButton
Private m_cboLang   As MSForms.ComboBox
Private m_langCodes() As String

Private Sub UserForm_Initialize()
    Me.Caption = T("frmLauncher.caption", "Language Settings")
    Me.Width = 280: Me.Height = 115

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblLang")
    lbl.Caption = T("frmLauncher.lbl_language", "Language:")
    lbl.Left = 12: lbl.Top = 16: lbl.Width = 68: lbl.Height = 18

    Set m_cboLang = Me.Controls.Add("Forms.ComboBox.1", "cboLanguage")
    m_cboLang.Left = 84: m_cboLang.Top = 14: m_cboLang.Width = 170: m_cboLang.Height = 18
    m_cboLang.Style = 2

    Set m_btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    m_btnOK.Caption = T("frmLauncher.btn_ok", "OK")
    m_btnOK.Left = 100: m_btnOK.Top = 52: m_btnOK.Width = 70: m_btnOK.Height = 24
    m_btnOK.Default = True

    Set m_btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    m_btnCancel.Caption = T("frmLauncher.btn_cancel", "Cancel")
    m_btnCancel.Left = 178: m_btnCancel.Top = 52: m_btnCancel.Width = 70: m_btnCancel.Height = 24
    m_btnCancel.Cancel = True

    If GetLangMeta("rtl") = "true" Then Me.RightToLeft = True

    PopulateLangList
End Sub

Private Sub PopulateLangList()
    Dim systemLabel As String
    systemLabel = T("frmLauncher.system_setting", "System setting")
    m_cboLang.AddItem systemLabel
    ReDim m_langCodes(0)
    m_langCodes(0) = ""

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim langDir As String: langDir = ThisWorkbook.Path & "\lang"
    If Not fso.FolderExists(langDir) Then GoTo SelectCurrent

    Dim f As Object
    For Each f In fso.GetFolder(langDir).Files
        If LCase(fso.GetExtensionName(f.Name)) = "ini" Then
            Dim code As String
            code = LCase(Left(f.Name, Len(f.Name) - 4))
            Dim dict As Object: Set dict = ParseINI(f.Path)
            Dim nativeName As String: nativeName = ""
            If Not dict Is Nothing Then
                If dict.Exists("meta.native_name") Then nativeName = dict("meta.native_name")
            End If
            If Len(nativeName) = 0 Then nativeName = code
            m_cboLang.AddItem nativeName
            ReDim Preserve m_langCodes(UBound(m_langCodes) + 1)
            m_langCodes(UBound(m_langCodes)) = code
        End If
    Next f

SelectCurrent:
    Dim currentCode As String
    currentCode = GetSetting("xlsm_devkit", "Language", "Code", "")
    If Len(currentCode) = 0 Then
        m_cboLang.ListIndex = 0
    Else
        Dim i As Long
        For i = 1 To UBound(m_langCodes)
            If m_langCodes(i) = currentCode Then
                m_cboLang.ListIndex = i: Exit For
            End If
        Next i
        If m_cboLang.ListIndex < 0 Then m_cboLang.ListIndex = 0
    End If
End Sub

Private Sub m_btnOK_Click()
    Dim idx As Long: idx = m_cboLang.ListIndex
    If idx < 0 Then idx = 0
    If idx = 0 Then
        DeleteSetting "xlsm_devkit", "Language", "Code"
        mLangCode = ""
        Set mLangCache = Nothing
        Set mEnCache = Nothing
    Else
        SetLang m_langCodes(idx)
    End If
    Unload Me
End Sub

Private Sub m_btnCancel_Click()
    Unload Me
End Sub



