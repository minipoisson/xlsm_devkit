VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} devkit_frmLauncher 
   Caption         =   "Launcher"
   ClientHeight    =   6240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "devkit_frmLauncher.frx":0000
   StartUpPosition =   1  '1 - CenterOwner
End
Attribute VB_Name = "devkit_frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_langCodes() As String
Private m_populating As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = t("frmLauncher.caption", "Language Settings")
    lbl_language.Caption = t("frmLauncher.lbl_language", "Language:")
    btnClose.Caption = t("frmLauncher.btn_close", "Close")
    frameImport.Caption = t("frmLauncher.import", "Import")
    frameExport.Caption = t("frmLauncher.export", "Export")
    Dim strAll As String, strMF As String, strS As String
    strAll = t("frmLauncher.all_modules_forms_sheets", "All modules, forms and sheets")
    strMF = t("frmLauncher.all_modules_forms", "All modules and forms")
    strS = t("frmLauncher.all_sheets", "All sheets")
    btnExpAll.Caption = strAll
    btnExpAllModulesForms.Caption = strMF
    btnExpAllSheets.Caption = strS
    btnImpAll.Caption = strAll
    btnImpAllModulesForms.Caption = strMF
    btnImpAllSheets.Caption = strS
    
    btnInsertDelete.Caption = t("frmLauncher.btn_insert_delete", "Insert / Delete")
    btnMoveSetup.Caption = t("frmLauncher.btn_move_setup", "Move Setup")
    btnSaveAsRelease.Caption = t("frmLauncher.btn_save_as_release", "Save as Release")
    btnSaveAsRelease.Enabled = (UCase(Left(ThisWorkbook.Name, 4)) = "DEV_")

    If GetLangMeta("rtl") = "true" Then Me.RightToLeft = True

    PopulateLangList
End Sub

Private Sub PopulateLangList()
    m_populating = True
    Dim systemLabel As String
    systemLabel = t("frmLauncher.system_setting", "System setting")
    cboLang.Clear
    cboLang.AddItem systemLabel
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
            cboLang.AddItem nativeName
            ReDim Preserve m_langCodes(UBound(m_langCodes) + 1)
            m_langCodes(UBound(m_langCodes)) = code
        End If
    Next f

SelectCurrent:
    Dim currentCode As String
    currentCode = GetSetting("xlsm_devkit", "Language", "Code", "")
    If Len(currentCode) = 0 Then
        cboLang.ListIndex = 0
    Else
        Dim i As Long
        For i = 1 To UBound(m_langCodes)
            If m_langCodes(i) = currentCode Then
                cboLang.ListIndex = i: Exit For
            End If
        Next i
        If cboLang.ListIndex < 0 Then cboLang.ListIndex = 0
    End If
    m_populating = False
End Sub
Private Sub cboLang_Change()
    If m_populating Then Exit Sub
    Dim idx As Long: idx = cboLang.ListIndex
    If idx < 0 Then idx = 0
    If idx = 0 Then
        SetLang ""
    Else
        SetLang m_langCodes(idx)
    End If
    UserForm_Initialize
End Sub

Private Sub btnExpAll_Click()
    Application.OnTime Now, "ExportAllModulesFormsSheetMaps"
    Unload Me
End Sub

Private Sub btnExpAllModulesForms_Click()
    Application.OnTime Now, "CallExportAllComponents"
    Unload Me
End Sub

Private Sub btnExpAllSheets_Click()
    Application.OnTime Now, "CallExportAllSheetMapsToMD"
    Unload Me
End Sub

Private Sub btnImpAll_Click()
    Application.OnTime Now, "ImportAllModulesFormsSheetMaps"
    Unload Me
End Sub

Private Sub btnImpAllModulesForms_Click()
    Application.OnTime Now, "CallImportAllComponents"
    Unload Me
End Sub

Private Sub btnImpAllSheets_Click()
    Application.OnTime Now, "CallImportAllSheetMapsFromMD"
    Unload Me
End Sub

Private Sub btnInsertDelete_Click()
    Application.OnTime Now, "RunShowInsertDeleteForm"
    Unload Me
End Sub

Private Sub btnMoveSetup_Click()
    Application.OnTime Now, "RunShowMoveSetupForm"
    Unload Me
End Sub

Private Sub btnSaveAsRelease_Click()
    Application.OnTime Now, "CallSaveAsRelease"
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
