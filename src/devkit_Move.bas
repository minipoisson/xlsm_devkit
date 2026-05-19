Attribute VB_Name = "devkit_Move"
Option Explicit

Private Const APP_KEY  As String = "xlsm_devkit"
Private Const REG_SECT As String = "MoveCapture"
Private Const REG_KEY  As String = "State"

Private mWsSource         As Worksheet
Private mWsDest           As Worksheet
Private mShowResultDialog As Boolean
Private mSnapshot         As Collection
Private mNewComponentName As String

' Optional feature: record and document cell-range move operations.
' Requires the following files in the same VBA project:
'   xlsm_devkit.bas        - CallExportAllComponents, ExportSheetToMDFile
'   devkit_frmMoveSetup.frm/frx   - setup dialog (shown by ShowMoveSetupForm)
'   devkit_frmMoveWait.frm/frx    - recording dialog (shown by RunMoveCapture)
'   devkit_frmInstruction.frm/frx - result/import dialog (shown by ExecuteMoveCapturePhase2)
' Entry point: ShowMoveSetupForm

Public Sub ShowMoveSetupForm()
    If MsgBox(t("msg.export_components_confirm", "Export all modules and forms to src/?"), _
              vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    CallExportAllComponents True
    devkit_frmMoveSetup.Show
End Sub

' ============================================================
' Phase 1
' ============================================================

Public Sub RunMoveCapture( _
    wsSource As Worksheet, _
    wsDest As Worksheet, _
    showResultDialog As Boolean _
)
    Set mWsSource = wsSource
    Set mWsDest = wsDest
    mShowResultDialog = showResultDialog
    mNewComponentName = ""

    Set mSnapshot = TakeSnapshot()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sheetFolder As String
    sheetFolder = ThisWorkbook.Path & "\sheet"
    If Not fso.FolderExists(sheetFolder) Then fso.CreateFolder sheetFolder

    ' Check for existing .old.md files
    If Not CheckAndConfirmOldMd(fso, wsSource, sheetFolder) Then Exit Sub
    If Not wsSource Is wsDest Then
        If Not CheckAndConfirmOldMd(fso, wsDest, sheetFolder) Then Exit Sub
    End If

    ' Export .old.md files
    On Error GoTo ErrExportOld
    ExportSheetToMDFile wsSource, sheetFolder & "\" & wsSource.codeName & ".old.md"
    If Not wsSource Is wsDest Then
        ExportSheetToMDFile wsDest, sheetFolder & "\" & wsDest.codeName & ".old.md"
    End If
    On Error GoTo 0

    ' Persist state to registry before showing form.
    ' Manual macro recording resets the VBA runtime, clearing module-level vars.
    SaveState
    devkit_frmMoveWait.Show vbModeless
    Exit Sub

ErrExportOld:
    MsgBox Fmt(t("msg.snapshot_export_failed", "Failed to export sheet snapshot: {0}"), Err.Description), vbExclamation
End Sub

Private Function CheckAndConfirmOldMd(fso As Object, ws As Worksheet, sheetFolder As String) As Boolean
    Dim p As String
    p = sheetFolder & "\" & ws.codeName & ".old.md"
    If fso.FileExists(p) Then
        If MsgBox(Fmt(t("msg.overwrite_old_md", "sheet\{0}.old.md already exists. Overwrite and continue?"), ws.codeName), _
                  vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    End If
    CheckAndConfirmOldMd = True
End Function

' ============================================================
' Phase 2 ? triggered by btnStop in frmMoveWait
' ============================================================

Public Sub ExecuteMoveCapturePhase2()
    If mWsSource Is Nothing Or mSnapshot Is Nothing Then
        If Not LoadState() Then
            MsgBox t("msg.state_lost", "Operation state was lost. Please restart from ShowMoveSetupForm."), vbExclamation
            Exit Sub
        End If
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sheetFolder As String
    sheetFolder = ThisWorkbook.Path & "\sheet"

    ' Identify recorded code
    Dim newModName As String
    Dim subCode As String
    Dim startLine As Long
    Dim addedLineCount As Long
    If Not IdentifyRecordedCode(newModName, subCode, startLine, addedLineCount) Then
        MsgBox t("msg.no_macro_found", "No recorded macro was found in the VBA project. The operation will be cancelled."), vbExclamation
        Call CleanupOldMdFiles(fso, sheetFolder)
        Exit Sub
    End If
    mNewComponentName = newModName

    ' Parse the recorded macro for Cut operations
    Dim srcSheetName As String, srcRange As String
    Dim dstSheetName As String, dstRange As String
    Dim unexpectedLines As String
    If Not ParseCutOperation(subCode, srcSheetName, srcRange, dstSheetName, dstRange, unexpectedLines) Then
        MsgBox Fmt(t("msg.no_move_detected", "No cell move was detected in the recorded macro. The operation will be cancelled." & vbLf & vbLf & "Recorded code (first 500 chars):" & vbLf & "{0}"), Left(subCode, 500)), vbExclamation
        Call CleanupOldMdFiles(fso, sheetFolder)
        Exit Sub
    End If

    ' Resolve default sheet names
    If srcSheetName = "" Then srcSheetName = mWsSource.Name
    If dstSheetName = "" Then dstSheetName = mWsDest.Name

    ' Check cross-sheet discrepancy
    Dim reducedAccuracy As Boolean
    reducedAccuracy = False
    If dstSheetName <> mWsDest.Name Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox(Fmt(t("msg.wrong_sheet_confirm", _
                          "The recorded move targets sheet '{0}', but '{1}' was selected in the setup dialog." & vbLf & _
                          "The .old.md for '{0}' was not saved before the move." & vbLf & vbLf & _
                          "Continue using the current state (reduced accuracy)?"), _
                         dstSheetName, mWsDest.Name), _
                     vbYesNo + vbDefaultButton2)
        If ans = vbNo Then
            Call CleanupOldMdFiles(fso, sheetFolder)
            Exit Sub
        End If
        reducedAccuracy = True
    End If

    ' Export post-move .md files
    On Error GoTo ErrExportMd
    ExportSheetToMDFile mWsSource, sheetFolder & "\" & mWsSource.codeName & ".md"
    If Not mWsSource Is mWsDest Then
        ExportSheetToMDFile mWsDest, sheetFolder & "\" & mWsDest.codeName & ".md"
    End If
    On Error GoTo 0

    ' Build instruction text, copy to clipboard, show dialog
    Dim isSameSheet As Boolean
    isSameSheet = (mWsSource Is mWsDest)
    Dim instructionText As String
    instructionText = BuildMoveInstructionText( _
        isSameSheet, srcSheetName, srcRange, dstSheetName, dstRange, _
        mWsSource, mWsDest, reducedAccuracy, unexpectedLines)
    Dim modDeleteNote As String
    If newModName <> "" Then
        modDeleteNote = ChrW(9888) & " The recorded macro module """ & newModName & _
                        """ is still in the VBA project." & vbCrLf & _
                        "  Please delete it from the VBE (Alt+F11) when done."
    Else
        modDeleteNote = ChrW(9888) & " Recorded macro lines were appended to an existing module." & vbCrLf & _
                        "  Please delete the recorded lines from the VBE (Alt+F11) when done."
    End If
    instructionText = instructionText & vbCrLf & vbCrLf & modDeleteNote
    CopyToClipboard instructionText
    ClearState
    If mShowResultDialog Then
        devkit_frmInstruction.txtInstruction.text = instructionText
        devkit_frmInstruction.Show
    End If
    Exit Sub

ErrExportMd:
    ClearState
    MsgBox t("msg.move_export_warning", "[WARNING] The cell move succeeded but the sheet export failed. The .old.md file(s) have been retained."), vbExclamation
End Sub

' ============================================================
' Cancel cleanup ? triggered by btnCancel in frmMoveWait
' ============================================================

Public Sub CancelMoveCapture()
    If mWsSource Is Nothing Or mSnapshot Is Nothing Then LoadState
    If mWsSource Is Nothing Then
        ClearState
        Exit Sub
    End If

    Dim newModName As String
    Dim subCode As String
    Dim startLine As Long
    Dim addedLineCount As Long
    IdentifyRecordedCode newModName, subCode, startLine, addedLineCount
    Call RemoveRecordedCode(newModName, startLine, addedLineCount)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sheetFolder As String
    sheetFolder = ThisWorkbook.Path & "\sheet"
    Call CleanupOldMdFiles(fso, sheetFolder)
    ClearState
End Sub

' ============================================================
' Private helpers
' ============================================================

Private Function TakeSnapshot() As Collection
    Dim snap As New Collection
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Dim lineCount As Long
        lineCount = comp.CodeModule.CountOfLines
        Dim codeText As String
        If lineCount > 0 Then
            codeText = comp.CodeModule.lines(1, lineCount)
        End If
        Dim item(3) As Variant
        item(0) = comp.Name
        item(1) = comp.Type
        item(2) = codeText
        item(3) = lineCount
        snap.Add item
    Next comp
    Set TakeSnapshot = snap
End Function

Private Function IdentifyRecordedCode( _
    ByRef newModName As String, _
    ByRef subCode As String, _
    ByRef startLine As Long, _
    ByRef addedLineCount As Long _
) As Boolean
    newModName = ""
    subCode = ""
    startLine = 0
    addedLineCount = 0

    ' Build lookup of snapshot names
    Dim snapNames As Object
    Set snapNames = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To mSnapshot.count
        Dim snapItem As Variant
        snapItem = mSnapshot(i)
        snapNames(snapItem(0)) = i
    Next i

    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If Not snapNames.Exists(comp.Name) Then
            ' Case A: brand new component added by recording
            newModName = comp.Name
            mNewComponentName = newModName
            Dim lc As Long
            lc = comp.CodeModule.CountOfLines
            If lc > 0 Then subCode = comp.CodeModule.lines(1, lc)
            IdentifyRecordedCode = True
            startLine = 0
            addedLineCount = 0
            Exit Function
        End If
    Next comp

    ' Case B: existing component had lines added
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If snapNames.Exists(comp.Name) Then
            Dim idx As Long
            idx = snapNames(comp.Name)
            Dim si As Variant
            si = mSnapshot(idx)
            Dim oldCount As Long
            oldCount = si(3)
            Dim newCount As Long
            newCount = comp.CodeModule.CountOfLines
            If newCount > oldCount Then
                ' Extract added lines (appended at end)
                addedLineCount = newCount - oldCount
                startLine = oldCount + 1
                subCode = comp.CodeModule.lines(startLine, addedLineCount)
                IdentifyRecordedCode = True
                newModName = ""
                Exit Function
            End If
        End If
    Next comp
End Function

Private Sub RemoveRecordedCode(newModName As String, startLine As Long, addedLineCount As Long)
    On Error Resume Next
    If newModName <> "" Then
        Dim comp As Object
        Set comp = ThisWorkbook.VBProject.VBComponents(newModName)
        If Not comp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove comp
        End If
    ElseIf startLine > 0 And addedLineCount > 0 Then
        ' Case B: find the module whose lines changed and delete the added lines
        Dim i As Long
        For i = 1 To mSnapshot.count
            Dim si As Variant
            si = mSnapshot(i)
            Dim c As Object
            On Error Resume Next
            Set c = ThisWorkbook.VBProject.VBComponents(si(0))
            On Error GoTo 0
            If Not c Is Nothing Then
                If c.CodeModule.CountOfLines >= startLine + addedLineCount - 1 Then
                    c.CodeModule.DeleteLines startLine, addedLineCount
                    Exit For
                End If
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub CleanupOldMdFiles(fso As Object, sheetFolder As String)
    On Error Resume Next
    Dim p As String
    p = sheetFolder & "\" & mWsSource.codeName & ".old.md"
    If fso.FileExists(p) Then fso.DeleteFile p, True
    If Not mWsSource Is mWsDest Then
        p = sheetFolder & "\" & mWsDest.codeName & ".old.md"
        If fso.FileExists(p) Then fso.DeleteFile p, True
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Macro parsing
' ============================================================

Private Function ParseCutOperation( _
    ByVal code As String, _
    ByRef srcSheetName As String, _
    ByRef srcRange As String, _
    ByRef dstSheetName As String, _
    ByRef dstRange As String, _
    ByRef unexpectedLines As String _
) As Boolean
    srcSheetName = ""
    srcRange = ""
    dstSheetName = ""
    dstRange = ""
    unexpectedLines = ""

    Dim lines() As String
    lines = Split(Replace(code, vbCrLf, vbLf), vbLf)

    ' State machine for Pattern B
    Dim pendingSrcSheet As String
    Dim pendingSrcRange As String
    Dim pendingDstSheet As String
    Dim pendingDstRange As String
    Dim afterCut As Boolean
    pendingSrcSheet = ""
    pendingSrcRange = ""
    pendingDstSheet = ""
    pendingDstRange = ""
    afterCut = False

    ' Track "active sheet" from Sheets("...").Select or .Activate
    Dim activeSheet As String
    activeSheet = ""

    Dim ln As String
    Dim i As Long
    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) = 0 Then GoTo NextLine
        ' Skip Sub/End Sub/comments
        If ln Like "Sub *" Or ln Like "End Sub" Or Left(ln, 1) = "'" Then GoTo NextLine

        ' Application.CutCopyMode / Selection.Copy ? expected, skip
        If InStr(1, ln, "Application.CutCopyMode", vbTextCompare) > 0 Then GoTo NextLine
        If InStr(1, ln, "Selection.Copy", vbTextCompare) > 0 Then GoTo NextLine

        ' Sheets("...").Select or .Activate ? track active sheet
        Dim sheetSel As String
        sheetSel = ExtractSheetName(ln)
        If sheetSel <> "" And (InStr(1, ln, ".Select", vbTextCompare) > 0 Or InStr(1, ln, ".Activate", vbTextCompare) > 0) Then
            activeSheet = sheetSel
            GoTo NextLine
        End If

        ' Pattern A: Range("...").Cut Destination:=Range("...")
        If InStr(1, ln, ".Cut", vbTextCompare) > 0 And InStr(1, ln, "Destination", vbTextCompare) > 0 Then
            Dim aSrcSheet As String, aSrcRange As String, aDstSheet As String, aDstRange As String
            If ParsePatternA(ln, activeSheet, aSrcSheet, aSrcRange, aDstSheet, aDstRange) Then
                srcSheetName = aSrcSheet
                srcRange = aSrcRange
                dstSheetName = aDstSheet
                dstRange = aDstRange
                ParseCutOperation = True
                GoTo NextLine
            End If
        End If

        ' Pattern B step 1: Range("...").Select ? potential source selection
        If InStr(1, ln, ".Select", vbTextCompare) > 0 And InStr(1, ln, "Range(", vbTextCompare) > 0 And Not afterCut Then
            Dim bSrcSheet As String, bSrcRange As String
            If ExtractRangeFromSelect(ln, activeSheet, bSrcSheet, bSrcRange) Then
                pendingSrcSheet = bSrcSheet
                pendingSrcRange = bSrcRange
                GoTo NextLine
            End If
        End If

        ' Pattern B step 2: Selection.Cut
        If ln Like "Selection.Cut" Then
            afterCut = True
            GoTo NextLine
        End If

        ' Pattern B step 3 (after cut): Range("...").Select ? destination
        If afterCut And InStr(1, ln, ".Select", vbTextCompare) > 0 And InStr(1, ln, "Range(", vbTextCompare) > 0 Then
            Dim bDstSheet As String, bDstRange As String
            If ExtractRangeFromSelect(ln, activeSheet, bDstSheet, bDstRange) Then
                pendingDstSheet = bDstSheet
                pendingDstRange = bDstRange
                GoTo NextLine
            End If
        End If

        ' Pattern B step 4: ActiveSheet.Paste (case-insensitive; may have args)
        If InStr(1, ln, "ActiveSheet.Paste", vbTextCompare) = 1 And afterCut Then
            srcSheetName = pendingSrcSheet
            srcRange = pendingSrcRange
            dstSheetName = pendingDstSheet
            dstRange = pendingDstRange
            ParseCutOperation = True
            afterCut = False
            pendingSrcSheet = ""
            pendingSrcRange = ""
            GoTo NextLine
        End If

        ' .PasteSpecial ? expected, skip
        If InStr(1, ln, ".PasteSpecial", vbTextCompare) > 0 Then GoTo NextLine

        ' Any remaining .Select or .Activate without Sheets prefix ? expected navigation
        If (InStr(1, ln, ".Select", vbTextCompare) > 0 Or InStr(1, ln, ".Activate", vbTextCompare) > 0) Then
            GoTo NextLine
        End If

        ' Unexpected line
        If Len(unexpectedLines) > 0 Then unexpectedLines = unexpectedLines & vbLf
        unexpectedLines = unexpectedLines & "  " & ln

NextLine:
    Next i

    ' Pattern B fallback: Cut+Select without explicit ActiveSheet.Paste
    ' Handles Enter-to-paste or other variants that don't emit an explicit Paste line
    If Not ParseCutOperation Then
        If afterCut And Len(pendingSrcRange) > 0 And Len(pendingDstRange) > 0 Then
            srcSheetName = pendingSrcSheet
            srcRange = pendingSrcRange
            dstSheetName = pendingDstSheet
            dstRange = pendingDstRange
            ParseCutOperation = True
        End If
    End If
End Function

Private Function ParsePatternA( _
    ByVal ln As String, _
    ByVal activeSheet As String, _
    ByRef srcSheet As String, _
    ByRef srcRng As String, _
    ByRef dstSheet As String, _
    ByRef dstRng As String _
) As Boolean
    ' Sheets("S1").Range("A1:B2").Cut Destination:=Sheets("S2").Range("C3")
    ' Range("A1:B2").Cut Destination:=Range("C3")
    srcSheet = ""
    srcRng = ""
    dstSheet = ""
    dstRng = ""

    Dim cutPos As Long
    cutPos = InStr(1, ln, ".Cut", vbTextCompare)
    If cutPos = 0 Then Exit Function

    Dim lhsPart As String
    lhsPart = Left(ln, cutPos - 1)
    srcSheet = ExtractSheetName(lhsPart)
    If srcSheet = "" Then srcSheet = activeSheet
    srcRng = ExtractRangeAddr(lhsPart)
    If srcRng = "" Then Exit Function

    Dim destIdx As Long
    destIdx = InStr(cutPos, ln, "Destination:=", vbTextCompare)
    If destIdx = 0 Then Exit Function
    Dim rhsPart As String
    rhsPart = Mid(ln, destIdx + Len("Destination:="))
    dstSheet = ExtractSheetName(rhsPart)
    If dstSheet = "" Then dstSheet = activeSheet
    dstRng = ExtractRangeAddr(rhsPart)
    If dstRng = "" Then Exit Function

    ParsePatternA = True
End Function

Private Function ExtractRangeFromSelect( _
    ByVal ln As String, _
    ByVal activeSheet As String, _
    ByRef sheetName As String, _
    ByRef rangeAddr As String _
) As Boolean
    sheetName = ExtractSheetName(ln)
    If sheetName = "" Then sheetName = activeSheet
    rangeAddr = ExtractRangeAddr(ln)
    ExtractRangeFromSelect = (rangeAddr <> "")
End Function

Private Function ExtractSheetName(ByVal s As String) As String
    ' Extract from Sheets("name") or Worksheets("name") patterns
    Dim pos As Long
    pos = InStr(1, s, "Sheets(""", vbTextCompare)
    If pos = 0 Then pos = InStr(1, s, "Worksheets(""", vbTextCompare)
    If pos = 0 Then Exit Function

    Dim q1 As Long, q2 As Long
    q1 = InStr(pos, s, """")
    If q1 = 0 Then Exit Function
    q2 = InStr(q1 + 1, s, """")
    If q2 = 0 Then Exit Function
    ExtractSheetName = Mid(s, q1 + 1, q2 - q1 - 1)
End Function

Private Function ExtractRangeAddr(ByVal s As String) As String
    ' Extract address from Range("addr") pattern
    Dim pos As Long
    pos = InStr(1, s, "Range(""", vbTextCompare)
    If pos = 0 Then Exit Function
    Dim q1 As Long, q2 As Long
    q1 = InStr(pos, s, """")
    If q1 = 0 Then Exit Function
    q2 = InStr(q1 + 1, s, """")
    If q2 = 0 Then Exit Function
    ExtractRangeAddr = Mid(s, q1 + 1, q2 - q1 - 1)
End Function

' ============================================================
' Instruction text builder
' ============================================================

Private Function BuildMoveInstructionText( _
    isSameSheet As Boolean, _
    srcSheetName As String, _
    srcRange As String, _
    dstSheetName As String, _
    dstRange As String, _
    wsSource As Worksheet, _
    wsDest As Worksheet, _
    reducedAccuracy As Boolean, _
    unexpectedLines As String _
) As String
    Dim t As String
    Dim WARN As String: WARN = ChrW(9888) & " "

    If isSameSheet Then
        t = "The following cell range move was performed on Excel sheet """ & wsSource.Name & """" & _
            " (VBA CodeName: " & wsSource.codeName & "):" & vbCrLf & vbCrLf
        t = t & "  Cut " & srcRange & ", pasted to " & dstRange & " on the same sheet." & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any VBA references" & vbCrLf
        t = t & "that target the source range " & srcRange & " on this sheet to point to " & dstRange & "." & vbCrLf & vbCrLf
        t = t & "Sheet map before: @sheet/" & wsSource.codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after:  @sheet/" & wsSource.codeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Range(""" & srcRange & """), Range(""" & FirstCell(srcRange) & """), etc." & vbCrLf
        t = t & "- Numeric-form: Cells(R, C) where R and C correspond to cells within " & srcRange & vbCrLf
        t = t & "- Dynamic-form: references constructed from the source range coordinates" & vbCrLf & vbCrLf
        t = t & "Note: Formulas inside the moved range that reference cells outside it" & vbCrLf
        t = t & "      are updated automatically by Excel. Focus on VBA code references only."
    Else
        ' Determine actual dest worksheet for CodeName
        Dim dstCodeName As String
        dstCodeName = wsDest.codeName
        ' If actual dest sheet differs from wsDest, try to find it by name
        If dstSheetName <> wsDest.Name Then
            Dim ws As Object
            For Each ws In ThisWorkbook.Sheets
                If ws.Type = xlWorksheet And ws.Name = dstSheetName Then
                    dstCodeName = ws.codeName
                    Exit For
                End If
            Next ws
        End If

        t = "The following cell range move was performed across sheets:" & vbCrLf & vbCrLf
        t = t & "  Source:      " & srcSheetName & " (CodeName: " & wsSource.codeName & ")" & _
                "  range " & srcRange & vbCrLf
        t = t & "  Destination: " & dstSheetName & " (CodeName: " & dstCodeName & ")" & _
                "  range " & dstRange & vbCrLf & vbCrLf
        t = t & "Please review all .bas files in the src/ folder and update any VBA references" & vbCrLf
        t = t & "that target " & srcSheetName & "!" & srcRange & _
                " to point to " & dstSheetName & "!" & dstRange & "." & vbCrLf & vbCrLf
        t = t & "Sheet map before (source): @sheet/" & wsSource.codeName & ".old.md" & vbCrLf
        t = t & "Sheet map after  (source): @sheet/" & wsSource.codeName & ".md" & vbCrLf
        t = t & "Sheet map before (dest):   @sheet/" & dstCodeName & ".old.md" & vbCrLf
        t = t & "Sheet map after  (dest):   @sheet/" & dstCodeName & ".md" & vbCrLf & vbCrLf
        t = t & "Reference formats to check:" & vbCrLf
        t = t & "- String-form:  Sheets(""" & srcSheetName & """).Range(""" & srcRange & """), Range(""" & srcRange & """), etc." & vbCrLf
        t = t & "- Numeric-form: Cells(R, C) where R and C correspond to cells within " & srcRange & vbCrLf
        t = t & "- Dynamic-form: references constructed from the source range coordinates" & vbCrLf & vbCrLf
        t = t & "Note: Formulas inside the moved range that reference cells outside it" & vbCrLf
        t = t & "      are updated automatically by Excel. Focus on VBA code references only."
    End If

    If Len(unexpectedLines) > 0 Then
        t = t & vbCrLf & vbCrLf
        t = t & WARN & "The recorded macro contained the following unexpected operations." & vbCrLf
        t = t & "  Please review whether these affect any VBA references:" & vbCrLf & vbCrLf
        t = t & unexpectedLines
    End If
    
    Dim langName As String: langName = GetLangMeta("english_name")
    If Len(langName) = 0 Then langName = GetLangMeta("native_name")
    If Len(langName) > 0 And langName <> "English" Then
        t = t & vbCrLf & vbCrLf & "Please respond in " & langName & "."
    End If

    If reducedAccuracy Then
        t = t & vbCrLf & vbCrLf
        t = t & WARN & "The sheet map for the destination sheet before the move was not captured." & vbCrLf
        t = t & "  The ""Sheet map before (dest)"" above reflects the post-move state." & vbCrLf
        t = t & "  Accuracy of reference detection for the destination sheet may be reduced."
    End If

    BuildMoveInstructionText = t
End Function

Private Function FirstCell(ByVal addr As String) As String
    Dim colon As Long
    colon = InStr(addr, ":")
    If colon > 0 Then
        FirstCell = Left(addr, colon - 1)
    Else
        FirstCell = addr
    End If
End Function

Private Sub CopyToClipboard(text As String)
    Dim obj As Object
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.SetText text
    obj.PutInClipboard
End Sub

' ============================================================
' State persistence (registry) ? survives VBA runtime reset
' ============================================================

Private Sub SaveState()
    ' Serialize: srcCodeName|dstCodeName|showDialog|name,lines;name,lines;...
    Dim snapStr As String
    Dim i As Long
    For i = 1 To mSnapshot.count
        Dim si As Variant: si = mSnapshot(i)
        If i > 1 Then snapStr = snapStr & ";"
        snapStr = snapStr & si(0) & "," & CStr(si(3))
    Next i
    Dim stateStr As String
    stateStr = mWsSource.codeName & "|" & mWsDest.codeName & "|" & _
               IIf(mShowResultDialog, "1", "0") & "|" & snapStr
    SaveSetting APP_KEY, REG_SECT, REG_KEY, stateStr
End Sub

Private Function LoadState() As Boolean
    Dim stateStr As String
    stateStr = GetSetting(APP_KEY, REG_SECT, REG_KEY, "")
    If Len(stateStr) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(stateStr, "|", 4)
    If UBound(parts) < 3 Then Exit Function

    ' Restore source/dest worksheets by CodeName
    Set mWsSource = Nothing
    Set mWsDest = Nothing
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If ws.Type = xlWorksheet Then
            If ws.codeName = parts(0) Then Set mWsSource = ws
            If ws.codeName = parts(1) Then Set mWsDest = ws
        End If
    Next ws
    If mWsSource Is Nothing Or mWsDest Is Nothing Then Exit Function

    mShowResultDialog = (parts(2) = "1")

    ' Rebuild snapshot collection (name + lineCount only; Phase 2 only needs these)
    Set mSnapshot = New Collection
    If Len(parts(3)) > 0 Then
        Dim compPairs() As String
        compPairs = Split(parts(3), ";")
        Dim j As Long
        For j = 0 To UBound(compPairs)
            Dim pair() As String
            pair = Split(compPairs(j), ",")
            If UBound(pair) >= 1 Then
                Dim item(3) As Variant
                item(0) = pair(0)
                item(1) = 0
                item(2) = ""
                item(3) = CLng(pair(1))
                mSnapshot.Add item
            End If
        Next j
    End If

    LoadState = True
End Function

Private Sub ClearState()
    On Error Resume Next
    DeleteSetting APP_KEY, REG_SECT
    On Error GoTo 0
End Sub
