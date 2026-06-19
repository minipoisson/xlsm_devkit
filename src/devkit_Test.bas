Attribute VB_Name = "devkit_Test"
Option Explicit

' Optional feature: test an Excel workbook like code (Phase 1 MVP).
'
' This module exposes a small set of non-interactive Public functions that a
' PowerShell runner (tools/Invoke-XlsmDevkitTest.ps1) calls via Application.Run.
' Every function takes and returns a JSON string; none of them shows a dialog.
'
' Design split (see the test module spec):
'   - AI / human author test specs as JSON (workbook.meta.json, *.test.json).
'   - PowerShell orchestrates: temp copy, input generation, calls these APIs.
'   - This module performs the in-workbook operations and reports JSON results.
'
' The module is self-contained on purpose (its own minimal JSON parser/serializer
' and UTF-8 writer) so it does not depend on Private helpers inside xlsm_devkit.bas.
' It carries the "devkit_" prefix, so xlsm_devkit's existing machinery skips it on
' import/export and strips it from released workbooks automatically.
'
' Constraints honoured:
'   - ASCII only (no Japanese, em-dash written as --).
'   - FSO / ADODB only (no Dir/Open/Kill/FreeFile).
'   - Sheet names etc. are passed in at runtime via JSON, never hardcoded.
'
' Public API:
'   DevkitTestPing()                 -> connectivity, version, workbook name
'   DevkitResolveInputs(metaJson)    -> resolve role=input ranges / unlocked cells
'   DevkitApplyInputs(inputsJson)    -> write values into target cells
'   DevkitCalculateFullRebuild()     -> Application.CalculateFullRebuild
'   DevkitAssertNoErrors(targetJson) -> scan targets for Excel error values
'   DevkitWriteResult(resultJson)    -> write result.json + result.md

Private Const DEVKIT_TEST_VERSION As String = "0.1.0"

' ===== JSON parser state (module-level; one parse at a time) =====
Private mJson As String
Private mPos As Long

' =====================================================================
' Public API
' =====================================================================

Public Function DevkitTestPing() As String
    On Error GoTo EH
    DevkitTestPing = "{""ok"":true,""version"":" & JsonStr(DEVKIT_TEST_VERSION) & _
                     ",""workbook"":" & JsonStr(ThisWorkbook.Name) & "}"
    Exit Function
EH:
    DevkitTestPing = JsonErr(Err.Description)
End Function

Public Function DevkitResolveInputs(ByVal metaJson As String) As String
    On Error GoTo EH
    Dim meta As Object
    Set meta = JsonParse(metaJson)

    Dim detectUnlocked As Boolean
    detectUnlocked = True
    If meta.Exists("input_detection") Then
        Dim idet As Object: Set idet = meta("input_detection")
        If idet.Exists("unlocked_cells_on_protected_sheets") Then _
            detectUnlocked = CBool(idet("unlocked_cells_on_protected_sheets"))
    End If

    Dim sb As String
    sb = "{""inputs"":["
    Dim first As Boolean: first = True

    If meta.Exists("sheets") Then
        Dim shts As Object: Set shts = meta("sheets")
        Dim k As Variant
        For Each k In shts.Keys
            Dim sdef As Object: Set sdef = shts(k)
            Dim role As String: role = ""
            If sdef.Exists("role") Then role = CStr(sdef("role"))

            Dim ws As Worksheet
            If sdef.Exists("code_name") Then
                Set ws = GetSheetByCodeName(CStr(sdef("code_name")))
            Else
                Set ws = GetSheetByName(CStr(k))
            End If
            If Not ws Is Nothing Then
                ' explicit input ranges
                If sdef.Exists("inputs") Then
                    Dim arr As Object: Set arr = sdef("inputs")
                    Dim ii As Long
                    For ii = 1 To arr.Count
                        Dim it As Object: Set it = arr(ii)
                        If it.Exists("range") Then
                            AppendInput sb, first, ws.Name, ws.CodeName, "", CStr(it("range")), ""
                        End If
                    Next ii
                End If
                ' auto-detect unlocked cells on protected sheets
                If detectUnlocked And IsInputRole(role) And ws.ProtectContents Then
                    Dim c As Range
                    For Each c In ws.UsedRange.Cells
                        If c.Locked = False Then
                            AppendInput sb, first, ws.Name, ws.CodeName, _
                                c.Address(False, False), "", CStr(c.Text)
                        End If
                    Next c
                End If
            End If
        Next k
    End If

    sb = sb & "]}"
    DevkitResolveInputs = sb
    Exit Function
EH:
    DevkitResolveInputs = JsonErr(Err.Description)
End Function

Public Function DevkitApplyInputs(ByVal inputsJson As String) As String
    On Error GoTo EH
    Dim payload As Object: Set payload = JsonParse(inputsJson)
    Dim arr As Object: Set arr = payload("inputs")

    Dim applied As Long: applied = 0
    Dim skipped As String: skipped = ""
    Dim sfirst As Boolean: sfirst = True

    Dim i As Long
    For i = 1 To arr.Count
        Dim it As Object: Set it = arr(i)
        Dim ws As Worksheet: Set ws = ResolveSheetRef(it)
        If ws Is Nothing Then
            AppendSkip skipped, sfirst, SheetRefLabel(it), "", "sheet not found"
        Else
            Dim target As Range
            Set target = Nothing
            On Error Resume Next
            If it.Exists("address") Then
                Set target = ws.Range(CStr(it("address")))
            ElseIf it.Exists("range") Then
                Set target = ws.Range(CStr(it("range")))
            End If
            On Error GoTo EH

            If target Is Nothing Then
                AppendSkip skipped, sfirst, SheetRefLabel(it), "", "bad address"
            Else
                Dim vtype As String: vtype = ""
                If it.Exists("type") Then vtype = LCase(CStr(it("type")))
                Dim vval As String: vval = ""
                If it.Exists("value") Then vval = CStr(it("value"))

                If ApplyValueToRange(target, vtype, vval) Then
                    applied = applied + 1
                Else
                    AppendSkip skipped, sfirst, SheetRefLabel(it), _
                        target.Address(False, False), "write failed (locked?)"
                End If
            End If
        End If
    Next i

    DevkitApplyInputs = "{""applied"":" & applied & ",""skipped"":[" & skipped & "]}"
    Exit Function
EH:
    DevkitApplyInputs = JsonErr(Err.Description)
End Function

Public Function DevkitCalculateFullRebuild() As String
    On Error GoTo EH
    Dim t0 As Double: t0 = Timer
    Application.CalculateFullRebuild
    Dim ms As Long: ms = CLng((Timer - t0) * 1000)
    If ms < 0 Then ms = 0    ' guard against midnight wrap of Timer
    DevkitCalculateFullRebuild = "{""ok"":true,""ms"":" & ms & "}"
    Exit Function
EH:
    DevkitCalculateFullRebuild = JsonErr(Err.Description)
End Function

Public Function DevkitAssertNoErrors(ByVal targetJson As String) As String
    On Error GoTo EH
    Dim payload As Object: Set payload = JsonParse(targetJson)
    Dim arr As Object: Set arr = payload("targets")

    Const MAX_FAIL As Long = 500
    Dim sb As String: sb = ""
    Dim first As Boolean: first = True
    Dim checked As Long: checked = 0
    Dim failCount As Long: failCount = 0
    Dim truncated As Boolean: truncated = False

    Dim i As Long
    For i = 1 To arr.Count
        Dim it As Object: Set it = arr(i)
        Dim ws As Worksheet: Set ws = ResolveSheetRef(it)
        If Not ws Is Nothing Then
            Dim rng As Range
            If it.Exists("used_range") And HasTrue(it, "used_range") Then
                Set rng = ws.UsedRange
            ElseIf it.Exists("range") Then
                Set rng = ws.Range(CStr(it("range")))
            Else
                Set rng = ws.UsedRange
            End If

            Dim c As Range
            For Each c In rng.Cells
                checked = checked + 1
                If IsError(c.Value2) Then
                    failCount = failCount + 1
                    If failCount <= MAX_FAIL Then
                        Dim fml As String: fml = ""
                        If c.HasFormula Then fml = c.Formula
                        If Not first Then sb = sb & ","
                        first = False
                        sb = sb & "{""sheet"":" & JsonStr(ws.Name) & _
                             ",""address"":" & JsonStr(c.Address(False, False)) & _
                             ",""error"":" & JsonStr(CStr(c.Text)) & _
                             ",""formula"":" & JsonStr(fml) & "}"
                    Else
                        truncated = True
                    End If
                End If
            Next c
        End If
    Next i

    DevkitAssertNoErrors = "{""failures"":[" & sb & "],""checked"":" & checked & _
                           ",""failCount"":" & failCount & _
                           ",""truncated"":" & LCase(CStr(truncated)) & "}"
    Exit Function
EH:
    DevkitAssertNoErrors = JsonErr(Err.Description)
End Function

Public Function DevkitWriteResult(ByVal resultJson As String) As String
    On Error GoTo EH
    Dim payload As Object: Set payload = JsonParse(resultJson)

    Dim outDir As String
    If payload.Exists("outputDir") Then
        outDir = CStr(payload("outputDir"))
    Else
        outDir = ThisWorkbook.Path & "\test-results\latest"
    End If
    EnsureFolder outDir

    Dim jsonPath As String: jsonPath = outDir & "\result.json"
    Dim mdPath As String: mdPath = outDir & "\result.md"

    WriteUTF8File jsonPath, JsonSerialize(payload)
    WriteUTF8File mdPath, RenderResultMd(payload)

    DevkitWriteResult = "{""ok"":true,""json"":" & JsonStr(jsonPath) & _
                        ",""md"":" & JsonStr(mdPath) & "}"
    Exit Function
EH:
    DevkitWriteResult = JsonErr(Err.Description)
End Function

' Applies one or more fixtures BEFORE a test run: seed specific cells with values
' or formulas. Each fixture supplies a sheet reference (code_name / sheet) plus a
' Markdown sheet-map table ("md"); only the listed cells are written (NON-destructive
' -- no ClearContents / ClearFormats / Validation.Delete). Reuses the core Markdown
' helpers ParseMDTableRow and UnescapeCellValue from xlsm_devkit.bas.
'   fixtureJson:
'     { "fixtures": [ { "code_name": "Sheet2",
'                       "md": "| Address | Name | Value | Formula | Style |\n...| B2 | | 5 | | |" } ] }
Public Function DevkitApplyFixture(ByVal fixtureJson As String) As String
    On Error GoTo EH
    Dim payload As Object: Set payload = JsonParse(fixtureJson)
    Dim arr As Object: Set arr = payload("fixtures")

    Dim applied As Long: applied = 0
    Dim skipped As String: skipped = ""
    Dim sfirst As Boolean: sfirst = True

    Dim fi As Long
    For fi = 1 To arr.Count
        Dim fx As Object: Set fx = arr(fi)
        Dim ws As Worksheet: Set ws = ResolveSheetRef(fx)
        If ws Is Nothing Then
            AppendSkip skipped, sfirst, SheetRefLabel(fx), "", "sheet not found"
        Else
            Dim md As String: md = ""
            If fx.Exists("md") Then md = CStr(fx("md"))
            Dim lines() As String
            lines = Split(Replace(Replace(md, vbCrLf, vbLf), vbCr, vbLf), vbLf)
            Dim li As Long
            For li = 0 To UBound(lines)
                Dim ln As String: ln = Trim(lines(li))
                If Len(ln) > 0 And Left(ln, 1) = "|" Then
                    Dim fields() As String
                    fields = ParseMDTableRow(ln)    ' Public in xlsm_devkit.bas
                    If UBound(fields) >= 0 Then
                        Dim addr As String: addr = Trim(fields(0))
                        If Not IsHeaderOrSeparatorRow(addr) Then
                            ' Trim alignment padding (hand-authored tables may pad columns
                            ' with multiple spaces; the core field trim only strips one).
                            Dim cValue As String: cValue = ""
                            Dim cFormula As String: cFormula = ""
                            If UBound(fields) >= 2 Then cValue = Trim(fields(2))
                            If UBound(fields) >= 3 Then cFormula = Trim(fields(3))
                            If ApplyFixtureCell(ws, addr, cValue, cFormula) Then
                                applied = applied + 1
                            Else
                                AppendSkip skipped, sfirst, SheetRefLabel(fx), addr, "write failed / bad address"
                            End If
                        End If
                    End If
                End If
            Next li
        End If
    Next fi

    DevkitApplyFixture = "{""applied"":" & applied & ",""skipped"":[" & skipped & "]}"
    Exit Function
EH:
    DevkitApplyFixture = JsonErr(Err.Description)
End Function

' =====================================================================
' Worksheet / value helpers
' =====================================================================

Private Function GetSheetByName(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set GetSheetByName = ws
            Exit Function
        End If
    Next ws
    Set GetSheetByName = Nothing
End Function

Private Function GetSheetByCodeName(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.CodeName, nm, vbTextCompare) = 0 Then
            Set GetSheetByCodeName = ws
            Exit Function
        End If
    Next ws
    Set GetSheetByCodeName = Nothing
End Function

' Resolves a worksheet from a reference dict carrying "code_name" and/or "sheet".
' The VBA CodeName takes priority when present (it is collision-free with tab names).
Private Function ResolveSheetRef(ByVal d As Object) As Worksheet
    If d.Exists("code_name") Then
        Set ResolveSheetRef = GetSheetByCodeName(CStr(d("code_name")))
        If Not ResolveSheetRef Is Nothing Then Exit Function
    End If
    If d.Exists("sheet") Then
        Set ResolveSheetRef = GetSheetByName(CStr(d("sheet")))
        Exit Function
    End If
    Set ResolveSheetRef = Nothing
End Function

' Human-readable label for a sheet reference, for skip/diagnostic messages.
Private Function SheetRefLabel(ByVal d As Object) As String
    If d.Exists("code_name") Then
        SheetRefLabel = CStr(d("code_name"))
    ElseIf d.Exists("sheet") Then
        SheetRefLabel = CStr(d("sheet"))
    Else
        SheetRefLabel = ""
    End If
End Function

Private Function IsInputRole(ByVal role As String) As Boolean
    IsInputRole = (InStr(1, role, "input", vbTextCompare) > 0)
End Function

' Writes a value into a range honouring an optional type hint.
'   type "blank"  -> clear contents
'   type "text"   -> force text (apostrophe prefix; avoids format mutation)
'   type "number" -> numeric (Val uses "." decimal, locale-independent)
'   type absent   -> infer: blank if "", number if numeric, else text
' Returns False if the write raised (e.g. target cell locked on a protected sheet).
Private Function ApplyValueToRange(ByVal target As Range, ByVal vtype As String, ByVal vval As String) As Boolean
    Dim ok As Boolean: ok = True
    On Error Resume Next
    Err.Clear
    Select Case vtype
        Case "blank"
            target.ClearContents
        Case "text"
            target.Value = "'" & vval
        Case "number"
            target.Value = Val(vval)
        Case Else
            If vval = "" Then
                target.ClearContents
            ElseIf IsNumeric(vval) Then
                target.Value = Val(vval)
            Else
                target.Value = vval
            End If
    End Select
    If Err.Number <> 0 Then ok = False
    On Error GoTo 0
    ApplyValueToRange = ok
End Function

' True for the Markdown table header row ("| Address | ... |") or a separator row
' ("|---|---|"), so fixture application skips them rather than treating them as cells.
Private Function IsHeaderOrSeparatorRow(ByVal addr As String) As Boolean
    If Len(addr) = 0 Then IsHeaderOrSeparatorRow = True: Exit Function
    If LCase(addr) = "address" Then IsHeaderOrSeparatorRow = True: Exit Function
    Dim i As Long, ch As String
    For i = 1 To Len(addr)
        ch = Mid(addr, i, 1)
        If ch <> "-" And ch <> ":" And ch <> " " Then Exit Function
    Next i
    IsHeaderOrSeparatorRow = True    ' only dashes/colons/spaces -> separator row
End Function

' Applies one fixture cell (value/formula only, non-destructive), mirroring the core
' import precedence: a formula wins over a literal value. Returns False on bad address
' or a write error. Reuses xlsm_devkit's Public UnescapeCellValue.
Private Function ApplyFixtureCell(ByVal ws As Worksheet, ByVal addr As String, _
                                  ByVal cValue As String, ByVal cFormula As String) As Boolean
    Dim target As Range
    Set target = Nothing
    On Error Resume Next
    Set target = ws.Range(addr)
    On Error GoTo 0
    If target Is Nothing Then Exit Function

    Dim ok As Boolean: ok = True
    On Error Resume Next
    Err.Clear
    If cFormula <> "-" And cFormula <> "" Then
        Dim fml As String: fml = cFormula
        If Left(fml, 1) = "`" And Right(fml, 1) = "`" Then fml = Mid(fml, 2, Len(fml) - 2)
        target.Formula = UnescapeCellValue(fml)
    ElseIf cValue <> "-" And cValue <> "" Then
        target.Value = UnescapeCellValue(cValue)
    End If
    If Err.Number <> 0 Then ok = False
    On Error GoTo 0
    ApplyFixtureCell = ok
End Function

Private Sub AppendInput(ByRef sb As String, ByRef first As Boolean, _
                        ByVal sheet As String, ByVal codeName As String, ByVal addr As String, _
                        ByVal rng As String, ByVal current As String)
    If Not first Then sb = sb & ","
    first = False
    sb = sb & "{""sheet"":" & JsonStr(sheet)
    If Len(codeName) > 0 Then sb = sb & ",""code_name"":" & JsonStr(codeName)
    If Len(addr) > 0 Then sb = sb & ",""address"":" & JsonStr(addr)
    If Len(rng) > 0 Then sb = sb & ",""range"":" & JsonStr(rng)
    sb = sb & ",""current"":" & JsonStr(current) & "}"
End Sub

Private Sub AppendSkip(ByRef sb As String, ByRef first As Boolean, _
                       ByVal sheet As String, ByVal addr As String, ByVal reason As String)
    If Not first Then sb = sb & ","
    first = False
    sb = sb & "{""sheet"":" & JsonStr(sheet) & ",""address"":" & JsonStr(addr) & _
         ",""reason"":" & JsonStr(reason) & "}"
End Sub

Private Function HasTrue(ByVal d As Object, ByVal key As String) As Boolean
    If Not d.Exists(key) Then Exit Function
    Dim v As Variant: v = d(key)
    If VarType(v) = vbBoolean Then
        HasTrue = v
    Else
        HasTrue = (LCase(CStr(v)) = "true")
    End If
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then EnsureFolder parent
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

' Writes content as UTF-8 without BOM (ADODB.Stream prepends a 3-byte BOM that we strip).
Private Sub WriteUTF8File(ByVal filePath As String, ByVal content As String)
    Dim stText As Object: Set stText = CreateObject("ADODB.Stream")
    stText.Type = 2
    stText.Charset = "UTF-8"
    stText.Open
    stText.WriteText content
    stText.Position = 0
    stText.Type = 1
    stText.Position = 3    ' skip BOM (EF BB BF)

    Dim stBin As Object: Set stBin = CreateObject("ADODB.Stream")
    stBin.Type = 1
    stBin.Open
    stBin.Write stText.Read
    stBin.SaveToFile filePath, 2    ' 2 = overwrite
    stBin.Close
    stText.Close
End Sub

' =====================================================================
' result.md rendering
' =====================================================================

Private Function RenderResultMd(ByVal payload As Object) As String
    Dim md As String
    md = "# Test Result" & vbCrLf & vbCrLf
    If payload.Exists("workbook") Then md = md & "- Workbook: " & CStr(payload("workbook")) & vbCrLf
    If payload.Exists("generatedAt") Then md = md & "- Generated: " & CStr(payload("generatedAt")) & vbCrLf

    If payload.Exists("summary") Then
        Dim sm As Object: Set sm = payload("summary")
        If sm.Exists("cases") Then md = md & "- Cases: " & NumPlain(sm("cases")) & vbCrLf
        If sm.Exists("failedCases") Then md = md & "- Failed cases: " & NumPlain(sm("failedCases")) & vbCrLf
        If sm.Exists("passed") Then md = md & "- Result: " & IIf(HasTrue(sm, "passed"), "PASS", "FAIL") & vbCrLf
    End If
    md = md & vbCrLf

    If payload.Exists("failures") Then
        Dim fl As Object: Set fl = payload("failures")
        If fl.Count > 0 Then
            md = md & "## Failures" & vbCrLf & vbCrLf
            md = md & "| Case | Sheet | Address | Error | Formula | Inputs |" & vbCrLf
            md = md & "|------|-------|---------|-------|---------|--------|" & vbCrLf
            Dim i As Long
            For i = 1 To fl.Count
                Dim f As Object: Set f = fl(i)
                md = md & "| " & MdField(f, "case") & " | " & MdField(f, "sheet") & _
                     " | " & MdField(f, "address") & " | " & MdField(f, "error") & _
                     " | " & MdCode(MdField(f, "formula")) & " | " & MdField(f, "inputs") & " |" & vbCrLf
            Next i
        Else
            md = md & "All cases passed. No Excel errors detected." & vbCrLf
        End If
    End If

    RenderResultMd = md
End Function

Private Function MdField(ByVal d As Object, ByVal key As String) As String
    Dim s As String
    If d.Exists(key) Then s = CStr(d(key))
    MdField = EscapeMd(s)
End Function

Private Function EscapeMd(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, "|", "\|")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    EscapeMd = s
End Function

Private Function MdCode(ByVal s As String) As String
    If Len(s) = 0 Then MdCode = "" Else MdCode = "`" & s & "`"
End Function

Private Function NumPlain(ByVal v As Variant) As String
    NumPlain = Replace(CStr(v), ",", ".")
End Function

' =====================================================================
' Minimal JSON serializer (Dictionary / Collection / scalar -> JSON)
' =====================================================================

Private Function JsonSerialize(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then
            JsonSerialize = "null"
        ElseIf TypeName(v) = "Dictionary" Then
            Dim o As String: o = "{"
            Dim k As Variant, f As Boolean: f = True
            For Each k In v.Keys
                If Not f Then o = o & ","
                f = False
                o = o & JsonStr(CStr(k)) & ":" & JsonSerialize(v(k))
            Next k
            JsonSerialize = o & "}"
        ElseIf TypeName(v) = "Collection" Then
            Dim a As String: a = "["
            Dim i As Long, g As Boolean: g = True
            For i = 1 To v.Count
                If Not g Then a = a & ","
                g = False
                a = a & JsonSerialize(v(i))
            Next i
            JsonSerialize = a & "]"
        Else
            JsonSerialize = "null"
        End If
    Else
        Select Case VarType(v)
            Case vbBoolean: JsonSerialize = LCase(CStr(v))
            Case vbEmpty, vbNull: JsonSerialize = "null"
            Case vbString: JsonSerialize = JsonStr(CStr(v))
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                JsonSerialize = NumPlain(v)
            Case Else: JsonSerialize = JsonStr(CStr(v))
        End Select
    End If
End Function

Private Function JsonStr(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    out = """"
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case ch
            Case """": out = out & "\"""
            Case "\": out = out & "\\"
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right$("000" & Hex$(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsonStr = out & """"
End Function

Private Function JsonErr(ByVal msg As String) As String
    JsonErr = "{""ok"":false,""error"":" & JsonStr(msg) & "}"
End Function

' =====================================================================
' Minimal recursive-descent JSON parser
'   object -> Scripting.Dictionary, array -> Collection,
'   string -> String, number -> Double, true/false -> Boolean, null -> Empty
' Input is produced by the PowerShell runner, so the grammar is the well-formed
' JSON subset we emit; no error recovery beyond surfacing a VBA runtime error.
' =====================================================================

Private Function JsonParse(ByVal s As String) As Object
    mJson = s
    mPos = 1
    Dim v As Variant
    AssignVar v, JsonParseValue()
    Set JsonParse = v
End Function

Private Function JsonParseValue() As Variant
    JsonSkipWs
    Dim ch As String: ch = Mid(mJson, mPos, 1)
    Select Case ch
        Case "{": Set JsonParseValue = JsonParseObject
        Case "[": Set JsonParseValue = JsonParseArray
        Case """": JsonParseValue = JsonParseString
        Case "t", "f": JsonParseValue = JsonParseBool
        Case "n": JsonParseNull: JsonParseValue = Empty
        Case Else: JsonParseValue = JsonParseNumber
    End Select
End Function

Private Function JsonParseObject() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    mPos = mPos + 1    ' consume {
    JsonSkipWs
    If Mid(mJson, mPos, 1) = "}" Then
        mPos = mPos + 1
        Set JsonParseObject = d
        Exit Function
    End If
    Dim sep As String
    Do
        JsonSkipWs
        Dim k As String: k = JsonParseString
        JsonSkipWs
        mPos = mPos + 1    ' consume :
        Dim v As Variant
        AssignVar v, JsonParseValue()
        If d.Exists(k) Then d.Remove k
        d.Add k, v
        JsonSkipWs
        sep = Mid(mJson, mPos, 1)
        mPos = mPos + 1    ' consume , or }
    Loop While sep = ","
    Set JsonParseObject = d
End Function

Private Function JsonParseArray() As Object
    Dim col As Object: Set col = New Collection
    mPos = mPos + 1    ' consume [
    JsonSkipWs
    If Mid(mJson, mPos, 1) = "]" Then
        mPos = mPos + 1
        Set JsonParseArray = col
        Exit Function
    End If
    Dim sep As String
    Do
        Dim v As Variant
        AssignVar v, JsonParseValue()
        col.Add v
        JsonSkipWs
        sep = Mid(mJson, mPos, 1)
        mPos = mPos + 1    ' consume , or ]
    Loop While sep = ","
    Set JsonParseArray = col
End Function

Private Function JsonParseString() As String
    mPos = mPos + 1    ' consume opening quote
    Dim out As String, ch As String, esc As String, hex4 As String
    Do While mPos <= Len(mJson)
        ch = Mid(mJson, mPos, 1)
        If ch = """" Then
            mPos = mPos + 1
            JsonParseString = out
            Exit Function
        ElseIf ch = "\" Then
            mPos = mPos + 1
            esc = Mid(mJson, mPos, 1)
            Select Case esc
                Case """": out = out & """"
                Case "\": out = out & "\"
                Case "/": out = out & "/"
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr(8)
                Case "f": out = out & Chr(12)
                Case "u"
                    hex4 = Mid(mJson, mPos + 1, 4)
                    out = out & ChrW(CLng("&H" & hex4))
                    mPos = mPos + 4
                Case Else: out = out & esc
            End Select
            mPos = mPos + 1
        Else
            out = out & ch
            mPos = mPos + 1
        End If
    Loop
    JsonParseString = out
End Function

Private Function JsonParseNumber() As Variant
    Dim startPos As Long: startPos = mPos
    Do While mPos <= Len(mJson)
        Select Case Mid(mJson, mPos, 1)
            Case "0" To "9", "-", "+", ".", "e", "E": mPos = mPos + 1
            Case Else: Exit Do
        End Select
    Loop
    JsonParseNumber = Val(Mid(mJson, startPos, mPos - startPos))
End Function

Private Function JsonParseBool() As Boolean
    If Mid(mJson, mPos, 4) = "true" Then
        mPos = mPos + 4
        JsonParseBool = True
    Else
        mPos = mPos + 5    ' false
        JsonParseBool = False
    End If
End Function

Private Sub JsonParseNull()
    mPos = mPos + 4    ' null
End Sub

Private Sub JsonSkipWs()
    Do While mPos <= Len(mJson)
        Select Case Mid(mJson, mPos, 1)
            Case " ", vbTab, vbCr, vbLf: mPos = mPos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub AssignVar(ByRef target As Variant, ByVal src As Variant)
    If IsObject(src) Then
        Set target = src
    Else
        target = src
    End If
End Sub
