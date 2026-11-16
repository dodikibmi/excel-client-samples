Attribute VB_Name = "modCore"
Option Explicit
' Quote & invoice builder: the core.
' The form, the Home buttons and the build script all go through these functions, so every document is numbered,
' totalled, checked and logged the same way. Functions return text so a caller (button, form or test) can show it:
'   NewDocument -> "Q-2026-0001" or "ERR: ..."     IssueDocument -> "" or "BLOCKED: ..." / "ERR: ..."
' Sheets are protected without a password; the core unprotects, writes and protects again.

Public Const REG_FIRST As Long = 6, LIN_FIRST As Long = 6, LOG_FIRST As Long = 6, MAX_LINES As Long = 15
' Register columns
Public Const R_NO = 1, R_TYPE = 2, R_DATE = 3, R_CUST = 4, R_NAME = 5, R_EMAIL = 6, R_LINES = 7, R_SUB = 8, _
    R_DPCT = 9, R_DISC = 10, R_TRATE = 11, R_TAX = 12, R_TOTAL = 13, R_STATUS = 14, R_DUE = 15, R_SENT = 16, _
    R_DONE = 17, R_LINK = 18, R_PDF = 19, R_CHECK = 20, R_NOTES = 21, R_FORMULAS = 22
' Lines columns
Public Const L_NO = 1, L_LINE = 2, L_CODE = 3, L_DESC = 4, L_UNIT = 5, L_QTY = 6, L_PRICE = 7, L_DISC = 8, _
    L_NET = 9, L_FORMULAS = 10
' Customers / Products columns
Public Const C_ID = 1, C_NAME = 2, C_CONTACT = 3, C_EMAIL = 4, C_CITY = 5, C_TERMS = 6, C_EXEMPT = 7
Public Const P_CODE = 1, P_CAT = 2, P_DESC = 3, P_UNIT = 4, P_LIST = 5, P_MIN = 6, P_MAXD = 7

Public ReplayClock As Date        ' set only by the build script when it replays the sample history

' ---------------------------------------------------------------- small helpers
Public Function S(nm As String) As Variant
    S = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Public Function Sh(nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Public Function LastRow(ws As Worksheet, firstRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < firstRow Then r = firstRow - 1
    LastRow = r
End Function

Public Function R2(ByVal x As Variant) As Variant      ' round half away from zero, like Excel's ROUND (VBA Round is banker's)
    Dim d As Variant
    d = CDec(x) * 100
    If d >= 0 Then R2 = Fix(d + CDec(0.5)) / 100 Else R2 = Fix(d - CDec(0.5)) / 100
End Function

' Number and date text built by hand: VBA's Format$ follows the PC's regional settings (decimal comma, local month
' names), and a quote must read the same on every PC.
Public Function N2(ByVal x As Variant, Optional ByVal grp As Boolean = True) As String
    Dim c As Variant, neg As Boolean, ip As String, fp As String, out As String
    c = R2(x)
    If c < 0 Then neg = True: c = -c
    ip = Trim$(Str$(Fix(c)))
    fp = Right$("00" & Trim$(Str$(CLng((c - Fix(c)) * 100))), 2)
    If grp Then
        Do While Len(ip) > 3
            out = "," & Right$(ip, 3) & out
            ip = Left$(ip, Len(ip) - 3)
        Loop
    End If
    N2 = IIf(neg, "-", "") & ip & out & "." & fp
End Function

Public Function DTxt(ByVal d As Date) As String
    DTxt = Day(d) & " " & Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", 3 * Month(d) - 2, 3) & " " & Year(d)
End Function

Public Function Money(ByVal x As Variant) As String
    Money = S("Cur") & N2(x)
End Function

Public Function Pct(ByVal x As Variant) As String
    Pct = Trim$(Str$(R2(CDec(x) * 100))) & "%"
End Function

Public Function Stamp() As Date
    If ReplayClock <> 0 Then Stamp = ReplayClock Else Stamp = Now
End Function

Public Sub SetClock(ByVal iso As String)            ' build script only: "2026-08-03 09:00"; "" = real clock
    If iso = "" Then ReplayClock = 0 Else ReplayClock = CDate(DateSerial(Val(Left$(iso, 4)), Val(Mid$(iso, 6, 2)), _
        Val(Mid$(iso, 9, 2))) + TimeSerial(Val(Mid$(iso, 12, 2)), Val(Mid$(iso, 15, 2)), 0))
End Sub

Public Sub SheetsOpen()
    Dim n As Variant
    For Each n In Array("Register", "Lines", "Log", "Document", "Home")
        Sh(CStr(n)).Unprotect
    Next
End Sub

Public Sub SheetsLock()
    Dim n As Variant
    For Each n In Array("Register", "Lines", "Log", "Document", "Home")
        Sh(CStr(n)).Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Next
End Sub

' The drop-downs of document numbers (Home "Selected document", Document page) follow the register.
Public Sub RefreshDocList()
    Dim last As Long, f As String, c As Variant
    last = LastRow(Sh("Register"), REG_FIRST)
    If last < REG_FIRST Then last = REG_FIRST
    f = "=Register!$A$" & REG_FIRST & ":$A$" & last
    For Each c In Array(Sh("Home").Range("SelDoc"), Sh("Document").Range("CurDoc"))
        c.Validation.Delete
        c.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=f
    Next
End Sub

Public Function FindRow(ws As Worksheet, firstRow As Long, key As String) As Long
    Dim m As Variant
    If key = "" Then Exit Function
    m = Application.Match(key, ws.Range(ws.Cells(firstRow, 1), ws.Cells(LastRow(ws, firstRow) + 1, 1)), 0)
    If Not IsError(m) Then FindRow = firstRow + m - 1
End Function

Public Function DocRow(no As String) As Long
    DocRow = FindRow(Sh("Register"), REG_FIRST, no)
End Function

Public Function CustRow(id As String) As Long
    CustRow = FindRow(Sh("Customers"), 6, id)
End Function

Public Function ProdRow(code As String) As Long
    ProdRow = FindRow(Sh("Products"), 6, code)
End Function

Public Sub LogIt(action As String, no As String, detail As String)
    Dim ws As Worksheet, r As Long
    Set ws = Sh("Log")
    r = LastRow(ws, LOG_FIRST) + 1
    ws.Cells(r, 1).Value = Stamp()
    ws.Cells(r, 2).Value = S("DeskUser")
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = no
    ws.Cells(r, 5).Value = detail
End Sub

' ---------------------------------------------------------------- lines: "CODE~qty~price~disc|CODE~..." (dot decimals, disc as a fraction)
Public Function ParseSpec(spec As String) As Variant
    Dim parts As Variant, i As Long, f As Variant, out() As Variant
    If Trim$(spec) = "" Then ParseSpec = Empty: Exit Function
    parts = Split(spec, "|")
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        f = Split(parts(i), "~")
        If UBound(f) < 3 Then out(i) = Array("", 0, 0, 0) Else out(i) = Array(Trim$(f(0)), CDec(Val(f(1))), CDec(Val(f(2))), CDec(Val(f(3))))
    Next
    ParseSpec = out
End Function

Public Function NumText(ByVal x As Variant) As String
    NumText = Trim$(Str$(x))
End Function

Public Function SpecOf(no As String) As String             ' the lines of a saved document, in line order
    Dim ws As Worksheet, r As Long, i As Long, n As Long, rows() As Long, out As String
    Set ws = Sh("Lines")
    For r = LIN_FIRST To LastRow(ws, LIN_FIRST)
        If ws.Cells(r, L_NO).Value = no Then n = n + 1
    Next
    If n = 0 Then Exit Function
    ReDim rows(1 To n)
    For r = LIN_FIRST To LastRow(ws, LIN_FIRST)
        If ws.Cells(r, L_NO).Value = no Then rows(ws.Cells(r, L_LINE).Value) = r
    Next
    For i = 1 To n
        r = rows(i)
        If i > 1 Then out = out & "|"
        out = out & ws.Cells(r, L_CODE).Value & "~" & NumText(ws.Cells(r, L_QTY).Value) & "~" & _
              NumText(ws.Cells(r, L_PRICE).Value) & "~" & NumText(ws.Cells(r, L_DISC).Value)
    Next
    SpecOf = out
End Function

Public Function LineNet(ByVal qty As Variant, ByVal price As Variant, ByVal disc As Variant) As Variant
    LineNet = R2(CDec(qty) * CDec(price) * (1 - CDec(disc)))
End Function

' Totals for a customer + discount + lines: Array(subtotal, discount, tax rate, tax, total)
Public Function Totals(custId As String, ByVal docDisc As Variant, spec As String) As Variant
    Dim L As Variant, i As Long, subT As Variant, disc As Variant, rate As Variant, tax As Variant, cr As Long
    subT = CDec(0)
    L = ParseSpec(spec)
    If Not IsEmpty(L) Then
        For i = 0 To UBound(L)
            subT = subT + LineNet(L(i)(1), L(i)(2), L(i)(3))
        Next
    End If
    disc = R2(subT * CDec(docDisc))
    cr = CustRow(custId)
    rate = CDec(S("TaxRate"))
    If cr > 0 Then If Sh("Customers").Cells(cr, C_EXEMPT).Value = "Yes" Then rate = CDec(0)
    tax = R2((subT - disc) * rate)
    Totals = Array(subT, disc, rate, tax, subT - disc + tax)
End Function

' Hard errors: the document cannot even be saved as a draft
Public Function InputErrors(custId As String, ByVal docDisc As Variant, spec As String) As String
    Dim L As Variant, i As Long, e As String
    If CustRow(custId) = 0 Then e = e & "Pick a customer." & vbLf
    If docDisc < 0 Or docDisc >= 1 Then e = e & "Document discount must be 0-99%." & vbLf
    L = ParseSpec(spec)
    If IsEmpty(L) Then
        e = e & "Add at least one line." & vbLf
    Else
        If UBound(L) + 1 > MAX_LINES Then e = e & "At most " & MAX_LINES & " lines per document." & vbLf
        For i = 0 To UBound(L)
            If ProdRow(CStr(L(i)(0))) = 0 Then e = e & "Line " & i + 1 & ": unknown product '" & L(i)(0) & "'." & vbLf
            If L(i)(1) <= 0 Then e = e & "Line " & i + 1 & ": quantity must be above 0." & vbLf
            If L(i)(2) < 0 Then e = e & "Line " & i + 1 & ": price cannot be negative." & vbLf
            If L(i)(3) < 0 Or L(i)(3) >= 1 Then e = e & "Line " & i + 1 & ": discount must be 0-99%." & vbLf
        Next
    End If
    If e <> "" Then e = Left$(e, Len(e) - 1)
    InputErrors = e
End Function

' Business checks before a document may be issued (a draft may break them; issuing is blocked until fixed)
Public Function CheckSpec(docType As String, custId As String, ByVal docDisc As Variant, spec As String) As String
    Dim e As String, L As Variant, i As Long, pr As Long, cr As Long, P As Worksheet
    e = InputErrors(custId, docDisc, spec)
    If e <> "" Then CheckSpec = e: Exit Function
    cr = CustRow(custId)
    If S("EmailMode") <> "Off" And Trim$(Sh("Customers").Cells(cr, C_EMAIL).Value) = "" Then _
        e = e & "Customer " & Sh("Customers").Cells(cr, C_NAME).Value & " has no email address (add it on Customers)." & vbLf
    If docDisc > S("MaxDocDisc") + 0.00001 Then _
        e = e & "Document discount " & Pct(docDisc) & " is above the " & Pct(S("MaxDocDisc")) & " limit." & vbLf
    Set P = Sh("Products")
    L = ParseSpec(spec)
    For i = 0 To UBound(L)
        pr = ProdRow(CStr(L(i)(0)))
        If L(i)(2) < CDec(P.Cells(pr, P_MIN).Value) Then e = e & "Line " & i + 1 & " " & L(i)(0) & ": price " & _
            N2(L(i)(2), False) & " is under the minimum " & N2(P.Cells(pr, P_MIN).Value, False) & "." & vbLf
        If L(i)(3) > CDec(P.Cells(pr, P_MAXD).Value) / 100 + 0.00001 Then e = e & "Line " & i + 1 & " " & L(i)(0) & _
            ": discount " & Pct(L(i)(3)) & " is above the " & P.Cells(pr, P_MAXD).Value & "% allowed." & vbLf
    Next
    If e <> "" Then e = Left$(e, Len(e) - 1)
    CheckSpec = e
End Function

Public Function CheckDocument(no As String) As String
    Dim r As Long, ws As Worksheet, e As String, q As Long
    r = DocRow(no)
    If r = 0 Then CheckDocument = "Document " & no & " not found.": Exit Function
    Set ws = Sh("Register")
    If ws.Cells(r, R_STATUS).Value <> "Draft" Then e = "Already issued (status " & ws.Cells(r, R_STATUS).Value & ")." & vbLf
    If Application.CountIf(ws.Columns(R_NO), no) > 1 Then e = e & "Number " & no & " is used twice in the register." & vbLf
    If ws.Cells(r, R_TYPE).Value = "Invoice" And ws.Cells(r, R_LINK).Value <> "" Then
        q = DocRow(CStr(ws.Cells(r, R_LINK).Value))
        If q = 0 Then
            e = e & "Quote " & ws.Cells(r, R_LINK).Value & " not found." & vbLf
        ElseIf ws.Cells(q, R_STATUS).Value <> "Accepted" Then
            e = e & "Quote " & ws.Cells(r, R_LINK).Value & " is not accepted." & vbLf
        End If
    End If
    e = e & CheckSpec(CStr(ws.Cells(r, R_TYPE).Value), CStr(ws.Cells(r, R_CUST).Value), ws.Cells(r, R_DPCT).Value, SpecOf(no))
    If Right$(e, 1) = vbLf Then e = Left$(e, Len(e) - 1)
    CheckDocument = e
End Function

' ---------------------------------------------------------------- numbering
Public Function NextNumber(docType As String, ByVal d As Date) As String
    Dim pre As String, ws As Worksheet, r As Long, mx As Long, v As String
    If docType = "Invoice" Then pre = S("InvPrefix") Else pre = S("QuotePrefix")
    pre = pre & Year(d) & "-"
    Set ws = Sh("Register")
    For r = REG_FIRST To LastRow(ws, REG_FIRST)
        v = CStr(ws.Cells(r, R_NO).Value)
        If Left$(v, Len(pre)) = pre Then
            If Val(Mid$(v, Len(pre) + 1)) > mx Then mx = Val(Mid$(v, Len(pre) + 1))
        End If
    Next
    NextNumber = pre & Format$(mx + 1, "0000")
End Function

' ---------------------------------------------------------------- writing documents
Private Sub WriteLines(no As String, spec As String)
    Dim ws As Worksheet, P As Worksheet, L As Variant, i As Long, r As Long, pr As Long
    Set ws = Sh("Lines"): Set P = Sh("Products")
    For r = LastRow(ws, LIN_FIRST) To LIN_FIRST Step -1          ' a draft being edited: drop its old lines
        If ws.Cells(r, L_NO).Value = no Then ws.Rows(r).Delete
    Next
    L = ParseSpec(spec)
    r = LastRow(ws, LIN_FIRST) + 1
    For i = 0 To UBound(L)
        pr = ProdRow(CStr(L(i)(0)))
        ws.Cells(r, L_NO).Value = no
        ws.Cells(r, L_LINE).Value = i + 1
        ws.Cells(r, L_CODE).Value = L(i)(0)
        ws.Cells(r, L_DESC).Value = P.Cells(pr, P_DESC).Value
        ws.Cells(r, L_UNIT).Value = P.Cells(pr, P_UNIT).Value
        ws.Cells(r, L_QTY).Value = CDbl(L(i)(1))
        ws.Cells(r, L_PRICE).Value = CDbl(L(i)(2))
        ws.Cells(r, L_DISC).Value = CDbl(L(i)(3))
        ws.Cells(r, L_NET).Value = CDbl(LineNet(L(i)(1), L(i)(2), L(i)(3)))
        ws.Range(ws.Cells(r, L_FORMULAS), ws.Cells(r, L_FORMULAS + 4)).FormulaR1C1 = Array( _
            "=RC1&""|""&RC2", _
            "=IF(ABS(ROUND(RC6*RC7*(1-RC8),2)-RC9)<0.005,""OK"",""Wrong"")", _
            "=IFERROR(INDEX(Products!C6,MATCH(RC3,Products!C1,0)),"""")", _
            "=IFERROR(INDEX(Register!C14,MATCH(RC1,Register!C1,0)),""Missing"")", _
            "=IF(AND(RC13<>""Draft"",ISNUMBER(RC12)),IF(RC7<RC12-0.004,""Under min"",""""),"""")")
        r = r + 1
    Next
End Sub

Private Sub WriteHeader(r As Long, no As String, docType As String, custId As String, ByVal d As Date, _
                        ByVal docDisc As Variant, notes As String, spec As String)
    Dim ws As Worksheet, C As Worksheet, cr As Long, t As Variant, L As Variant
    Set ws = Sh("Register"): Set C = Sh("Customers")
    cr = CustRow(custId)
    t = Totals(custId, docDisc, spec)
    L = ParseSpec(spec)
    ws.Cells(r, R_NO).Value = no
    ws.Cells(r, R_TYPE).Value = docType
    ws.Cells(r, R_DATE).Value = d
    ws.Cells(r, R_CUST).Value = custId
    ws.Cells(r, R_NAME).Value = C.Cells(cr, C_NAME).Value
    ws.Cells(r, R_EMAIL).Value = C.Cells(cr, C_EMAIL).Value
    ws.Cells(r, R_LINES).Value = UBound(L) + 1
    ws.Cells(r, R_SUB).Value = CDbl(t(0))
    ws.Cells(r, R_DPCT).Value = CDbl(docDisc)
    ws.Cells(r, R_DISC).Value = CDbl(t(1))
    ws.Cells(r, R_TRATE).Value = CDbl(t(2))
    ws.Cells(r, R_TAX).Value = CDbl(t(3))
    ws.Cells(r, R_TOTAL).Value = CDbl(t(4))
    ws.Cells(r, R_STATUS).Value = "Draft"
    If docType = "Invoice" Then
        ws.Cells(r, R_DUE).Value = d + C.Cells(cr, C_TERMS).Value
    Else
        ws.Cells(r, R_DUE).Value = d + S("QuoteValidDays")
    End If
    ws.Cells(r, R_NOTES).Value = notes
    ws.Range(ws.Cells(r, R_FORMULAS), ws.Cells(r, R_FORMULAS + 4)).FormulaR1C1 = Array( _
        "=IF(AND(RC14=""Sent"",RC15<AsOf),IF(RC2=""Quote"",""Expired"",""Overdue""),RC14)", _
        "=IF(RC22=""Overdue"",AsOf-RC15,"""")", _
        "=IF(ABS(SUMIFS(Lines!C9,Lines!C1,RC1)-RC8)>=0.005,""Lines"",IF(ABS(RC8-RC10+RC12-RC13)>=0.005,""Total"",""OK""))", _
        "=IF(RC22=""Overdue"",""Overdue ""&RC23&"" days"",IF(AND(RC2=""Quote"",RC22=""Sent"",RC15-AsOf<=7),""Expires ""&TEXT(RC15,""d mmm""),IF(AND(RC14=""Draft"",RC20<>""""),""Blocked: ""&RC20,IF(RC14=""Draft"",""Draft, not sent"",""""))))", _
        "=IF(RC2<>""Invoice"","""",IF(IFERROR(AND(INDEX(C14,MATCH(RC18,C1,0))=""Accepted"",INDEX(C18,MATCH(RC18,C1,0))=RC1),FALSE),""OK"",""Bad""))")
End Sub

Private Sub StoreCheck(no As String)
    Dim r As Long
    r = DocRow(no)
    If Sh("Register").Cells(r, R_STATUS).Value = "Draft" Then
        Sh("Register").Cells(r, R_CHECK).Value = Replace(CheckDocument(no), vbLf, " ")
    Else
        Sh("Register").Cells(r, R_CHECK).Value = ""
    End If
End Sub

Public Function NewDocument(docType As String, custId As String, ByVal d As Date, ByVal docDisc As Double, _
                            notes As String, spec As String) As String
    Dim e As String, no As String, r As Long, t As Variant
    If docType <> "Quote" And docType <> "Invoice" Then NewDocument = "ERR: type must be Quote or Invoice": Exit Function
    e = InputErrors(custId, docDisc, spec)
    If e <> "" Then NewDocument = "ERR: " & Replace(e, vbLf, " "): Exit Function
    no = NextNumber(docType, d)
    If DocRow(no) > 0 Then NewDocument = "ERR: number " & no & " is already used": Exit Function
    SheetsOpen
    r = LastRow(Sh("Register"), REG_FIRST) + 1
    WriteHeader r, no, docType, custId, d, docDisc, notes, spec
    WriteLines no, spec
    StoreCheck no
    t = Totals(custId, docDisc, spec)
    RefreshDocList
    LogIt "Created", no, docType & " for " & Sh("Register").Cells(r, R_NAME).Value & ", " & _
          Sh("Register").Cells(r, R_LINES).Value & " line(s), total " & Money(t(4))
    SheetsLock
    NewDocument = no
End Function

Public Function UpdateDraft(no As String, custId As String, ByVal d As Date, ByVal docDisc As Double, _
                            notes As String, spec As String) As String
    Dim r As Long, e As String, before As Variant, ws As Worksheet
    Set ws = Sh("Register")
    r = DocRow(no)
    If r = 0 Then UpdateDraft = "ERR: " & no & " not found": Exit Function
    If ws.Cells(r, R_STATUS).Value <> "Draft" Then UpdateDraft = "ERR: " & no & " is " & ws.Cells(r, R_STATUS).Value & _
        "; issued documents are frozen, create a new one": Exit Function
    e = InputErrors(custId, docDisc, spec)
    If e <> "" Then UpdateDraft = "ERR: " & Replace(e, vbLf, " "): Exit Function
    before = ws.Cells(r, R_TOTAL).Value
    SheetsOpen
    Dim link As Variant: link = ws.Cells(r, R_LINK).Value
    WriteHeader r, no, CStr(ws.Cells(r, R_TYPE).Value), custId, d, docDisc, notes, spec
    ws.Cells(r, R_LINK).Value = link
    WriteLines no, spec
    StoreCheck no
    LogIt "Edited", no, "Total " & Money(before) & " -> " & Money(ws.Cells(r, R_TOTAL).Value)
    SheetsLock
    UpdateDraft = ""
End Function

' ---------------------------------------------------------------- issuing: checks -> PDF -> email -> Sent
Public Function FolderOf(nm As String) As String
    Dim f As String
    f = CStr(S(nm))
    If InStr(f, ":") = 0 And Left$(f, 2) <> "\\" Then f = ThisWorkbook.Path & "\" & f
    If Dir(f, vbDirectory) = "" Then MkDir f
    FolderOf = f
End Function

Public Function SafeName(ByVal s0 As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "&")
        s0 = Replace(s0, ch, " ")
    Next
    SafeName = Application.WorksheetFunction.Trim(s0)
End Function

Public Function IssueDocument(no As String, ByVal d As Date) As String
    Dim e As String, r As Long, ws As Worksheet, f As String, pdf As String, mail As String
    e = CheckDocument(no)
    If e <> "" Then
        SheetsOpen
        StoreCheck no
        LogIt "Issue blocked", no, Replace(e, vbLf, " ")
        SheetsLock
        IssueDocument = "BLOCKED: " & e
        Exit Function
    End If
    Set ws = Sh("Register")
    r = DocRow(no)
    SheetsOpen
    ws.Cells(r, R_STATUS).Value = "Sent"
    ws.Cells(r, R_SENT).Value = d
    ws.Cells(r, R_CHECK).Value = ""
    f = no & " " & SafeName(CStr(ws.Cells(r, R_NAME).Value)) & ".pdf"
    pdf = FolderOf("PdfFolder") & "\" & f
    Sh("Document").Range("CurDoc").Value = no
    Application.Calculate
    On Error GoTo failed
    Sh("Document").ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdf, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0
    ws.Cells(r, R_PDF).Value = f
    mail = SendMail(r, pdf)
    LogIt "Issued", no, ws.Cells(r, R_TYPE).Value & " " & Money(ws.Cells(r, R_TOTAL).Value) & ", PDF " & f & mail
    SheetsLock
    IssueDocument = ""
    Exit Function
failed:
    ws.Cells(r, R_STATUS).Value = "Draft"
    ws.Cells(r, R_SENT).Value = ""
    LogIt "Issue failed", no, "PDF not written: " & Err.Description
    SheetsLock
    IssueDocument = "ERR: PDF not written: " & Err.Description
End Function

Private Function SendMail(r As Long, pdf As String) As String
    Dim ws As Worksheet, mode As String, subj As String, body As String, cr As Long, kind As String, whenTxt As String
    mode = S("EmailMode")
    If mode = "Off" Then Exit Function
    Set ws = Sh("Register")
    cr = CustRow(CStr(ws.Cells(r, R_CUST).Value))
    kind = LCase$(ws.Cells(r, R_TYPE).Value)
    If kind = "quote" Then whenTxt = "valid until " Else whenTxt = "due on "
    subj = ws.Cells(r, R_TYPE).Value & " " & ws.Cells(r, R_NO).Value & " from " & S("CoName")
    body = "Dear " & Sh("Customers").Cells(cr, C_CONTACT).Value & "," & vbCrLf & vbCrLf & _
           "Please find attached " & kind & " " & ws.Cells(r, R_NO).Value & " for " & Money(ws.Cells(r, R_TOTAL).Value) & _
           ", " & whenTxt & DTxt(ws.Cells(r, R_DUE).Value) & "." & vbCrLf & vbCrLf & _
           "Kind regards," & vbCrLf & S("DeskUser") & vbCrLf & S("CoName") & vbCrLf & S("CoPhone")
    If mode = "Outlook draft" Then
        Dim ol As Object, m As Object
        Set ol = CreateObject("Outlook.Application")
        Set m = ol.CreateItem(0)
        m.To = ws.Cells(r, R_EMAIL).Value
        m.Subject = subj
        m.Body = body
        m.Attachments.Add pdf
        m.Display
        SendMail = "; Outlook draft opened"
    Else
        Dim eml As String
        eml = FolderOf("OutboxFolder") & "\" & ws.Cells(r, R_NO).Value & ".eml"
        WriteEml eml, CStr(S("CoEmail")), CStr(ws.Cells(r, R_EMAIL).Value), subj, body, pdf
        SendMail = "; email " & ws.Cells(r, R_NO).Value & ".eml in " & S("OutboxFolder")
    End If
End Function

Private Function FileBase64(path As String) As String
    Dim st As Object, x As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.LoadFromFile path
    Set x = CreateObject("MSXML2.DOMDocument").createElement("b")
    x.DataType = "bin.base64"
    x.nodeTypedValue = st.Read
    st.Close
    FileBase64 = Replace(x.Text, vbLf, vbCrLf)
End Function

' An unsent email (.eml, "X-Unsent: 1"): Outlook / Windows Mail open it as a draft ready to send.
Public Sub WriteEml(path As String, fromAddr As String, toAddr As String, subj As String, body As String, attach As String)
    Dim b As String, t As String, fn As Integer, nm As String
    b = "----=_QIB_" & Format$(Stamp(), "yyyymmddhhnnss")
    nm = Mid$(attach, InStrRev(attach, "\") + 1)
    t = "From: " & fromAddr & vbCrLf & "To: " & toAddr & vbCrLf & "Subject: " & subj & vbCrLf & _
        "X-Unsent: 1" & vbCrLf & "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: multipart/mixed; boundary=""" & b & """" & vbCrLf & vbCrLf & _
        "--" & b & vbCrLf & "Content-Type: text/plain; charset=""us-ascii""" & vbCrLf & vbCrLf & body & vbCrLf & vbCrLf & _
        "--" & b & vbCrLf & "Content-Type: application/pdf; name=""" & nm & """" & vbCrLf & _
        "Content-Transfer-Encoding: base64" & vbCrLf & _
        "Content-Disposition: attachment; filename=""" & nm & """" & vbCrLf & vbCrLf & FileBase64(attach) & vbCrLf & _
        "--" & b & "--" & vbCrLf
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, t;
    Close #fn
End Sub

' ---------------------------------------------------------------- status changes
Public Function SetStatus(no As String, newStatus As String, ByVal d As Date) As String
    Dim r As Long, ws As Worksheet, cur As String, typ As String, e As String
    Set ws = Sh("Register")
    r = DocRow(no)
    If r = 0 Then SetStatus = "ERR: " & no & " not found": Exit Function
    cur = ws.Cells(r, R_STATUS).Value: typ = ws.Cells(r, R_TYPE).Value
    Select Case newStatus
        Case "Accepted", "Declined"
            If typ <> "Quote" Then
                e = "only a quote can be " & LCase$(newStatus)
            ElseIf cur <> "Sent" Then
                e = no & " is " & cur & "; only a sent quote can be " & LCase$(newStatus)
            ElseIf newStatus = "Accepted" And d > ws.Cells(r, R_DUE).Value Then
                e = no & " expired on " & DTxt(ws.Cells(r, R_DUE).Value) & "; issue a new quote"
            End If
        Case "Paid"
            If typ <> "Invoice" Then
                e = "only an invoice can be paid"
            ElseIf cur <> "Sent" Then
                e = no & " is " & cur & "; only a sent invoice can be paid"
            End If
        Case Else
            e = "unknown status " & newStatus
    End Select
    If e = "" And d < ws.Cells(r, R_DATE).Value Then e = "the date is before the document date"
    If e <> "" Then SetStatus = "ERR: " & e: Exit Function
    SheetsOpen
    ws.Cells(r, R_STATUS).Value = newStatus
    ws.Cells(r, R_DONE).Value = d
    LogIt newStatus, no, typ & " " & Money(ws.Cells(r, R_TOTAL).Value) & " " & LCase$(newStatus) & " on " & DTxt(d)
    SheetsLock
    SetStatus = ""
End Function

Public Function ConvertToInvoice(qno As String, ByVal d As Date) As String
    Dim r As Long, ws As Worksheet, inv As String, ri As Long
    Set ws = Sh("Register")
    r = DocRow(qno)
    If r = 0 Then ConvertToInvoice = "ERR: " & qno & " not found": Exit Function
    If ws.Cells(r, R_TYPE).Value <> "Quote" Then ConvertToInvoice = "ERR: " & qno & " is not a quote": Exit Function
    If ws.Cells(r, R_STATUS).Value <> "Accepted" Then ConvertToInvoice = "ERR: " & qno & " is " & ws.Cells(r, R_STATUS).Value & _
        "; only an accepted quote becomes an invoice": Exit Function
    If ws.Cells(r, R_LINK).Value <> "" Then ConvertToInvoice = "ERR: " & qno & " was already invoiced as " & _
        ws.Cells(r, R_LINK).Value: Exit Function
    ' the invoice copies the quote's own lines and prices: a later price-list change does not touch an agreed quote
    inv = NewDocument("Invoice", CStr(ws.Cells(r, R_CUST).Value), d, ws.Cells(r, R_DPCT).Value, _
                      "From quote " & qno, SpecOf(qno))
    If Left$(inv, 4) = "ERR:" Then ConvertToInvoice = inv: Exit Function
    ri = DocRow(inv)
    SheetsOpen
    ws.Cells(ri, R_LINK).Value = qno
    ws.Cells(r, R_LINK).Value = inv
    StoreCheck inv
    LogIt "Converted", qno, "Quote " & qno & " -> invoice " & inv
    SheetsLock
    ConvertToInvoice = inv
End Function

' ---------------------------------------------------------------- reset: an empty book for a new business
Public Sub ResetDocuments()
    Dim n As Variant, ws As Worksheet, last As Long
    SheetsOpen
    For Each n In Array("Register", "Lines", "Log")
        Set ws = Sh(CStr(n))
        last = LastRow(ws, 6)
        If last >= 6 Then ws.Rows("6:" & last).Delete
    Next
    LogIt "Reset", "", "All documents, lines and log entries cleared; customers, products and settings kept"
    Sh("Document").Range("CurDoc").Value = ""
    Sh("Home").Range("SelDoc").Value = ""
    RefreshDocList
    SheetsLock
End Sub
