Attribute VB_Name = "modCore"
Option Explicit
' Event rental desk: the core.
' Forms, Home buttons, the build script and the tests all call these functions, so every booking is checked against
' stock, priced, moved through its statuses and logged the same way. Functions return text:
'   NewBooking -> "B-2027-0001" | "BLOCKED: ..." | "ERR: ..."      other actions -> "" | "BLOCKED: ..." | "ERR: ..."
' Rules: sign in first; Cancel and Lock period need a Manager; nothing can be recorded on a date inside a locked
' period; an item can be booked only if (owned - written off - held by overlapping bookings) covers every day.
' Sheets are protected without a password (it stops mistakes, it is not security); the core unprotects, writes, protects.

Public Const FIRST As Long = 6, MAX_LINES As Long = 20
' Bookings columns
Public Const B_NO = 1, B_STATUS = 2, B_CUST = 3, B_NAME = 4, B_EVENT = 5, B_OUT = 6, B_DUE = 7, B_BACK = 8, B_DAYS = 9, _
    B_RENT = 10, B_DELIV = 11, B_LATE = 12, B_DMG = 13, B_SUB = 14, B_TAX = 15, B_TOTAL = 16, B_DEPDUE = 17, B_PAID = 18, _
    B_BAL = 19, B_INV = 20, B_INVDATE = 21, B_PDF = 22, B_BY = 23, B_AT = 24, B_NOTES = 25
' Lines columns (12-16 are formulas)
Public Const L_NO = 1, L_LINE = 2, L_CODE = 3, L_ITEM = 4, L_QTY = 5, L_RATE = 6, L_DAYS = 7, L_AMT = 8, L_LATEF = 9, _
    L_REPL = 10, L_DMG = 11
' Catalog / Customers / Users / Payments
Public Const K_CODE = 1, K_CAT = 2, K_ITEM = 3, K_OWNED = 4, K_RATE = 5, K_LATEF = 6, K_REPL = 7
Public Const C_ID = 1, C_NAME = 2
Public Const U_USER = 1, U_NAME = 2, U_ROLE = 3, U_HASH = 4
Public Const P_NO = 1, P_DATE = 2, P_KIND = 3, P_AMT = 4, P_BY = 5, P_AT = 6

Public ReplayClock As Date        ' set only by the build script when it replays the sample history

' ---------------------------------------------------------------- small helpers
Public Function S(nm As String) As Variant
    S = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Public Sub SetS(nm As String, v As Variant)
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

Public Function Sh(nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Public Function LastRow(ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < FIRST Then r = FIRST - 1
    LastRow = r
End Function

Public Function R2(ByVal x As Variant) As Variant      ' round half away from zero, like Excel's ROUND
    Dim d As Variant
    d = CDec(x) * 100
    If d >= 0 Then R2 = Fix(d + CDec(0.5)) / 100 Else R2 = Fix(d - CDec(0.5)) / 100
End Function

Public Function N2(ByVal x As Variant) As String       ' 1,234.50 on every PC (Format$ follows regional settings)
    Dim c As Variant, neg As Boolean, ip As String, out As String
    c = R2(x)
    If c < 0 Then neg = True: c = -c
    ip = Trim$(Str$(Fix(c)))
    Do While Len(ip) > 3
        out = "," & Right$(ip, 3) & out
        ip = Left$(ip, Len(ip) - 3)
    Loop
    N2 = IIf(neg, "-", "") & ip & out & "." & Right$("00" & Trim$(Str$(CLng((c - Fix(c)) * 100))), 2)
End Function

Public Function Money(ByVal x As Variant) As String
    Money = S("Cur") & N2(x)
End Function

Public Function DTxt(ByVal d As Date) As String
    DTxt = Day(d) & " " & Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", 3 * Month(d) - 2, 3) & " " & Year(d)
End Function

Public Function Stamp() As Date
    If ReplayClock <> 0 Then Stamp = ReplayClock Else Stamp = Now
End Function

Public Function TodayD() As Date                       ' the desk's "today": replay clock, else Settings As-of, else the PC date
    If ReplayClock <> 0 Then
        TodayD = Int(ReplayClock)
    ElseIf IsDate(S("AsOfDate")) And CStr(S("AsOfDate")) <> "" Then
        TodayD = CDate(S("AsOfDate"))
    Else
        TodayD = Date
    End If
End Function

Public Sub SetClock(ByVal iso As String)              ' build script only: "2027-01-04 09:00"; "" = real clock
    If iso = "" Then
        ReplayClock = 0
    Else
        ReplayClock = DateSerial(Val(Left$(iso, 4)), Val(Mid$(iso, 6, 2)), Val(Mid$(iso, 9, 2))) + _
            TimeSerial(Val(Mid$(iso, 12, 2)), Val(Mid$(iso, 15, 2)), 0)
    End If
End Sub

Public Function ParseDate(ByVal t As String) As Variant  ' "2027-03-20" -> date (never the PC's date order)
    t = Trim$(t)
    If t Like "####-##-##" Then
        ParseDate = DateSerial(Val(Left$(t, 4)), Val(Mid$(t, 6, 2)), Val(Right$(t, 2)))
        If Format$(ParseDate, "yyyymmdd") <> Replace(t, "-", "") Then ParseDate = Empty
    Else
        ParseDate = Empty
    End If
End Function

Public Sub SheetsOpen()
    Dim n As Variant
    For Each n In Array("Bookings", "Lines", "Payments", "Log", "Home", "Invoice", "Settings", "Users")
        Sh(CStr(n)).Unprotect
    Next
End Sub

Public Sub SheetsLock()
    Dim n As Variant
    For Each n In Array("Bookings", "Lines", "Payments", "Log", "Home", "Invoice", "Settings", "Users")
        Sh(CStr(n)).Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Next
End Sub

Public Function FindRow(ws As Worksheet, key As String) As Long
    Dim r As Variant
    r = Application.Match(key, ws.Range("A" & FIRST & ":A" & (LastRow(ws) + 1)), 0)
    If IsError(r) Then FindRow = 0 Else FindRow = FIRST + r - 1
End Function

Public Function BookRow(no As String) As Long: BookRow = FindRow(Sh("Bookings"), no): End Function
Public Function ItemRow(code As String) As Long: ItemRow = FindRow(Sh("Catalog"), code): End Function
Public Function CustRow(id As String) As Long: CustRow = FindRow(Sh("Customers"), id): End Function
Public Function UserRow(u As String) As Long: UserRow = FindRow(Sh("Users"), LCase$(u)): End Function

' ---------------------------------------------------------------- users, roles, audit trail
Public Function PinHash(ByVal user As String, ByVal pin As String) As String
    ' The workbook keeps only this number, never the PIN. A PIN keeps honest people honest: Excel is not a vault.
    Dim h As Double, t As String, i As Long, k As Long, v As Double
    t = LCase$(user) & ":" & pin & ":lumen-linen"
    h = 7
    For k = 1 To 300
        For i = 1 To Len(t)
            v = h * 131 + AscW(Mid$(t, i, 1)) + k
            h = v - Int(v / 2147483629#) * 2147483629#
        Next
    Next
    PinHash = Trim$(Str$(h))
End Function

Public Function CurUser() As String: CurUser = CStr(S("SignedUser")): End Function
Public Function CurRole() As String: CurRole = CStr(S("SignedRole")): End Function

Public Sub LogIt(action As String, no As String, result As String, detail As String)
    Dim ws As Worksheet, r As Long
    Set ws = Sh("Log")
    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = Stamp()
    ws.Cells(r, 2).Value = IIf(CurUser() = "", "(nobody)", CurUser())
    ws.Cells(r, 3).Value = CurRole()
    ws.Cells(r, 4).Value = action
    ws.Cells(r, 5).Value = no
    ws.Cells(r, 6).Value = result
    ws.Cells(r, 7).Value = detail
End Sub

Private Function Refuse(action As String, no As String, why As String) As String
    SheetsOpen
    LogIt action, no, "BLOCKED", why
    SetS "HomeMsg", "BLOCKED: " & why
    SheetsLock
    Refuse = "BLOCKED: " & why
End Function

Private Function Done(action As String, no As String, detail As String, Optional ret As String = "") As String
    LogIt action, no, "OK", detail
    SetS "HomeMsg", action & " " & no & ": " & detail
    SheetsLock
    Done = ret
End Function

Private Function Guard(action As String, no As String, ByVal d As Date, Optional managerOnly As Boolean = False) As String
    ' "" = allowed; otherwise the reason (already logged)
    If CurUser() = "" Then Guard = Refuse(action, no, "sign in first"): Exit Function
    If managerOnly And CurRole() <> "Manager" Then
        Guard = Refuse(action, no, action & " needs a Manager (" & CurUser() & " is " & CurRole() & ")"): Exit Function
    End If
    If IsDate(S("LockedThrough")) Then
        If d <= CDate(S("LockedThrough")) Then
            Guard = Refuse(action, no, "the period is closed through " & DTxt(CDate(S("LockedThrough"))) & " (date given: " & DTxt(d) & ")")
            Exit Function
        End If
    End If
    Guard = ""
End Function

Public Function SignIn(ByVal user As String, ByVal pin As String) As String
    Dim r As Long
    user = LCase$(Trim$(user))
    r = UserRow(user)
    SheetsOpen
    If r = 0 Then
        SetS "SignedUser", "": SetS "SignedRole", ""
        LogIt "Sign in", "", "BLOCKED", "unknown user '" & user & "'"
        SetS "HomeMsg", "BLOCKED: unknown user"
        Call SheetsLock: SignIn = "BLOCKED: unknown user": Exit Function
    End If
    If PinHash(user, pin) <> CStr(Sh("Users").Cells(r, U_HASH).Value) Then
        SetS "SignedUser", "": SetS "SignedRole", ""
        LogIt "Sign in", "", "BLOCKED", "wrong PIN for '" & user & "'"
        SetS "HomeMsg", "BLOCKED: wrong PIN"
        Call SheetsLock: SignIn = "BLOCKED: wrong PIN": Exit Function
    End If
    SetS "SignedUser", user
    SetS "SignedRole", Sh("Users").Cells(r, U_ROLE).Value
    SignIn = Done("Sign in", "", Sh("Users").Cells(r, U_NAME).Value & " (" & CurRole() & ")")
End Function

Public Sub ReplayAs(ByVal user As String)
    ' build script only: the sample history is replayed as the person who did each step (no sign-in line per step)
    Dim r As Long
    r = UserRow(user)
    Sh("Settings").Unprotect
    If r = 0 Then                                  ' (IIf would read Cells(0, ...) even when r = 0)
        SetS "SignedUser", "": SetS "SignedRole", ""
    Else
        SetS "SignedUser", LCase$(user): SetS "SignedRole", Sh("Users").Cells(r, U_ROLE).Value
    End If
End Sub

Public Sub SignOut()
    SheetsOpen
    If CurUser() <> "" Then LogIt "Sign out", "", "OK", ""
    SetS "SignedUser", "": SetS "SignedRole", ""
    SetS "HomeMsg", "Signed out"
    SheetsLock
End Sub

Public Function LockPeriod(ByVal through As Date, ByVal d As Date) As String
    Dim g As String
    g = Guard("Lock period", "", d, True)
    If g <> "" Then LockPeriod = g: Exit Function
    If through >= d Then LockPeriod = Refuse("Lock period", "", "can only lock dates before today (" & DTxt(d) & ")"): Exit Function
    SheetsOpen
    SetS "LockedThrough", through
    SetS "LockedAt", Stamp()
    LockPeriod = Done("Lock period", "", "closed through " & DTxt(through))
End Function

' ---------------------------------------------------------------- stock
' Which bookings hold stock, and until when: Confirmed / Out / Returned / Invoiced / Paid hold their items from the
' out date to the due date; a booking still Out after its due date keeps holding until it comes back (today).
Private Function HoldsStock(st As String) As Boolean
    HoldsStock = (st <> "Enquiry" And st <> "Cancelled" And st <> "")
End Function

Public Function FreeQty(code As String, ByVal d As Date, Optional excludeNo As String = "") As Long
    Dim wl As Worksheet, wb As Worksheet, k As Long, r As Long, br As Long, used As Long, lost As Long
    Dim st As String, dOut As Date, dEnd As Date
    Set wl = Sh("Lines"): Set wb = Sh("Bookings")
    k = ItemRow(code)
    If k = 0 Then FreeQty = 0: Exit Function
    For r = FIRST To LastRow(wl)
        If wl.Cells(r, L_CODE).Value = code Then
            br = BookRow(CStr(wl.Cells(r, L_NO).Value))
            st = CStr(wb.Cells(br, B_STATUS).Value)
            If wl.Cells(r, L_DMG).Value > 0 And IsDate(wb.Cells(br, B_BACK).Value) Then
                If CDate(wb.Cells(br, B_BACK).Value) <= d Then lost = lost + wl.Cells(r, L_DMG).Value
            End If
            If HoldsStock(st) And CStr(wl.Cells(r, L_NO).Value) <> excludeNo Then
                dOut = wb.Cells(br, B_OUT).Value: dEnd = wb.Cells(br, B_DUE).Value
                If st = "Out" And TodayD() > dEnd Then dEnd = TodayD()
                If dOut <= d And d <= dEnd Then used = used + wl.Cells(r, L_QTY).Value
            End If
        End If
    Next
    FreeQty = Sh("Catalog").Cells(k, K_OWNED).Value - lost - used
End Function

Public Function MinFree(code As String, ByVal dOut As Date, ByVal dDue As Date, Optional excludeNo As String = "", _
                        Optional ByRef worstDay As Date) As Long
    Dim d As Date, f As Long, m As Long
    m = 2147483647
    For d = dOut To dDue
        f = FreeQty(code, d, excludeNo)
        If f < m Then m = f: worstDay = d
    Next
    MinFree = m
End Function

' spec = "CODE~qty|CODE~qty" -> (code, qty) pairs, same code added up
Public Function ParseSpec(spec As String) As Collection
    Dim c As New Collection, d As Object, p As Variant, f As Variant, k As Variant
    Set d = CreateObject("Scripting.Dictionary")             ' keeps the order the items were added
    For Each p In Split(spec, "|")
        f = Split(p, "~")
        If UBound(f) >= 1 Then
            If d.Exists(Trim$(f(0))) Then d(Trim$(f(0))) = d(Trim$(f(0))) + CLng(Val(f(1))) Else d.Add Trim$(f(0)), CLng(Val(f(1)))
        End If
    Next
    For Each k In d.Keys
        c.Add Array(CStr(k), d(k))
    Next
    Set ParseSpec = c
End Function

Public Function StockProblem(spec As String, ByVal dOut As Date, ByVal dDue As Date, Optional excludeNo As String = "") As String
    Dim c As Collection, i As Long, f As Long, worst As Date
    Set c = ParseSpec(spec)
    For i = 1 To c.Count
        f = MinFree(CStr(c(i)(0)), dOut, dDue, excludeNo, worst)
        If c(i)(1) > f Then
            StockProblem = Sh("Catalog").Cells(ItemRow(CStr(c(i)(0))), K_ITEM).Value & ": " & c(i)(1) & " asked, only " & _
                IIf(f < 0, 0, f) & " free on " & DTxt(worst)
            Exit Function
        End If
    Next
    StockProblem = ""
End Function

Public Function InputProblem(custId As String, ByVal dEvent As Date, ByVal dOut As Date, ByVal dDue As Date, spec As String) As String
    Dim c As Collection, i As Long
    If CustRow(custId) = 0 Then InputProblem = "pick a customer": Exit Function
    If Not (dOut <= dEvent And dEvent <= dDue) Then InputProblem = "dates must run out <= event <= back": Exit Function
    If dDue - dOut > S("MaxDays") Then InputProblem = "a hire can last " & S("MaxDays") & " days at most": Exit Function
    Set c = ParseSpec(spec)
    If c.Count = 0 Then InputProblem = "add at least one item": Exit Function
    If c.Count > MAX_LINES Then InputProblem = "at most " & MAX_LINES & " lines": Exit Function
    For i = 1 To c.Count
        If ItemRow(CStr(c(i)(0))) = 0 Then InputProblem = "unknown item " & c(i)(0): Exit Function
        If c(i)(1) <= 0 Then InputProblem = "quantity must be above 0": Exit Function
    Next
    InputProblem = ""
End Function

' ---------------------------------------------------------------- money
Public Function HireDays(ByVal dOut As Date, ByVal dDue As Date) As Long
    HireDays = IIf(dDue - dOut < 1, 1, dDue - dOut)
End Function

Public Sub Reprice(no As String)
    ' totals of one booking from its lines, fees and payments (Bookings values are frozen numbers, not formulas)
    Dim wb As Worksheet, wl As Worksheet, wp As Worksheet, r As Long, i As Long
    Dim rent As Variant, late As Variant, dmg As Variant, sbt As Variant, paid As Variant, lateDays As Long
    Set wb = Sh("Bookings"): Set wl = Sh("Lines"): Set wp = Sh("Payments")
    r = BookRow(no)
    If IsDate(wb.Cells(r, B_BACK).Value) Then
        lateDays = CDate(wb.Cells(r, B_BACK).Value) - CDate(wb.Cells(r, B_DUE).Value) - S("GraceDays")
        If lateDays < 0 Then lateDays = 0
    End If
    rent = 0: late = 0: dmg = 0
    For i = FIRST To LastRow(wl)
        If wl.Cells(i, L_NO).Value = no Then
            rent = rent + wl.Cells(i, L_AMT).Value
            late = late + R2(wl.Cells(i, L_QTY).Value * wl.Cells(i, L_LATEF).Value * lateDays)
            dmg = dmg + R2(wl.Cells(i, L_DMG).Value * wl.Cells(i, L_REPL).Value)
        End If
    Next
    paid = 0
    For i = FIRST To LastRow(wp)
        If wp.Cells(i, P_NO).Value = no Then paid = paid + wp.Cells(i, P_AMT).Value
    Next
    sbt = R2(rent + wb.Cells(r, B_DELIV).Value + late + dmg)      ' every stored amount rounded to the cent (no 1E-13 leftovers)
    wb.Cells(r, B_RENT).Value = R2(rent)
    wb.Cells(r, B_LATE).Value = R2(late)
    wb.Cells(r, B_DMG).Value = R2(dmg)
    wb.Cells(r, B_SUB).Value = sbt
    wb.Cells(r, B_TAX).Value = R2(sbt * S("TaxRate"))
    wb.Cells(r, B_TOTAL).Value = R2(sbt + R2(sbt * S("TaxRate")))
    wb.Cells(r, B_PAID).Value = R2(paid)
    wb.Cells(r, B_BAL).Value = R2(sbt + R2(sbt * S("TaxRate")) - paid)
End Sub

Public Function NextNo(prefix As String, ByVal d As Date) As String
    Dim ws As Worksheet, col As Long, r As Long, t As String, n As Long, best As Long, stem As String
    If prefix = "B" Then Set ws = Sh("Bookings"): col = B_NO Else Set ws = Sh("Bookings"): col = B_INV
    stem = prefix & "-" & Year(d) & "-"
    For r = FIRST To LastRow(ws)
        t = CStr(ws.Cells(r, col).Value)
        If Left$(t, Len(stem)) = stem Then
            n = Val(Mid$(t, Len(stem) + 1))
            If n > best Then best = n
        End If
    Next
    NextNo = stem & Format$(best + 1, "0000")
End Function

' ---------------------------------------------------------------- actions
Public Function NewBooking(custId As String, ByVal dEvent As Date, ByVal dOut As Date, ByVal dDue As Date, spec As String, _
                           ByVal delivery As Double, notes As String, ByVal confirm As Boolean, ByVal d As Date) As String
    Dim g As String, e As String, no As String, wb As Worksheet, wl As Worksheet, r As Long, rl As Long, c As Collection
    Dim i As Long, k As Long, days As Long, est As Variant
    g = Guard(IIf(confirm, "New booking", "New enquiry"), "", d)
    If g <> "" Then NewBooking = g: Exit Function
    e = InputProblem(custId, dEvent, dOut, dDue, spec)
    If e <> "" Then NewBooking = Refuse("New booking", "", e): Exit Function
    If confirm Then
        e = StockProblem(spec, dOut, dDue)
        If e <> "" Then NewBooking = Refuse("New booking", "", e): Exit Function
    End If
    Set wb = Sh("Bookings"): Set wl = Sh("Lines")
    SheetsOpen
    no = NextNo("B", d)
    r = LastRow(wb) + 1
    days = HireDays(dOut, dDue)
    wb.Cells(r, B_NO).Value = no
    wb.Cells(r, B_STATUS).Value = IIf(confirm, "Confirmed", "Enquiry")
    wb.Cells(r, B_CUST).Value = custId
    wb.Cells(r, B_NAME).Value = Sh("Customers").Cells(CustRow(custId), C_NAME).Value
    wb.Cells(r, B_EVENT).Value = dEvent: wb.Cells(r, B_OUT).Value = dOut: wb.Cells(r, B_DUE).Value = dDue
    wb.Cells(r, B_DAYS).Value = days
    wb.Cells(r, B_DELIV).Value = delivery
    wb.Cells(r, B_BY).Value = CurUser(): wb.Cells(r, B_AT).Value = Stamp(): wb.Cells(r, B_NOTES).Value = notes
    Set c = ParseSpec(spec)
    rl = LastRow(wl)
    For i = 1 To c.Count
        k = ItemRow(CStr(c(i)(0)))
        rl = rl + 1
        wl.Cells(rl, L_NO).Value = no: wl.Cells(rl, L_LINE).Value = i
        wl.Cells(rl, L_CODE).Value = c(i)(0)
        wl.Cells(rl, L_ITEM).Value = Sh("Catalog").Cells(k, K_ITEM).Value
        wl.Cells(rl, L_QTY).Value = c(i)(1)
        wl.Cells(rl, L_RATE).Value = Sh("Catalog").Cells(k, K_RATE).Value
        wl.Cells(rl, L_DAYS).Value = days
        wl.Cells(rl, L_AMT).Value = R2(c(i)(1) * wl.Cells(rl, L_RATE).Value * days)
        wl.Cells(rl, L_LATEF).Value = Sh("Catalog").Cells(k, K_LATEF).Value
        wl.Cells(rl, L_REPL).Value = Sh("Catalog").Cells(k, K_REPL).Value
        wl.Cells(rl, L_DMG).Value = 0
        LineFormulas rl
    Next
    Reprice no
    est = wb.Cells(r, B_TOTAL).Value
    wb.Cells(r, B_DEPDUE).Value = R2(est * S("DepositPct"))
    Reprice no
    NewBooking = Done(IIf(confirm, "New booking", "New enquiry"), no, wb.Cells(r, B_NAME).Value & ", " & c.Count & " lines, " & _
        DTxt(dOut) & " to " & DTxt(dDue) & ", " & Money(est) & IIf(confirm, ", stock held", ", no stock held yet"), no)
End Function

Public Sub LineFormulas(rl As Long)
    ' helpers for the Availability sheet and the checks: status, out date, end of hold, holds stock?, back date
    Dim wl As Worksheet, m As String
    Set wl = Sh("Lines")
    m = "MATCH($A" & rl & ",Bookings!$A:$A,0)"
    wl.Cells(rl, 12).Formula = "=INDEX(Bookings!$B:$B," & m & ")"
    wl.Cells(rl, 13).Formula = "=INDEX(Bookings!$F:$F," & m & ")"
    wl.Cells(rl, 14).Formula = "=IF($L" & rl & "=""Out"",MAX(INDEX(Bookings!$G:$G," & m & "),AsOf),INDEX(Bookings!$G:$G," & m & "))"
    wl.Cells(rl, 15).Formula = "=IF(OR($L" & rl & "=""Enquiry"",$L" & rl & "=""Cancelled""),0,1)"
    wl.Cells(rl, 16).Formula = "=IF(INDEX(Bookings!$H:$H," & m & ")="""","""",INDEX(Bookings!$H:$H," & m & "))"
End Sub

Private Function NeedStatus(action As String, no As String, r As Long, ok As String) As String
    Dim st As String
    If r = 0 Then NeedStatus = Refuse(action, no, "no booking " & no): Exit Function
    st = CStr(Sh("Bookings").Cells(r, B_STATUS).Value)
    If InStr("," & ok & ",", "," & st & ",") = 0 Then
        NeedStatus = Refuse(action, no, "booking is " & st & " (needs " & Replace(ok, ",", " or ") & ")")
    Else
        NeedStatus = ""
    End If
End Function

Public Function ConfirmBooking(no As String, ByVal d As Date) As String
    Dim g As String, r As Long, e As String, wb As Worksheet
    g = Guard("Confirm", no, d): If g <> "" Then ConfirmBooking = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Confirm", no, r, "Enquiry"): If g <> "" Then ConfirmBooking = g: Exit Function
    Set wb = Sh("Bookings")
    e = StockProblem(SpecOf(no), wb.Cells(r, B_OUT).Value, wb.Cells(r, B_DUE).Value, no)
    If e <> "" Then ConfirmBooking = Refuse("Confirm", no, e): Exit Function
    SheetsOpen
    wb.Cells(r, B_STATUS).Value = "Confirmed"
    ConfirmBooking = Done("Confirm", no, "stock held " & DTxt(wb.Cells(r, B_OUT).Value) & " to " & DTxt(wb.Cells(r, B_DUE).Value))
End Function

Public Function SpecOf(no As String) As String
    Dim wl As Worksheet, i As Long, t As String
    Set wl = Sh("Lines")
    For i = FIRST To LastRow(wl)
        If wl.Cells(i, L_NO).Value = no Then t = t & IIf(t = "", "", "|") & wl.Cells(i, L_CODE).Value & "~" & wl.Cells(i, L_QTY).Value
    Next
    SpecOf = t
End Function

Public Function CheckOut(no As String, ByVal d As Date) As String
    Dim g As String, r As Long
    g = Guard("Check out", no, d): If g <> "" Then CheckOut = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Check out", no, r, "Confirmed"): If g <> "" Then CheckOut = g: Exit Function
    SheetsOpen
    Sh("Bookings").Cells(r, B_STATUS).Value = "Out"
    CheckOut = Done("Check out", no, Sh("Bookings").Cells(r, B_NAME).Value & ", due back " & DTxt(Sh("Bookings").Cells(r, B_DUE).Value))
End Function

Public Function CheckIn(no As String, ByVal d As Date, dmgSpec As String) As String
    ' dmgSpec = "CODE~qty damaged or missing|..." ("" = all back fine). Damaged / missing items are charged at
    ' replacement cost and leave the stock from this date.
    Dim g As String, r As Long, wl As Worksheet, c As Collection, i As Long, j As Long, ok As Boolean, wb As Worksheet, txt As String
    g = Guard("Check in", no, d): If g <> "" Then CheckIn = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Check in", no, r, "Out"): If g <> "" Then CheckIn = g: Exit Function
    Set wl = Sh("Lines"): Set wb = Sh("Bookings")
    Set c = ParseSpec(dmgSpec)
    For i = 1 To c.Count                                   ' every damaged item must be on the booking, not more than hired
        ok = False
        For j = FIRST To LastRow(wl)
            If wl.Cells(j, L_NO).Value = no And wl.Cells(j, L_CODE).Value = c(i)(0) Then ok = (c(i)(1) >= 0 And c(i)(1) <= wl.Cells(j, L_QTY).Value)
        Next
        If Not ok Then CheckIn = Refuse("Check in", no, "damaged " & c(i)(0) & " x" & c(i)(1) & " is not on this booking or more than hired"): Exit Function
    Next
    SheetsOpen
    For j = FIRST To LastRow(wl)
        If wl.Cells(j, L_NO).Value = no Then
            wl.Cells(j, L_DMG).Value = 0
            For i = 1 To c.Count
                If wl.Cells(j, L_CODE).Value = c(i)(0) Then wl.Cells(j, L_DMG).Value = c(i)(1)
            Next
        End If
    Next
    wb.Cells(r, B_BACK).Value = d
    wb.Cells(r, B_STATUS).Value = "Returned"
    Reprice no
    txt = IIf(wb.Cells(r, B_LATE).Value > 0, "late fee " & Money(wb.Cells(r, B_LATE).Value) & " (" & (d - wb.Cells(r, B_DUE).Value) & " days late)", "on time")
    If wb.Cells(r, B_DMG).Value > 0 Then txt = txt & ", damaged/missing " & Money(wb.Cells(r, B_DMG).Value)
    CheckIn = Done("Check in", no, txt)
End Function

Public Function FolderOf(nm As String) As String
    Dim f As String
    f = CStr(S(nm))
    If InStr(f, ":") = 0 And Left$(f, 2) <> "\\" Then f = ThisWorkbook.Path & "\" & f
    If Dir(f, vbDirectory) = "" Then MkDir f
    FolderOf = f
End Function

Public Function SafeName(ByVal s0 As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "&", "'", ".")
        s0 = Replace(s0, ch, " ")
    Next
    SafeName = Application.WorksheetFunction.Trim(s0)
End Function

Public Function MakeInvoice(no As String, ByVal d As Date) As String
    Dim g As String, r As Long, wb As Worksheet, inv As String, f As String
    g = Guard("Invoice", no, d): If g <> "" Then MakeInvoice = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Invoice", no, r, "Returned"): If g <> "" Then MakeInvoice = g: Exit Function
    Set wb = Sh("Bookings")
    SheetsOpen
    inv = NextNo("INV", d)
    wb.Cells(r, B_INV).Value = inv: wb.Cells(r, B_INVDATE).Value = d
    Reprice no
    wb.Cells(r, B_STATUS).Value = IIf(wb.Cells(r, B_BAL).Value <= 0, "Paid", "Invoiced")
    f = inv & " " & SafeName(CStr(wb.Cells(r, B_NAME).Value)) & ".pdf"
    SetS "CurInv", no
    Application.Calculate
    On Error GoTo failed
    Sh("Invoice").ExportAsFixedFormat Type:=xlTypePDF, Filename:=FolderOf("PdfFolder") & "\" & f, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0
    wb.Cells(r, B_PDF).Value = f
    MakeInvoice = Done("Invoice", no, inv & " " & Money(wb.Cells(r, B_TOTAL).Value) & ", balance " & Money(wb.Cells(r, B_BAL).Value) & ", PDF " & f)
    Exit Function
failed:
    wb.Cells(r, B_INV).Value = "": wb.Cells(r, B_INVDATE).Value = "": wb.Cells(r, B_STATUS).Value = "Returned"
    LogIt "Invoice", no, "ERR", "PDF not written: " & Err.Description
    SheetsLock
    MakeInvoice = "ERR: PDF not written: " & Err.Description
End Function

Public Function RecordPayment(no As String, ByVal amount As Double, ByVal d As Date) As String
    Dim g As String, r As Long, wb As Worksheet, wp As Worksheet, rp As Long, kind As String
    g = Guard("Payment", no, d): If g <> "" Then RecordPayment = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Payment", no, r, "Confirmed,Out,Returned,Invoiced"): If g <> "" Then RecordPayment = g: Exit Function
    Set wb = Sh("Bookings"): Set wp = Sh("Payments")
    If amount <= 0 Then RecordPayment = Refuse("Payment", no, "amount must be above 0"): Exit Function
    If R2(amount) > R2(wb.Cells(r, B_BAL).Value) Then
        RecordPayment = Refuse("Payment", no, Money(amount) & " is more than the balance " & Money(wb.Cells(r, B_BAL).Value)): Exit Function
    End If
    kind = IIf(wb.Cells(r, B_INV).Value = "", "Deposit", "Balance")
    SheetsOpen
    rp = LastRow(wp) + 1
    wp.Cells(rp, P_NO).Value = no: wp.Cells(rp, P_DATE).Value = d: wp.Cells(rp, P_KIND).Value = kind
    wp.Cells(rp, P_AMT).Value = R2(amount): wp.Cells(rp, P_BY).Value = CurUser(): wp.Cells(rp, P_AT).Value = Stamp()
    Reprice no
    If wb.Cells(r, B_STATUS).Value = "Invoiced" And wb.Cells(r, B_BAL).Value <= 0 Then wb.Cells(r, B_STATUS).Value = "Paid"
    RecordPayment = Done("Payment", no, kind & " " & Money(amount) & ", balance " & Money(wb.Cells(r, B_BAL).Value))
End Function

Public Function CancelBooking(no As String, ByVal d As Date, reason As String) As String
    Dim g As String, r As Long
    g = Guard("Cancel", no, d, True): If g <> "" Then CancelBooking = g: Exit Function
    r = BookRow(no)
    g = NeedStatus("Cancel", no, r, "Enquiry,Confirmed"): If g <> "" Then CancelBooking = g: Exit Function
    If Trim$(reason) = "" Then CancelBooking = Refuse("Cancel", no, "give a reason"): Exit Function
    SheetsOpen
    Sh("Bookings").Cells(r, B_STATUS).Value = "Cancelled"
    Sh("Bookings").Cells(r, B_NOTES).Value = Trim$(Sh("Bookings").Cells(r, B_NOTES).Value & " Cancelled: " & reason)
    CancelBooking = Done("Cancel", no, "stock released; reason: " & reason)
End Function

Public Function DeskDate() As Date                      ' Home's "Action date" if typed, else today
    If IsDate(S("ActionDate")) And CStr(S("ActionDate")) <> "" Then DeskDate = CDate(S("ActionDate")) Else DeskDate = TodayD()
End Function
