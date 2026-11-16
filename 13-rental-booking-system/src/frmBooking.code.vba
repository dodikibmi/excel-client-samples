Option Explicit
' New booking: customer, dates, items. While you type, the form shows how many of each item are free for those
' dates (the same rule the core uses), the hire price, tax, total and deposit. "Confirm booking" holds the stock;
' the core checks everything again and refuses if anything does not fit.

Private codes() As String, qtys() As Long, n As Long

Public Sub StartNew()
    Dim ws As Worksheet, r As Long, cats As Object
    n = 0: ReDim codes(1 To MAX_LINES): ReDim qtys(1 To MAX_LINES)
    Set ws = Sh("Customers")
    cboCustomer.Clear
    For r = FIRST To LastRow(ws)
        cboCustomer.AddItem ws.Cells(r, C_ID).Value
        cboCustomer.List(cboCustomer.ListCount - 1, 1) = ws.Cells(r, C_NAME).Value
    Next
    Set ws = Sh("Catalog"): Set cats = CreateObject("Scripting.Dictionary")
    cboCategory.Clear
    For r = FIRST To LastRow(ws)
        If Not cats.Exists(CStr(ws.Cells(r, K_CAT).Value)) Then cats.Add CStr(ws.Cells(r, K_CAT).Value), 1: cboCategory.AddItem ws.Cells(r, K_CAT).Value
    Next
    cboItem.Clear
    txtEvent.Text = "": txtOut.Text = "": txtDue.Text = "": txtQty.Text = "": txtDelivery.Text = "0": txtNotes.Text = ""
    lblMsg.Caption = "": lblFree.Caption = ""
    lblUser.Caption = "Signed in: " & CurUser() & " (" & CurRole() & ")"
    Refresh
End Sub

Private Function Dates(ByRef dEv As Date, ByRef dOut As Date, ByRef dDue As Date) As Boolean
    Dim a As Variant, b As Variant, c As Variant
    a = ParseDate(txtEvent.Text): b = ParseDate(txtOut.Text): c = ParseDate(txtDue.Text)
    If IsEmpty(a) Or IsEmpty(b) Or IsEmpty(c) Then Dates = False: Exit Function
    dEv = a: dOut = b: dDue = c
    Dates = (dOut <= dEv And dEv <= dDue)
End Function

Private Sub cboCategory_Change()
    Dim ws As Worksheet, r As Long
    Set ws = Sh("Catalog")
    cboItem.Clear
    For r = FIRST To LastRow(ws)
        If ws.Cells(r, K_CAT).Value = cboCategory.Value Then
            cboItem.AddItem ws.Cells(r, K_CODE).Value
            cboItem.List(cboItem.ListCount - 1, 1) = ws.Cells(r, K_ITEM).Value
        End If
    Next
    ShowFree
End Sub

Private Sub cboItem_Change(): ShowFree: End Sub
Private Sub txtQty_Change(): ShowFree: End Sub
Private Sub txtEvent_Change(): Refresh: End Sub
Private Sub txtOut_Change(): Refresh: End Sub
Private Sub txtDue_Change(): Refresh: End Sub
Private Sub txtDelivery_Change(): Refresh: End Sub

Private Sub ShowFree()
    Dim dEv As Date, dOut As Date, dDue As Date, f As Long, k As Long
    If cboItem.ListIndex < 0 Then lblFree.Caption = "": Exit Sub
    k = ItemRow(cboItem.List(cboItem.ListIndex, 0))
    If Not Dates(dEv, dOut, dDue) Then lblFree.Caption = "type the dates to see what is free": Exit Sub
    f = MinFree(cboItem.List(cboItem.ListIndex, 0), dOut, dDue)
    If f < 0 Then f = 0
    lblFree.Caption = f & " of " & Sh("Catalog").Cells(k, K_OWNED).Value & " free for these dates  ·  " & _
        Money(Sh("Catalog").Cells(k, K_RATE).Value) & " a day" & _
        IIf(InList(cboItem.List(cboItem.ListIndex, 0)) > 0, "  ·  " & InList(cboItem.List(cboItem.ListIndex, 0)) & " already on this booking", "")
    lblFree.ForeColor = IIf(Val(txtQty.Text) + InList(cboItem.List(cboItem.ListIndex, 0)) > f, RGB(184, 58, 46), RGB(102, 113, 126))
End Sub

Private Function InList(code As String) As Long
    Dim i As Long
    For i = 1 To n
        If codes(i) = code Then InList = InList + qtys(i)
    Next
End Function

Public Sub AddLine()
    Dim q As Long, i As Long
    If cboItem.ListIndex < 0 Then lblMsg.Caption = "Pick an item": Exit Sub
    q = Val(txtQty.Text)
    If q <= 0 Then lblMsg.Caption = "Quantity must be above 0": Exit Sub
    For i = 1 To n
        If codes(i) = cboItem.List(cboItem.ListIndex, 0) Then qtys(i) = qtys(i) + q: GoTo shown
    Next
    If n = MAX_LINES Then lblMsg.Caption = "At most " & MAX_LINES & " lines": Exit Sub
    n = n + 1: codes(n) = cboItem.List(cboItem.ListIndex, 0): qtys(n) = q
shown:
    txtQty.Text = "": lblMsg.Caption = ""
    Refresh
End Sub

Public Sub RemoveLine()
    Dim i As Long, j As Long
    i = lstLines.ListIndex + 1
    If i < 1 Then Exit Sub
    For j = i To n - 1
        codes(j) = codes(j + 1): qtys(j) = qtys(j + 1)
    Next
    n = n - 1
    Refresh
End Sub

Public Function Spec() As String
    Dim i As Long
    For i = 1 To n
        Spec = Spec & IIf(i > 1, "|", "") & codes(i) & "~" & qtys(i)
    Next
End Function

Public Sub Refresh()
    Dim dEv As Date, dOut As Date, dDue As Date, ok As Boolean, days As Long, i As Long, k As Long
    Dim amt As Variant, rent As Variant, sbt As Variant, tax As Variant, f As Long, short As Long
    ok = Dates(dEv, dOut, dDue)
    days = IIf(ok, HireDays(dOut, dDue), 0)
    lblDays.Caption = IIf(ok, days & " hire day" & IIf(days = 1, "", "s") & "  ·  " & DTxt(dOut) & " to " & DTxt(dDue), _
        "dates as yyyy-mm-dd: out <= event <= back")
    lstLines.Clear
    rent = 0
    For i = 1 To n
        k = ItemRow(codes(i))
        amt = IIf(ok, R2(qtys(i) * Sh("Catalog").Cells(k, K_RATE).Value * days), 0)
        rent = rent + amt
        lstLines.AddItem codes(i)
        lstLines.List(i - 1, 1) = Sh("Catalog").Cells(k, K_ITEM).Value
        lstLines.List(i - 1, 2) = qtys(i)
        lstLines.List(i - 1, 3) = Money(Sh("Catalog").Cells(k, K_RATE).Value)
        lstLines.List(i - 1, 4) = Money(amt)
        If ok Then
            f = MinFree(codes(i), dOut, dDue)
            lstLines.List(i - 1, 5) = IIf(qtys(i) > f, "only " & IIf(f < 0, 0, f) & " free", f & " free")
            If qtys(i) > f Then short = short + 1
        End If
    Next
    sbt = rent + Val(txtDelivery.Text)
    tax = R2(sbt * S("TaxRate"))
    lblRent.Caption = Money(rent)
    lblTax.Caption = Money(tax)
    lblTotal.Caption = Money(sbt + tax)
    lblDeposit.Caption = Money(R2((sbt + tax) * S("DepositPct")))
    lblCheck.Caption = IIf(short > 0, short & " line(s) above what is free: Confirm will be refused", IIf(n > 0 And ok, "Every line fits the stock for these dates", ""))
    lblCheck.ForeColor = IIf(short > 0, RGB(184, 58, 46), RGB(47, 138, 78))
    ShowFree
End Sub

Public Function SaveNow(ByVal confirm As Boolean) As String
    Dim dEv As Date, dOut As Date, dDue As Date, res As String
    If cboCustomer.ListIndex < 0 Then lblMsg.Caption = "Pick a customer": SaveNow = "ERR": Exit Function
    If Not Dates(dEv, dOut, dDue) Then lblMsg.Caption = "Check the dates (yyyy-mm-dd, out <= event <= back)": SaveNow = "ERR": Exit Function
    res = NewBooking(cboCustomer.List(cboCustomer.ListIndex, 0), dEv, dOut, dDue, Spec(), Val(txtDelivery.Text), txtNotes.Text, confirm, DeskDate())
    Application.Calculate
    If Left$(res, 2) = "B-" Then
        Sh("Home").Unprotect: SetS "SelBooking", res: Sh("Home").Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
        Application.Calculate
        lblMsg.ForeColor = RGB(47, 138, 78)
        lblMsg.Caption = res & IIf(confirm, " confirmed: stock held", " saved as an enquiry")
    Else
        lblMsg.ForeColor = RGB(184, 58, 46)
        lblMsg.Caption = res
    End If
    SaveNow = res
End Function

Private Sub cmdAdd_Click(): AddLine: End Sub
Private Sub cmdRemove_Click(): RemoveLine: End Sub
Private Sub cmdEnquiry_Click(): If Left$(SaveNow(False), 2) = "B-" Then Unload Me
End Sub
Private Sub cmdConfirm_Click(): If Left$(SaveNow(True), 2) = "B-" Then Unload Me
End Sub
Private Sub cmdClose_Click(): Unload Me: End Sub
