Option Explicit
' New quote / edit a draft. The form only collects input: totals, checks, numbering, saving and issuing are done by
' modCore, the same code the Home buttons and the automated test use.

Private mNo As String        ' "" = not saved yet
Private mType As String
Private mLines As Collection ' each item: Array(code, description, unit, qty, price, disc)
Private mLoading As Boolean

Private Const GOOD As Long = &H8ABF5B   ' RGB(91,191,138)
Private Const BAD As Long = &H5F73E5    ' RGB(229,115,95)
Private Const MUTED As Long = &HA3988E  ' RGB(142,152,163)

Public Sub StartNew(docType As String)
    mNo = "": mType = docType
    Set mLines = New Collection
    mLoading = True
    cboCustomer.ListIndex = -1
    txtDate.Value = Format$(Date, "yyyy-mm-dd")
    txtDocDisc.Value = "0"
    txtNotes.Value = ""
    mLoading = False
    lblTitle.Caption = "New " & LCase$(docType)
    lblNo.Caption = "Number given when saved"
    lblMsg.Caption = ""
    cmdClose.Caption = "Cancel"
    Refresh
End Sub

Public Sub LoadDraft(no As String)
    Dim r As Long, ws As Worksheet, L As Variant, i As Long, pr As Long
    Set ws = Sh("Register")
    r = DocRow(no)
    mNo = no: mType = ws.Cells(r, R_TYPE).Value
    Set mLines = New Collection
    mLoading = True
    SetCustomer CStr(ws.Cells(r, R_CUST).Value)
    txtDate.Value = Format$(ws.Cells(r, R_DATE).Value, "yyyy-mm-dd")
    txtDocDisc.Value = NumText(ws.Cells(r, R_DPCT).Value * 100)
    txtNotes.Value = ws.Cells(r, R_NOTES).Value
    L = ParseSpec(SpecOf(no))
    If Not IsEmpty(L) Then
        For i = 0 To UBound(L)
            pr = ProdRow(CStr(L(i)(0)))
            mLines.Add Array(L(i)(0), Sh("Products").Cells(pr, P_DESC).Value, Sh("Products").Cells(pr, P_UNIT).Value, _
                             L(i)(1), L(i)(2), L(i)(3))
        Next
    End If
    mLoading = False
    lblTitle.Caption = "Edit draft " & LCase$(mType)
    lblNo.Caption = no & "  ·  Draft"
    lblMsg.Caption = ""
    cmdClose.Caption = "Cancel"
    Refresh
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, r As Long, cats As New Collection, c As Variant
    Set ws = Sh("Customers")
    cboCustomer.Clear
    For r = 6 To LastRow(ws, 6)
        cboCustomer.AddItem ws.Cells(r, C_ID).Value & "  ·  " & ws.Cells(r, C_NAME).Value
    Next
    Set ws = Sh("Products")
    On Error Resume Next
    For r = 6 To LastRow(ws, 6)
        cats.Add CStr(ws.Cells(r, P_CAT).Value), CStr(ws.Cells(r, P_CAT).Value)
    Next
    On Error GoTo 0
    cboCategory.Clear
    For Each c In cats
        cboCategory.AddItem c
    Next
    txtQty.Value = "1": txtDisc.Value = "0"
    If mLines Is Nothing Then StartNew "Quote"
End Sub

' ---------------------------------------------------------------- customer
Public Function CustId() As String
    If cboCustomer.ListIndex >= 0 Then CustId = Trim$(Split(cboCustomer.Value, "·")(0))
End Function

Public Sub SetCustomer(id As String)
    Dim i As Long
    For i = 0 To cboCustomer.ListCount - 1
        If Trim$(Split(cboCustomer.List(i), "·")(0)) = id Then cboCustomer.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub cboCustomer_Change()
    Dim cr As Long, C As Worksheet, em As String
    cr = CustRow(CustId())
    If cr = 0 Then lblCust.Caption = "": Refresh: Exit Sub
    Set C = Sh("Customers")
    em = Trim$(C.Cells(cr, C_EMAIL).Value)
    lblCust.Caption = C.Cells(cr, C_CONTACT).Value & "  ·  " & IIf(em = "", "NO EMAIL ADDRESS", em) & "  ·  " & _
        C.Cells(cr, C_CITY).Value & "  ·  pays in " & C.Cells(cr, C_TERMS).Value & " days" & _
        IIf(C.Cells(cr, C_EXEMPT).Value = "Yes", "  ·  tax exempt", "")
    lblCust.ForeColor = IIf(em = "", BAD, MUTED)
    Refresh
End Sub

' ---------------------------------------------------------------- category -> product (cascading)
Private Sub cboCategory_Change()
    Dim ws As Worksheet, r As Long
    Set ws = Sh("Products")
    cboProduct.Clear
    For r = 6 To LastRow(ws, 6)
        If ws.Cells(r, P_CAT).Value = cboCategory.Value Then cboProduct.AddItem ws.Cells(r, P_CODE).Value & "  ·  " & ws.Cells(r, P_DESC).Value
    Next
    lblProd.Caption = ""
End Sub

Public Function ProdCode() As String
    If cboProduct.ListIndex >= 0 Then ProdCode = Trim$(Split(cboProduct.Value, "·")(0))
End Function

Private Sub cboProduct_Change()
    Dim pr As Long, P As Worksheet
    pr = ProdRow(ProdCode())
    If pr = 0 Then lblProd.Caption = "": Exit Sub
    Set P = Sh("Products")
    txtPrice.Value = N2(P.Cells(pr, P_LIST).Value, False)
    txtDisc.Value = "0"
    lblProd.Caption = "per " & P.Cells(pr, P_UNIT).Value & "  ·  list " & N2(P.Cells(pr, P_LIST).Value, False) & _
        "  ·  minimum " & N2(P.Cells(pr, P_MIN).Value, False) & "  ·  max discount " & P.Cells(pr, P_MAXD).Value & "%"
    lblProd.ForeColor = MUTED
End Sub

Private Function Num(ByVal t As String, ByRef v As Variant) As Boolean
    t = Trim$(Replace(t, ",", ""))
    If t = "" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    v = CDec(Val(t))
    Num = True
End Function

Public Sub AddLine()
    Dim q As Variant, p As Variant, d As Variant, pr As Long
    pr = ProdRow(ProdCode())
    If pr = 0 Then Say "Pick a category and a product first.", BAD: Exit Sub
    If Not Num(txtQty.Value, q) Then Say "Quantity must be a number.", BAD: Exit Sub
    If q <= 0 Then Say "Quantity must be above 0.", BAD: Exit Sub
    If Not Num(txtPrice.Value, p) Then Say "Unit price must be a number.", BAD: Exit Sub
    If p < 0 Then Say "Unit price cannot be negative.", BAD: Exit Sub
    If Not Num(txtDisc.Value, d) Then d = 0
    If d < 0 Or d >= 100 Then Say "Line discount must be 0-99%.", BAD: Exit Sub
    If mLines.Count >= MAX_LINES Then Say "At most " & MAX_LINES & " lines per document.", BAD: Exit Sub
    mLines.Add Array(ProdCode(), Sh("Products").Cells(pr, P_DESC).Value, Sh("Products").Cells(pr, P_UNIT).Value, q, p, d / 100)
    Say "Line added.", GOOD
    Refresh
End Sub

Public Sub RemoveLine()
    If lstLines.ListIndex < 0 Then Say "Click a line in the list first.", BAD: Exit Sub
    mLines.Remove lstLines.ListIndex + 1
    Say "Line removed.", GOOD
    Refresh
End Sub

Private Sub cmdAdd_Click(): AddLine: End Sub
Private Sub cmdRemove_Click(): RemoveLine: End Sub
Private Sub txtDocDisc_Change(): If Not mLoading Then Refresh
End Sub

Public Function DocDisc() As Variant
    Dim v As Variant
    If Num(txtDocDisc.Value, v) Then DocDisc = v / 100 Else DocDisc = CDec(0)
End Function

Public Function Spec() As String
    Dim i As Long, s0 As String, L As Variant
    For i = 1 To mLines.Count
        L = mLines(i)
        If i > 1 Then s0 = s0 & "|"
        s0 = s0 & L(0) & "~" & NumText(L(3)) & "~" & NumText(L(4)) & "~" & NumText(L(5))
    Next
    Spec = s0
End Function

Public Function DocDate() As Date
    Dim t As String
    t = Trim$(txtDate.Value)
    If t Like "####-##-##" Then DocDate = DateSerial(Val(Left$(t, 4)), Val(Mid$(t, 6, 2)), Val(Right$(t, 2)))
End Function

' ---------------------------------------------------------------- live list, totals and checks
Public Sub Refresh()
    Dim i As Long, L As Variant, t As Variant, chk As String
    If mLoading Or mLines Is Nothing Then Exit Sub
    lstLines.Clear
    For i = 1 To mLines.Count
        L = mLines(i)
        lstLines.AddItem i
        lstLines.List(i - 1, 1) = L(0)
        lstLines.List(i - 1, 2) = L(1)
        lstLines.List(i - 1, 3) = NumText(L(3)) & " " & L(2)
        lstLines.List(i - 1, 4) = N2(L(4))
        lstLines.List(i - 1, 5) = IIf(L(5) = 0, "", Pct(L(5)))
        lstLines.List(i - 1, 6) = N2(LineNet(L(3), L(4), L(5)))
    Next
    t = Totals(CustId(), DocDisc(), Spec())
    lblSub.Caption = Money(t(0))
    lblDiscCap.Caption = "Discount " & Pct(DocDisc())
    lblDiscAmt.Caption = IIf(t(1) = 0, "-", "-" & Money(t(1)))
    lblTaxCap.Caption = S("TaxLabel") & " " & Pct(t(2))
    lblTax.Caption = Money(t(3))
    lblTotal.Caption = Money(t(4))
    chk = CheckSpec(mType, CustId(), DocDisc(), Spec())
    If DocDate() = 0 Then chk = "Date must be yyyy-mm-dd." & IIf(chk = "", "", vbLf & chk)
    If chk = "" Then
        lblChecks.Caption = "Ready to issue: every check passed."
        lblChecks.ForeColor = GOOD
    Else
        lblChecks.Caption = "Not ready to issue yet (you can still save a draft):" & vbLf & "- " & Replace(chk, vbLf, vbLf & "- ")
        lblChecks.ForeColor = BAD
    End If
End Sub

Private Sub Say(t As String, col As Long)
    lblMsg.Caption = t
    lblMsg.ForeColor = col
End Sub

' ---------------------------------------------------------------- save / issue
Public Function SaveDraft() As Boolean
    Dim res As String
    If DocDate() = 0 Then Say "Date must be yyyy-mm-dd.", BAD: Exit Function
    If mNo = "" Then
        res = NewDocument(mType, CustId(), DocDate(), DocDisc(), CStr(txtNotes.Value), Spec())
        If Left$(res, 4) = "ERR:" Then Say Mid$(res, 6), BAD: Exit Function
        mNo = res
    Else
        res = UpdateDraft(mNo, CustId(), DocDate(), DocDisc(), CStr(txtNotes.Value), Spec())
        If res <> "" Then Say Mid$(res, 6), BAD: Exit Function
    End If
    lblNo.Caption = mNo & "  ·  Draft"
    lblTitle.Caption = "Edit draft " & LCase$(mType)
    Sh("Home").Range("SelDoc").Value = mNo
    Say "Saved " & mNo & " as a draft.", GOOD
    SaveDraft = True
End Function

Public Sub IssueNow()
    Dim res As String
    If Not SaveDraft() Then Exit Sub
    res = IssueDocument(mNo, Date)
    If res = "" Then
        lblNo.Caption = mNo & "  ·  Sent"
        lblTitle.Caption = mType & " issued"
        Say mNo & " issued: PDF saved" & IIf(S("EmailMode") = "Off", ".", ", email ready (" & S("EmailMode") & ")."), GOOD
        cmdSave.Enabled = False: cmdIssue.Enabled = False: cmdAdd.Enabled = False: cmdRemove.Enabled = False
        cmdClose.Caption = "Close"
    ElseIf Left$(res, 8) = "BLOCKED:" Then
        Say "Saved as a draft, NOT issued: fix the points above first.", BAD
    Else
        Say Mid$(res, 6), BAD
    End If
End Sub

Private Sub cmdSave_Click(): SaveDraft: End Sub
Private Sub cmdIssue_Click(): IssueNow: End Sub
Private Sub cmdClose_Click(): Unload Me: End Sub
