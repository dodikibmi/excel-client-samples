Attribute VB_Name = "modButtons"
Option Explicit
' Home buttons. Each one calls the core with the booking picked on Home and the action date (blank = today); the
' result is written to the message bar on Home (no pop-ups), and every attempt, allowed or not, is in the Log.

Private Function Sel() As String
    Sel = Trim$(CStr(S("SelBooking")))
End Function

Private Sub ShowResult(res As String)
    Application.Calculate
    If res <> "" And Left$(res, 2) <> "B-" Then SetMsg res
End Sub

Public Sub SetMsg(t As String)
    Sh("Home").Unprotect
    SetS "HomeMsg", t
    Sh("Home").Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
End Sub

Public Sub BtnSignIn(): frmSignIn.Show: End Sub
Public Sub BtnSignOut(): SignOut: Application.Calculate: End Sub

Public Sub BtnNew()
    If CurUser() = "" Then SetMsg "BLOCKED: sign in first": Exit Sub
    frmBooking.StartNew
    frmBooking.Show
End Sub

Public Sub BtnConfirm(): ShowResult ConfirmBooking(Sel(), DeskDate()): End Sub
Public Sub BtnCheckOut(): ShowResult CheckOut(Sel(), DeskDate()): End Sub

Public Sub BtnCheckIn()
    ' damaged / missing quantities typed next to the booking's lines on Home
    Dim codes As Range, qty As Range, i As Long, spec As String
    Set codes = ThisWorkbook.Names("HomeLineCode").RefersToRange
    Set qty = ThisWorkbook.Names("HomeDamaged").RefersToRange
    For i = 1 To codes.Rows.Count
        If CStr(codes.Cells(i, 1).Value) <> "" And Val(CStr(qty.Cells(i, 1).Value)) > 0 Then
            spec = spec & IIf(spec = "", "", "|") & codes.Cells(i, 1).Value & "~" & Val(CStr(qty.Cells(i, 1).Value))
        End If
    Next
    Dim res As String
    res = CheckIn(Sel(), DeskDate(), spec)
    If res = "" Then Sh("Home").Unprotect: qty.ClearContents: Sh("Home").Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    ShowResult res
End Sub

Public Sub BtnInvoice(): ShowResult MakeInvoice(Sel(), DeskDate()): End Sub

Public Sub BtnPay()
    Dim a As Variant
    a = S("PayAmount")
    If Not IsNumeric(a) Or CStr(a) = "" Then SetMsg "Type the amount received in the yellow box first": Exit Sub
    ShowResult RecordPayment(Sel(), CDbl(a), DeskDate())
End Sub

Public Sub BtnCancel(): ShowResult CancelBooking(Sel(), DeskDate(), CStr(S("CancelReason"))): End Sub

Public Sub BtnLock()
    If Not IsDate(S("LockDate")) Then SetMsg "Type the last date to close in the yellow box first": Exit Sub
    ShowResult LockPeriod(CDate(S("LockDate")), DeskDate())
End Sub

Public Sub GoHome(): Sh("Home").Activate: Sh("Home").Range("A1").Select: End Sub
Public Sub GoAvailability(): Sh("Availability").Activate: Sh("Availability").Range("A1").Select: End Sub
Public Sub GoBookings(): Sh("Bookings").Activate: Sh("Bookings").Range("A1").Select: End Sub
Public Sub GoLog(): Sh("Log").Activate: Sh("Log").Cells(LastRow(Sh("Log")), 1).Select: End Sub
