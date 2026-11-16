Option Explicit
' Every time the book opens: nobody is signed in, and the protection (UserInterfaceOnly is not saved in the file)
' is switched on again, so people cannot type over bookings, lines, payments or the log; the macros still can.
Private Sub Workbook_Open()
    On Error Resume Next
    SheetsOpen
    SetS "SignedUser", ""
    SetS "SignedRole", ""
    SetS "HomeMsg", "Sign in to start (Sign in button, top right)"
    SheetsLock
    Sh("Home").Activate
End Sub
