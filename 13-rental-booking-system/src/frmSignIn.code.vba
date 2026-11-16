Option Explicit
' Sign in: pick your name, type your PIN. The workbook keeps a hash of each PIN, never the PIN itself.

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, r As Long
    Set ws = Sh("Users")
    cboUser.Clear
    For r = FIRST To LastRow(ws)
        cboUser.AddItem ws.Cells(r, U_USER).Value
        cboUser.List(cboUser.ListCount - 1, 1) = ws.Cells(r, U_NAME).Value & " (" & ws.Cells(r, U_ROLE).Value & ")"
    Next
    lblMsg.Caption = ""
End Sub

Public Function DoSignIn() As String
    Dim res As String
    If cboUser.ListIndex < 0 Then lblMsg.Caption = "Pick your name first": DoSignIn = "ERR": Exit Function
    res = SignIn(cboUser.List(cboUser.ListIndex, 0), txtPin.Text)
    Application.Calculate
    If res = "" Then
        DoSignIn = ""
        Me.Hide
    Else
        lblMsg.Caption = Replace(res, "BLOCKED: ", "") & ". Try again."
        txtPin.Text = ""
        DoSignIn = res
    End If
End Function

Private Sub cmdSignIn_Click(): DoSignIn: End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
Private Sub txtPin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0: DoSignIn
End Sub
