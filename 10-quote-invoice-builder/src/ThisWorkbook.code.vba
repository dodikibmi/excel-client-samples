Option Explicit
' Protection with UserInterfaceOnly is not saved in the file, so it is switched on again every time the book opens:
' people cannot type over the register, the macros still can.
Private Sub Workbook_Open()
    On Error Resume Next
    SheetsOpen
    RefreshDocList
    SheetsLock
    Sh("Home").Activate
End Sub
