Attribute VB_Name = "modButtons"
Option Explicit
' The buttons on Home. Each one works on the document picked in the yellow "Selected document" cell (SelDoc),
' or on the row you are on in the Register, and shows the core's answer in a message box.

Private Function Picked() As String
    If ActiveSheet.Name = "Register" And ActiveCell.Row >= REG_FIRST Then
        Picked = CStr(ActiveSheet.Cells(ActiveCell.Row, R_NO).Value)
    Else
        Picked = CStr(Sh("Home").Range("SelDoc").Value)
    End If
    If Picked = "" Then MsgBox "Pick a document in the yellow 'Selected document' cell on Home first.", vbInformation
End Function

Private Sub Tell(res As String, okText As String)
    If res = "" Then
        MsgBox okText, vbInformation
    ElseIf Left$(res, 8) = "BLOCKED:" Then
        MsgBox "Not issued. Fix these first:" & vbLf & vbLf & "- " & Replace(Mid$(res, 10), vbLf, vbLf & "- "), vbExclamation
    Else
        MsgBox Replace(res, "ERR: ", ""), vbExclamation
    End If
End Sub

Public Sub BtnNewQuote()
    frmDocument.StartNew "Quote"
    frmDocument.Show
End Sub

Public Sub BtnEditDraft()
    Dim no As String: no = Picked(): If no = "" Then Exit Sub
    If DocRow(no) = 0 Then MsgBox no & " not found.", vbExclamation: Exit Sub
    If Sh("Register").Cells(DocRow(no), R_STATUS).Value <> "Draft" Then
        MsgBox no & " is already issued. Issued documents are frozen; create a new one.", vbInformation: Exit Sub
    End If
    frmDocument.LoadDraft no
    frmDocument.Show
End Sub

Public Sub BtnIssue()
    Dim no As String: no = Picked(): If no = "" Then Exit Sub
    Tell IssueDocument(no, Date), no & " issued: PDF saved in the " & S("PdfFolder") & " folder" & _
        IIf(S("EmailMode") = "Off", ".", " and the email prepared (" & S("EmailMode") & ").")
End Sub

Public Sub BtnAccept()
    Dim no As String: no = Picked(): If no = "" Then Exit Sub
    Tell SetStatus(no, "Accepted", Date), no & " marked Accepted. Next: Convert to invoice."
End Sub

Public Sub BtnDecline()
    Dim no As String: no = Picked(): If no = "" Then Exit Sub
    Tell SetStatus(no, "Declined", Date), no & " marked Declined."
End Sub

Public Sub BtnConvert()
    Dim no As String, res As String
    no = Picked(): If no = "" Then Exit Sub
    res = ConvertToInvoice(no, Date)
    If Left$(res, 4) = "ERR:" Then Tell res, "": Exit Sub
    Sh("Home").Range("SelDoc").Value = res
    MsgBox "Invoice " & res & " created as a draft from " & no & ". Click Issue to send it.", vbInformation
End Sub

Public Sub BtnPaid()
    Dim no As String, v As Variant
    no = Picked(): If no = "" Then Exit Sub
    Dim d As Date
    v = Application.InputBox("Date the payment arrived, as yyyy-mm-dd:", "Mark paid", Format$(Date, "yyyy-mm-dd"), Type:=2)
    If VarType(v) = vbBoolean Then Exit Sub
    If Not v Like "####-##-##" Then MsgBox "Not a yyyy-mm-dd date: " & v, vbExclamation: Exit Sub
    d = DateSerial(Val(Left$(v, 4)), Val(Mid$(v, 6, 2)), Val(Right$(v, 2)))
    Tell SetStatus(no, "Paid", d), no & " marked Paid on " & DTxt(d) & "."
End Sub

Public Sub BtnShowDoc()
    Dim no As String: no = Picked(): If no = "" Then Exit Sub
    Sh("Document").Range("CurDoc").Value = no
    Sh("Document").Activate
End Sub

Public Sub BtnOpenPdf()
    Dim no As String, r As Long, f As String
    no = Picked(): If no = "" Then Exit Sub
    r = DocRow(no)
    If r = 0 Then Exit Sub
    f = CStr(Sh("Register").Cells(r, R_PDF).Value)
    If f = "" Then MsgBox no & " has no PDF yet: it is still a draft.", vbInformation: Exit Sub
    f = FolderOf("PdfFolder") & "\" & f
    If Dir(f) = "" Then MsgBox "PDF not found: " & f, vbExclamation: Exit Sub
    ThisWorkbook.FollowHyperlink f
End Sub

Public Sub BtnReset()
    If MsgBox("Clear ALL quotes, invoices, lines and the log, to start your own book?" & vbLf & _
              "Customers, products and settings are kept. PDF files are not deleted.", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    If MsgBox("Are you sure? This cannot be undone (keep a copy of the file first).", vbYesNo + vbCritical) <> vbYes Then Exit Sub
    ResetDocuments
    MsgBox "Done. Numbering starts again at " & NextNumber("Quote", Date) & ".", vbInformation
End Sub

Public Sub GoHome(): Sh("Home").Activate: End Sub
Public Sub GoDocument(): Sh("Document").Activate: End Sub
Public Sub GoRegister(): Sh("Register").Activate: End Sub
Public Sub GoLines(): Sh("Lines").Activate: End Sub
Public Sub GoCustomers(): Sh("Customers").Activate: End Sub
Public Sub GoProducts(): Sh("Products").Activate: End Sub
Public Sub GoLog(): Sh("Log").Activate: End Sub
Public Sub GoChecks(): Sh("Checks").Activate: End Sub
Public Sub GoSettings(): Sh("Settings").Activate: End Sub
Public Sub GoHelp(): Sh("How to use").Activate: End Sub
