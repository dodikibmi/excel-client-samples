Attribute VB_Name = "ExportStatements"
' Optional macro: exports one PDF statement per sales rep for the month picked on the Statement sheet.
' The workbook ships macro-free. To use this: Save As "Excel Macro-Enabled Workbook (*.xlsm)",
' Alt+F11 > File > Import File... > this .bas file, then run ExportAllStatements (Alt+F8).
' PDFs are written next to the workbook, e.g. "2026-10 Lena Fischer.pdf".
Option Explicit

Public Sub ExportAllStatements()
    Const FIRST_LINE As Long = 47            ' first order line row on the Statement sheet
    Dim ws As Worksheet, repCell As Range, rep As Range, n As Long, folder As String, keepRep As Variant
    Set ws = ThisWorkbook.Worksheets("Statement")
    Set repCell = ThisWorkbook.Names("StmtRep").RefersToRange
    keepRep = repCell.Value
    folder = ThisWorkbook.Path & Application.PathSeparator
    ws.Unprotect ""
    For Each rep In ThisWorkbook.Names("RepList").RefersToRange.Cells
        If Len(rep.Value) > 0 Then
            repCell.Value = rep.Value
            Application.Calculate
            n = Application.WorksheetFunction.Count(ws.Range("I" & FIRST_LINE & ":I" & FIRST_LINE + 99))
            ws.PageSetup.PrintArea = "$A$7:$J$" & (FIRST_LINE + n + 1)
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & _
                Format(ThisWorkbook.Names("StmtMonth").RefersToRange.Value, "yyyy-mm") & " " & rep.Value & ".pdf"
        End If
    Next rep
    repCell.Value = keepRep
    ws.Protect "", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    MsgBox "Statements saved in " & folder, vbInformation
End Sub
