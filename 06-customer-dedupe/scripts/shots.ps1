# Renders docs/*.png from Customer-Cleanup.xlsx (sheet -> PDF page -> PNG cropped to content).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'e6-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Customer-Cleanup.xlsx'), 0, $true)
    $areas = @{ 'Summary' = 'A1:M29'; 'Customers' = 'A1:P20'; 'Needs review' = 'A1:N12'; 'Merge log' = 'A1:L20'; 'CRM import' = 'A1:L18'; 'Cleaning log' = 'A1:D32' }
    foreach ($s in $areas.Keys) {
        $ws = $wb.Worksheets.Item($s)
        $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        $ws.PageSetup.PrintArea = $areas[$s]
        $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp (($s -replace ' ', '-') + '.pdf')))
    }
    $wb.Close($false)
    # the 'before' picture of the sales team's spreadsheet: title rows, merged company cells
    $wb = $xl.Workbooks.Open((Join-Path $root 'data\sales-team-accounts.xlsx'), 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
    $ws.PageSetup.PrintArea = 'A1:K16'; $ws.PageSetup.PrintGridlines = $true
    $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
    $ws.ExportAsFixedFormat(0, (Join-Path $tmp 'before-sales-sheet.pdf'))
    $wb.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
python -c @"
import fitz, pathlib
tmp = pathlib.Path(r'$tmp'); docs = pathlib.Path(r'$root') / 'docs'
for pdf in tmp.glob('*.pdf'):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect); bb = None
    for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
    for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
    bb = (bb or r) + (-8, -8, 8, 8)
    page.get_pixmap(dpi=170, clip=bb & r).save(docs / (pdf.stem.lower() + '.png'))
    print('docs/' + pdf.stem.lower() + '.png')
"@
