# Renders docs/*.png from Sales-Commission-Calculator.xlsx (sheet -> PDF page -> PNG cropped to content).
# The workbook is opened read-only and never saved.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'e3-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
function Shot($ws, $area, $name, $orient) {
    $ws.PageSetup.Orientation = $orient; $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
    $ws.PageSetup.PrintArea = $area; $ws.PageSetup.PrintTitleRows = ''
    $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
    $ws.ExportAsFixedFormat(0, (Join-Path $tmp ($name + '.pdf')))
}
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Sales-Commission-Calculator.xlsx'), 0, $true)
    foreach ($n in 'Commission', 'Statement', 'Orders', 'Test cases') { $wb.Worksheets.Item($n).Unprotect('') }
    Shot $wb.Worksheets.Item('Commission') 'A1:L61' 'commission' 1
    $s = $wb.Worksheets.Item('Statement'); $s.Range('C4').Value2 = 'Lena Fischer'; $xl.CalculateFull()
    Shot $s 'A1:J40' 'statement' 1
    $wb.Worksheets.Item('Plan').Range('B9').Value2 = 'Stepped'; $xl.CalculateFull()
    Shot $s 'A1:J16' 'statement-stepped' 1
    $wb.Worksheets.Item('Plan').Range('B9').Value2 = 'Whole amount'; $xl.CalculateFull()
    Shot $wb.Worksheets.Item('Orders') 'A1:O30' 'orders' 2
    Shot $wb.Worksheets.Item('Test cases') 'A1:S29' 'test-cases' 2
    Shot $wb.Worksheets.Item('Plan') 'A1:R35' 'plan' 2
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
