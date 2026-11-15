# Renders docs/*.png from Sales-Commission-Calculator.xlsx (sheet area -> PDF page -> PNG cropped to content).
# The workbook is opened read-only and never saved. Run it under the shared Excel lock:
#   python tools/with_lock.py E3 -- powershell -ExecutionPolicy Bypass -File scripts/shots.ps1
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
    $ws.PageSetup.LeftMargin = 10; $ws.PageSetup.RightMargin = 10; $ws.PageSetup.TopMargin = 10; $ws.PageSetup.BottomMargin = 10
    $ws.PageSetup.CenterHorizontally = $false
    $ws.ExportAsFixedFormat(0, (Join-Path $tmp ($name + '.pdf')))
}
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Sales-Commission-Calculator.xlsx'), 0, $true)
    foreach ($ws in $wb.Worksheets) { $ws.Unprotect('') }
    $method = $wb.Names.Item('Method').RefersToRange
    $rep = $wb.Names.Item('StmtRep').RefersToRange
    Shot $wb.Worksheets.Item('Dashboard') 'A1:AK49' 'dashboard' 2
    Shot $wb.Worksheets.Item('Commission') 'A1:P80' 'commission' 1
    Shot $wb.Worksheets.Item('What-if') 'A1:M53' 'what-if' 2
    Shot $wb.Worksheets.Item('Payouts') 'A1:P36' 'payouts' 2
    $s = $wb.Worksheets.Item('Statement')
    $rep.Value2 = 'Lena Fischer'; $xl.CalculateFull()
    Shot $s 'A7:J72' 'statement' 1
    $method.Value2 = 'Stepped'; $xl.CalculateFull()
    $rep.Value2 = 'Grace Holloway'; $xl.CalculateFull()
    Shot $s 'A7:J44' 'statement-stepped' 1
    $method.Value2 = 'Whole amount'; $rep.Value2 = 'Lena Fischer'; $xl.CalculateFull()
    Shot $wb.Worksheets.Item('Orders') 'A1:S32' 'orders' 2
    Shot $wb.Worksheets.Item('Returns') 'A1:W24' 'returns' 2
    Shot $wb.Worksheets.Item('Test cases') 'A1:Q60' 'test-cases' 1
    Shot $wb.Worksheets.Item('Plan') 'A1:Q34' 'plan' 2
    $wb.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
python -c @"
import fitz, pathlib
fitz.mupdf.fz_set_graphics_aa_level(0)  # no hairline seams between filled cells
tmp = pathlib.Path(r'$tmp'); docs = pathlib.Path(r'$root') / 'docs'
for pdf in tmp.glob('*.pdf'):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect); bb = None
    for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
    for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
    bb = (bb or r) + (-2, -2, 2, 2)
    page.get_pixmap(dpi=170, clip=bb & r).save(docs / (pdf.stem.lower() + '.png'))
    print('docs/' + pdf.stem.lower() + '.png')
"@
