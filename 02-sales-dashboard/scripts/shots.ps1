# Renders docs/*.png from Sales-Dashboard.xlsx (sheet -> PDF page -> PNG cropped to content).
# dashboard-all = Oct 2026, all regions; dashboard-west = Oct 2026, West only (workbook is opened read-only, never saved).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'e2-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Sales-Dashboard.xlsx'), 0, $true)
    $d = $wb.Worksheets.Item('Dashboard')
    foreach ($region in 'All', 'West') {
        $d.Range('G4').Value2 = $region
        $xl.CalculateFull()
        $d.ExportAsFixedFormat(0, (Join-Path $tmp ('dashboard-' + $region.ToLower() + '.pdf')))
    }
    $areas = @{ 'Calc' = 'A1:N31'; 'Sales data' = 'A1:N25'; 'How to use' = 'A1:A15' }
    foreach ($s in $areas.Keys) {
        $ws = $wb.Worksheets.Item($s)
        $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        $ws.PageSetup.PrintArea = $areas[$s]
        $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp (($s -replace ' ', '-') + '.pdf')))
    }
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
