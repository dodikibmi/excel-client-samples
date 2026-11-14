# Renders docs/*.png from Project-Cost-Tracker.xlsx (sheet -> PDF page -> PNG cropped to content).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'e8-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Project-Cost-Tracker.xlsx'), 0, $true)
    $areas = [ordered]@{ 'Dashboard' = @('A1:P60', 'dashboard', 1); 'Dashboard ' = @('A1:P31', 'dashboard-top', 2);
                         'Weekly trend' = @('A1:N37', 'weekly-trend', 2); 'P-2402' = @('A1:M27', 'pm-sheet', 2);
                         'Snapshots' = @('A1:L41', 'snapshots', 1) }
    foreach ($k in $areas.Keys) {
        $s = $k.Trim(); $a = $areas[$k]
        $ws = $wb.Worksheets.Item($s)
        $ws.PageSetup.Orientation = $a[2]; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        $ws.PageSetup.PrintArea = $a[0]
        $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp ($a[1] + '.pdf')))
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
    page.get_pixmap(dpi=170, clip=bb & r).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"@
