# Renders docs/*.png from Bookkeeping-Trial-Balance.xlsx (sheet -> PDF page -> PNG cropped to content).
# Run under the shared screen lock:  python tools/with_lock.py E7 -- powershell -ExecutionPolicy Bypass -File scripts/shots.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$py = 'python'  # needs PyMuPDF (fitz)
$tmp = Join-Path $env:TEMP 'e7-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Bookkeeping-Trial-Balance.xlsx'), 0, $true)
    # report sheets keep the print area set by build_workbook.py; working tabs get a fixed window
    $areas = [ordered]@{ 'Summary' = ''; 'P&L' = ''; 'Trial balance' = ''; 'Balance sheet' = ''; 'Cash flow' = ''; 'Bank rec' = '';
                         'Review' = ''; 'Month-end close' = ''; 'Payees' = '';
                         'Transfers' = 'A1:I8'; 'Rules' = 'A1:E24'; 'Transactions' = 'A1:N22'; 'Ledger' = 'A1:K24'; 'Cleaning log' = 'A1:D15' }
    foreach ($s in $areas.Keys) {
        $ws = $wb.Worksheets.Item($s)
        if ($areas[$s]) { $ws.PageSetup.PrintArea = $areas[$s] }
        $rng = $ws.Range($ws.PageSetup.PrintArea)
        $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        if ($rng.Height -gt 0.72 * $rng.Width) { $ws.PageSetup.Orientation = 1 } else { $ws.PageSetup.Orientation = 2 }
        $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
        $name = ($s -replace '&', 'and' -replace ' ', '-')
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp ($name + '.pdf')))
    }
    $wb.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
& $py -c @"
import fitz, pathlib
tmp = pathlib.Path(r'$tmp'); docs = pathlib.Path(r'$root') / 'docs'
for pdf in tmp.glob('*.pdf'):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect); bb = None
    for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
    for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
    bb = (bb or r) & r
    dpi = max(150, min(260, int(1900 / (bb.width / 72))))   # about 1900 px wide whatever the page scale
    page.get_pixmap(dpi=dpi, clip=bb).save(docs / (pdf.stem.lower() + '.png'))
    print('docs/' + pdf.stem.lower() + '.png', dpi)
"@
