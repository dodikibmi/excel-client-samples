# Renders docs/*.png from Project-Cost-Tracker.xlsx (sheet -> PDF page -> PNG cropped to content).
# Each sheet uses the print area set by build_workbook.py; 'dashboard-top' is the upper part of the Dashboard.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'e8-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
Get-ChildItem $tmp -Filter *.pdf | Remove-Item
New-Item -ItemType Directory -Force (Join-Path $root 'docs') | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'Project-Cost-Tracker.xlsx'), 0, $true)
    # sheet, png name, print area ('' = the one in the workbook), orientation (1 portrait, 2 landscape)
    $shots = @(@('Dashboard', 'dashboard', '', 1), @('Dashboard', 'dashboard-top', 'A1:S29', 2),
               @('Project view', 'project-view', '', 2), @('S-curve', 's-curve', '', 2),
               @('COs & risks', 'cos-risks', '', 2), @('WIP & cash', 'wip-cash', '', 2),
               @('Weekly trend', 'weekly-trend', '', 2), @('P-2402', 'pm-sheet', '', 2))
    foreach ($a in $shots) {
        $ws = $wb.Worksheets.Item($a[0])
        $ws.PageSetup.Orientation = $a[3]; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        if ($a[2] -ne '') { $ws.PageSetup.PrintArea = $a[2] }
        $ws.PageSetup.LeftMargin = 18; $ws.PageSetup.RightMargin = 18; $ws.PageSetup.TopMargin = 18; $ws.PageSetup.BottomMargin = 18
        $ws.PageSetup.CenterHorizontally = $false
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
    for img in page.get_image_info(): bb = fitz.Rect(img['bbox']) if bb is None else bb | fitz.Rect(img['bbox'])
    bb = (bb or r) & r
    zoom = 1900 / bb.width if bb.width * 170 / 72 > 1900 else 170 / 72
    page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=bb).save(docs / (pdf.stem + '.png'))
    print('docs/' + pdf.stem + '.png')
"@
