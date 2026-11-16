# Asks the model that is open in Power BI Desktop for its numbers (DAX over the local Analysis Services port) and
# writes them to test/model-snapshot.json; test/verify.py compares that snapshot with its own Python calculation.
# Use: open Larkspur-Sales.pbip (or .pbix) in Power BI Desktop, Refresh, then
#      powershell -ExecutionPolicy Bypass -File test\query_model.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pkg = Get-AppxPackage -Name 'Microsoft.MicrosoftPowerBIDesktop' -ErrorAction SilentlyContinue
$dll = if ($pkg) { Join-Path $pkg.InstallLocation 'bin\Microsoft.PowerBI.AdomdClient.dll' }
else { 'C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.PowerBI.AdomdClient.dll' }
Add-Type -Path $dll

# the Analysis Services engine Power BI Desktop starts listens on a random local port: take it from the process
$engine = Get-Process msmdsrv -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
if (-not $engine) { throw 'Power BI Desktop is not running (no msmdsrv process): open the report first' }
$port = (Get-NetTCPConnection -State Listen -OwningProcess $engine.Id | Select-Object -First 1).LocalPort
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection "Data Source=localhost:$port"
$conn.Open()

function Run($dax) {
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $dax
    $r = $cmd.ExecuteReader(); $rows = @()
    while ($r.Read()) {
        $o = [ordered]@{}
        for ($i = 0; $i -lt $r.FieldCount; $i++) {
            $name = $r.GetName($i) -replace '^.*\[(.*)\]$', '$1'
            $v = $r.GetValue($i); if ($v -is [DBNull]) { $v = $null } elseif ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd') }
            $o[$name] = $v
        }
        $rows += [pscustomobject]$o
    }
    $r.Close(); return , $rows
}

$snap = [ordered]@{}
$snap.months = Run @"
EVALUATE SUMMARIZECOLUMNS ( 'Date'[Month No],
    "Net", [Net Sales], "Gross", [Gross Sales], "Discounts", [Discounts], "Returns", [Returns], "Orders", [Orders], "Units", [Units],
    "COGS", [COGS], "New", [New Customers], "Returning", [Returning Orders], "Shipping", [Shipping Charged], "Taxes", [Taxes Collected],
    "Ads", [Ad Spend], "CC", [Net Sales CC], "FX", [FX Effect], "AOV", [Average Order Value] )
"@
$snap.currencies = Run @"
EVALUATE SUMMARIZECOLUMNS ( 'Currency'[Currency], 'Date'[Month No], "Local", [Net Sales Local], "Net", [Net Sales] )
"@
$snap.products = Run @"
EVALUATE SUMMARIZECOLUMNS ( Sales[SKU], "Units", [Units], "Net", [Net Sales], "COGS", [COGS] )
"@
$snap.markets = Run @"
EVALUATE CALCULATETABLE ( SUMMARIZECOLUMNS ( Orders[Country], "Net", [Net Sales], "LY", [Net Sales LY] ), 'Date'[Year] = 2026 )
"@
$snap.cohorts = Run @"
EVALUATE SUMMARIZECOLUMNS ( Customer[Cohort No], 'Months Later'[Months Later], "R", [Cohort Retention %] )
"@
$snap.kpi = Run @"
EVALUATE CALCULATETABLE ( ROW ( "Net", [Net Sales], "LY", [Net Sales LY], "YoY", [Net Sales YoY %], "Orders", [Orders],
    "AOV", [Average Order Value], "AOVLY", [AOV LY], "GM", [Gross Margin %], "GMLY", [Gross Margin LY %], "Ret", [Returning Share %],
    "Refund", [Refund Rate %], "CC", [Net Sales CC], "CCG", [CC Growth %], "FX", [FX Effect], "NonUSD", [Non-USD Share %],
    "MER", [MER], "CAC", [CAC] ), 'Date'[Year] = 2026 )
"@
$snap.insights = Run @"
EVALUATE CALCULATETABLE ( ADDCOLUMNS ( VALUES ( Insights[No] ), "Text", [Insight] ), 'Date'[Year] = 2026 )
"@
$snap.checks = Run @"
EVALUATE ADDCOLUMNS ( VALUES ( Checks[No] ), "Status", [Check Status], "Detail", [Check Detail] )
"@
$snap.fx = Run @"
EVALUATE SUMMARIZECOLUMNS ( 'Currency'[Currency], "Latest", [Latest Rate] )
"@
$snap.fxinfo = Run @"
EVALUATE ROW ( "ApiRates", [API Rates], "Latest", [Rates Latest Date], "Source", [Rate Source] )
"@
$snap.orders = Run @"
EVALUATE SELECTCOLUMNS ( Orders, "Order", Orders[Order ID], "Rate", Orders[Rate], "Type", Orders[Customer Type], "Since", Orders[Months Since First] )
"@
$conn.Close()
$out = Join-Path $root 'test\model-snapshot.json'
[IO.File]::WriteAllText($out, ($snap | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
Write-Output "model-snapshot.json written from localhost:$port ($($snap.orders.Count) orders, $($snap.checks.Count) checks)"
