# Asks the model that is open in Power BI Desktop for its numbers (DAX over the local Analysis Services port) and
# writes them to test/model-snapshot.json; test/verify.py compares that snapshot with its own Python calculation.
# Use: open Copperleaf-Finance.pbip (or .pbix) in Power BI Desktop, Refresh, then
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
    "Income", [Total Income], "CostOfSales", [Cost of Sales], "OperatingExpenses", [Operating Expenses],
    "Expense", [Total Expense], "GrossProfit", [Gross Profit], "NetProfit", [Net Profit], "Budget", [Budget],
    "Variance", [Variance], "Cash", [Cash Balance], "InTransit", [In Transit], "IncomeMoM", [Income MoM %],
    "NetProfitPrevMonth", [Net Profit Previous Month], "NetProfitYTD", [Net Profit YTD] )
"@
$snap.total = Run @"
EVALUATE ROW ( "Income", [Total Income], "Expense", [Total Expense], "NetProfit", [Net Profit], "Budget", [Budget],
    "Variance", [Variance], "Cash", [Cash Balance], "IncomeMoM", [Income MoM %], "LedgerRows", COUNTROWS ( Ledger ),
    "Debits", SUM ( Ledger[Debit] ), "Credits", SUM ( Ledger[Credit] ), "ChecksPassed", [Checks Passed], "ChecksFailed", [Checks Failed] )
"@
$snap.accounts = Run @"
EVALUATE FILTER ( SUMMARIZECOLUMNS ( Account[Account No], 'Date'[Month No], "Actual", [Actual], "Budget", [Budget] ),
    NOT ISBLANK ( [Actual] ) || NOT ISBLANK ( [Budget] ) )
"@
$snap.statements = Run @"
EVALUATE SUMMARIZECOLUMNS ( Account[Account No], 'Date'[Month No],
    "Books", [Books Balance], "Statement", [Statement Balance], "Difference", [Difference vs Statement] )
"@
$snap.cashflow = Run @"
EVALUATE SUMMARIZECOLUMNS ( Ledger[Cash Flow Line], 'Date'[Month No], "Net", [Cash Movement] )
"@
$snap.checks = Run @"
EVALUATE ADDCOLUMNS ( VALUES ( Checks[No] ), "Status", [Check Status], "Detail", [Check Detail] )
"@
$snap.lines = Run @"
EVALUATE SELECTCOLUMNS ( FILTER ( Ledger, Ledger[Side] = "Register" ),
    "Line", Ledger[Line ID], "Date", Ledger[Date], "Register", Ledger[Register No], "Account", Ledger[Offset Account No],
    "Amount", Ledger[Amount] )
"@
$snap.kpi = Run @"
EVALUATE ROW ( "IncomeVsBudget", [Income vs Budget %], "ExpenseVsBudget", [Expense vs Budget %], "CashChange", [Cash Change] )
"@
$snap.insights = Run @"
EVALUATE ADDCOLUMNS ( VALUES ( Insights[No] ), "Text", [Insight] )
"@
$snap.forecast = Run @"
EVALUATE SUMMARIZECOLUMNS ( 'Forecast Week'[Week No], "Net", [Forecast Amount], "Balance", [Forecast Balance] )
"@
$snap.forecastKpi = Run @"
EVALUATE ROW ( "Today", [Cash Today], "Lowest", [Lowest Forecast Cash], "Headroom", [Headroom], "Below", [Weeks Below Buffer], "Closing", [Forecast Closing] )
"@
$snap.revenue = Run @"
EVALUATE ROW ( "GM", [Gross Margin %], "WS", [Wholesale Share %], "Payroll", [Payroll % of Income], "Spend", [Spend], "CounterDay", [Counter Sales per Day] )
"@
$snap.payees = Run @"
EVALUATE SUMMARIZECOLUMNS ( Ledger[Payee], "Spend", [Spend] )
"@
$snap.running = Run @"
EVALUATE CALCULATETABLE ( SUMMARIZECOLUMNS ( Ledger[Line ID], Ledger[Date], "Balance", [Running Balance] ), Account[Account Label] = "1000 Business checking" )
"@
$conn.Close()
$out = Join-Path $root 'test\model-snapshot.json'
[IO.File]::WriteAllText($out, ($snap | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
Write-Output "model-snapshot.json written from localhost:$port ($($snap.lines.Count) register lines, $($snap.checks.Count) checks)"
