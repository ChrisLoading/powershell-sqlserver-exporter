param(
    [string]$Server = "<YOUR_SQL_SERVER_NAME>",
    [string]$Database = "<YOUR_DATABASE_NAME>",
    [string]$Outdir, # Allow temporary directory override
    [string]$OutdirLog, # Allow temporary directory override
    [ValidateSet("csv", "xlsx")]
    [string]$Format = "csv",
    [switch]$PauseOnError
)

# PowerShell settings 
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Path bootstrap
$ScriptRoot = $PSScriptRoot
$TaskRoot = Split-Path $ScriptRoot -Parent
$LocalLogRoot = Join-Path $TaskRoot "Logs"

$ShareRoot = "D:\<SMB>\<CATEGORY>\<TASK>"
$ShareOutdir = Join-Path $ShareRoot "Exports"
$ShareLogRoot = Join-Path $ShareRoot "Logs" 

if ([string]::IsNullOrWhiteSpace($Outdir)) { $Outdir = $ShareOutdir }
if ([string]::IsNullOrWhiteSpace($OutdirLog)) { $OutdirLog = $ShareLogRoot }

# Ensure output directories exist
New-Item -ItemType Directory -Force -Path $Outdir, $OutdirLog, $LocalLogRoot | Out-Null

# Build file path
$now = Get-Date
$stamp = $now.ToString("yyyyMMdd_HHmmss")

# Start logging locally
$localLogFile = Join-Path $LocalLogRoot ("Visit_{0}.log" -f $stamp)
New-Item -ItemType Directory -Force -Path $LocalLogRoot | Out-Null

Start-Transcript -Path $localLogFile -Append -ErrorAction Stop | Out-Null

# Query range: 48 hour rolling window
# $end = $now
# $start = $end.AddDays(-2)

# Query range: 2 days ago to now (down to second)
# $end = $now
# $start = $end.Date.AddDays(-2)

# Query range: 2 days ago to today (full days)
$today = (Get-Date).Date
$end = $today.AddDays(1) # 00:00:00 of the next day when the script runs
$start = $today.AddDays(-2) # 00:00:00 of 2 days ago

# SQL Query
$query = @"
SET NOCOUNT ON;

DECLARE @Start datetime2(0) = @pStart;
DECLARE @End datetime2(0) = @pEnd;

SELECT *
FROM dbo.vVisit_WithVisitDT
WHERE VisitDT IS NOT NULL
  AND VisitDT >= @Start
  AND VisitDT < @End
ORDER BY VisitDT DESC;
"@

$conn = $null
$cmd = $null
$da = $null
$dt = $null
$ShareLogFile = Join-Path $OutdirLog ("Visit_{0}.log" -f $stamp)

try {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Start export... Server=$Server DB=$Database Format=$Format"

    # Database connection via .NET (instead of Invoke-Sqlcmd)
    $connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)

    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query

    $p1 = $cmd.Parameters.Add("@pStart", [System.Data.SqlDbType]::DateTime2);
    $p1.Value = $start
    $p2 = $cmd.Parameters.Add("@pEnd", [System.Data.SqlDbType]::DateTime2);
    $p2.Value = $end

    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable

    # Throw exception if SQL/conn/cmd fails 
    [void]$da.Fill($dt)
    $conn.Close()

    Write-Host "Rows retrieved: $($dt.Rows.Count)"

    # Export data by selected format
    if ($Format -eq "csv") {
        $outFile = Join-Path $OutDir ("Visit_{0}.csv" -f $stamp)
        $dt | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    }
    else {
        # xlsx: Module ImportExcel required (No Microsoft Office dependency)

        # Installation: 
        # Install-Module ImportExcel -Scope AllUsers
        Import-Module ImportExcel -ErrorAction Stop
        $outFile = Join-Path $OutDir ("Visit_{0}.xlsx" -f $stamp)
        $dt | Export-Excel -Path $outFile -WorksheetName "data" -AutoSize -BoldTopRow -FreezeTopRow -ErrorAction Stop
    }
    Write-Host "Exported to: $outFile"
    exit 0
}
catch {
    Write-Error "FAILED: $($_.Exception.Message)"
    if ($PauseOnError -and $Host.Name -eq 'ConsoleHost') {
        Read-Host "Error. Press Enter to close...(Check log: $localLogFile)"
    }
    exit 1
}
finally {
    # Cleanup
    if ($conn) { $conn.Dispose() }
    if ($da) { $da.Dispose() }
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch {}

    $SyncLogRoot = Join-Path $LocalLogRoot "SyncLogs"
    New-Item -ItemType Directory -Force -Path $SyncLogRoot | Out-Null
    $syncLogFile = Join-Path $SyncLogRoot "_sync.log"

    # Copy local log to share and record in sync log
    try {
        New-Item -ItemType Directory -Force -Path $OutdirLog | Out-Null
        Copy-Item -Path $localLogFile -Destination $ShareLogFile -Force

        Add-Content -Path $syncLogFile -Encoding UTF8 -Value (
            "[{0}] OK.  Log copied to -> {1}" -f $stamp, $shareLogFile
        )
    }
    catch {
        # best effort logging: in case of failure, do not throw sync error to override main result
        $syncErrMsg = 
            "[{0}] Error. Failed to copy log. Local={1}; Share={2}; Error={3}" -f $stamp, $localLogFile, $shareLogFile, $_.Exception.Message
        try { Add-Content -Path $syncLogFile -Encoding UTF8 -Value $syncErrMsg }
        catch {}
    }
}
