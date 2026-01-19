# SQL Server Data Export Automation

> PowerShell and T-SQL scripts for automated visit data extraction from SQL Server with support for CSV and XLSX exports.

## Overview

This project provides automated data export functionality for the 行政網路系統 (Administrative Network System) database. It includes:

- **PowerShell script** for querying and exporting visit data with logging
- **T-SQL view** that converts ROC calendar dates to standard datetime2 format
- Configurable export formats (CSV/XLSX)
- Local and network-based (SMB) logging options

## Requirements

- **PowerShell** 5.0 or later
- **SQL Server** 2016 or later (with Integrated Authentication)
- **.NET Framework** 4.7.2+ (for SqlClient)
- **ImportExcel module** (Optional: for XLSX export only)
  
  ```powershell
  Install-Module ImportExcel -Scope AllUsers
  ```

## Quick start

### Basic CSV Export (Default)

```powershell
.\Export-Visit.ps1
```

### Export to XLSX

```powershell
.\Export-Visit.ps1 -Format xlsx
```

### Custom Server and Output Directory

```powershell
.\Export-Visit.ps1 -Server "<YOUR_SQL_SERVER_NAME>" -Database "<YOUR_DATABASE_NAME>" -Outdir "C:\CUSTOM\PATH" -OutdirLog "D:\<SMB>\CUSTOM_LOG_PATH" -Format csv
```

### With Error Pause (For Manual Execution)

```powershell
.\Export-Visit.ps1 -PauseOnError
```

## Configuration

### Script Parameters

| Parameter | Type | Default | Description |
| --------- | ---- | ------- | ----------- |
| `-Server` | string | `.\SQLEXPRESS` | SQL Server instance |
| `-Database` | string | `Northwind` | Target database name |
| `-Outdir` | string | `C:\Custom\Path` | Export output directory |
| `-OutdirLog` | string | `D:\<SMB>\LogPath` | Network log directory |
| `-Format` | string | `csv` | Export format: `csv` or `xlsx` |
| `-PauseOnError` | switch | false | Pause on error (interactive/test mode) |

## Logging

Logs are created in two locations:

- **Local**: `.\Logs\Visit_{timestamp}.log`
- **Network (SMB)**: `.\<SMB>\LogPath\Visit_{timestamp}.log`
- **Sync Log**: `.\Logs\SyncLogs\_sync.log`

## Database Setup

Run [create_view.sql](create_view.sql) to create the required view:

```sql
    CREATE OR ALTER VIEW dbo.vVisit_WithVisitDT AS
        ...
    FROM dbo.<YourTable>;
```

This view converts ROC calendar dates (民國) to standard datetime format.

> **Note**: Consider **computed columns (persisted/unpersisted)** for performance optimization / indexing if needed. Refer to [Microsoft Docs on Computed Columns](https://learn.microsoft.com/en-us/sql/relational-databases/tables/specify-computed-columns-in-a-table?view=sql-server-ver16).

## Troubleshooting

- Ensure SQL Server allows Integrated Authentication.
- Verify network paths and permissions for SMB logging.
- Check PowerShell execution policy if scripts fail to run.

| Issue | Solution |
| ----- | -------- |
| **"Module ImportExcel not found"** | Run: `Install-Module ImportExcel -Scope AllUsers` |
| **SQL connection fails** | Verify SQL Server is running and accessible; check `-Server` parameter |
| **"Access Denied" on network share** | Verify user has write permissions to `D:\<SMB>` |
| **No rows exported** | Check visit date format in source table; verify date range in script |
| **Special characters corrupted** | **UTF-8 BOM** may be required if special characters are not displayed properly. Otherwise, file is UTF-8 encoded; ensure your viewer supports UTF-8 |

## Windows Task Scheduler Setup

1. Open Windows Task Scheduler
2. Create Task (recommended over Basic Task for more options)
3. Trigger: Daily at desired time
4. Action: Start program
   - Program: `powershell.exe`
   - (Optional) Program: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
     > Fully specify full path to avoid PowerShell version issues. (32-bit vs 64-bit)
   - Arguments: `-NoProfile -ExecutionPolicy Bypass -File "C:\path\to\Export-Visit.ps1"`
     > With optional parameters as mentioned above.
