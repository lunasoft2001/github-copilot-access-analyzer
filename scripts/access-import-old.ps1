# ============================================================================
# Access Database Import Script
# Re-imports modified VBA code and queries back into Access database
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$true)]
    [string]$SourceFolder,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateBackup = $true
)

# Validate inputs
if (-not (Test-Path $DatabasePath)) {
    Write-Error "Database file not found: $DatabasePath"
    exit 1
}

if (-not (Test-Path $SourceFolder)) {
    Write-Error "Source folder not found: $SourceFolder"
    exit 1
}

Write-Host "Starting Access database import..." -ForegroundColor Cyan
Write-Host "  Database: $DatabasePath"
Write-Host "  Source:   $SourceFolder"
Write-Host ""

# Create backup first
if ($CreateBackup) {
    Write-Host "Creating backup..." -ForegroundColor Yellow
    $backupScript = Join-Path $PSScriptRoot "access-backup.ps1"
    $backupPath = & $backupScript -DatabasePath $DatabasePath
    Write-Host ""
}

try {
    # Create Access Application COM object
    Write-Host "Opening Access application..." -ForegroundColor Yellow
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    
    # Open database
    Write-Host "Opening database..." -ForegroundColor Yellow
    $access.OpenCurrentDatabase($DatabasePath, $true) # $true = exclusive mode for imports
    
    $importLog = @()
    $successCount = 0
    $failCount = 0
    
    # Import VBA Modules
    Write-Host "Importing VBA modules..." -ForegroundColor Yellow
    $vbaFolder = Join-Path $SourceFolder "06_Codigo_VBA"
    if (Test-Path $vbaFolder) {
        $vbProject = $access.VBE.ActiveVBProject
        $basFiles = Get-ChildItem -Path $vbaFolder -Filter "*.bas"
        
        foreach ($basFile in $basFiles) {
            try {
                $moduleName = $basFile.BaseName
                
                # Remove existing module if exists
                foreach ($component in $vbProject.VBComponents) {
                    if ($component.Name -eq $moduleName) {
                        $vbProject.VBComponents.Remove($component)
                        break
                    }
                }
                
                # Import new version
                $vbProject.VBComponents.Import($basFile.FullName) | Out-Null
                
                $importLog += "✓ VBA Module: $moduleName"
                $successCount++
                Write-Host "  ✓ $moduleName" -ForegroundColor Green
            }
            catch {
                $importLog += "✗ VBA Module: $moduleName - Error: $_"
                $failCount++
                Write-Host "  ✗ $moduleName - $_" -ForegroundColor Red
            }
        }
    }
    
    # Import Queries
    Write-Host "Importing queries..." -ForegroundColor Yellow
    $queryFolder = Join-Path $SourceFolder "02_Consultas"
    if (Test-Path $queryFolder) {
        $db = $access.CurrentDb()
        $sqlFiles = Get-ChildItem -Path $queryFolder -Filter "*.sql"
        
        foreach ($sqlFile in $sqlFiles) {
            try {
                # Skip the index file
                if ($sqlFile.Name -eq "00_Lista_Consultas.txt") { continue }
                
                $queryName = $sqlFile.BaseName
                $sqlContent = Get-Content $sqlFile.FullName -Raw -Encoding UTF8
                
                # Remove comment lines
                $sqlContent = ($sqlContent -split "`n" | Where-Object { $_ -notmatch "^--" }) -join "`n"
                $sqlContent = $sqlContent.Trim()
                
                if ([string]::IsNullOrWhiteSpace($sqlContent)) { continue }
                
                # Delete existing query if exists
                foreach ($qry in $db.QueryDefs) {
                    if ($qry.Name -eq $queryName) {
                        $db.QueryDefs.Delete($queryName)
                        break
                    }
                }
                
                # Create new query
                $newQuery = $db.CreateQueryDef($queryName, $sqlContent)
                
                $importLog += "✓ Query: $queryName"
                $successCount++
                Write-Host "  ✓ $queryName" -ForegroundColor Green
            }
            catch {
                $importLog += "✗ Query: $queryName - Error: $_"
                $failCount++
                Write-Host "  ✗ $queryName - $_" -ForegroundColor Red
            }
        }
    }
    
    # Note: Form and Report code is embedded in their definitions
    # Full re-import of forms/reports would require LoadFromText
    # which is complex and may lose formatting. Typically only VBA
    # modules and queries are re-imported after refactoring.
    
    Write-Host ""
    Write-Host "Closing database..." -ForegroundColor Yellow
    $access.CloseCurrentDatabase()
    $access.Quit()
    
    # Release COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Save import log
    $logPath = Join-Path $SourceFolder "IMPORT_LOG_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $importLog | Out-File -FilePath $logPath -Encoding UTF8
    
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Import Complete!" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Successful: $successCount"
    Write-Host "  Failed:     $failCount"
    Write-Host "  Log file:   $logPath"
    Write-Host ""
    
    if ($failCount -gt 0) {
        Write-Host "⚠ Some imports failed. Check log file for details." -ForegroundColor Yellow
    }
    
    return @{
        Success = $successCount
        Failed = $failCount
        LogPath = $logPath
    }
}
catch {
    Write-Error "Import failed: $_"
    Write-Error $_.Exception.Message
    
    # Cleanup on error
    if ($access) {
        try {
            $access.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        } catch {}
    }
    
    exit 1
}
