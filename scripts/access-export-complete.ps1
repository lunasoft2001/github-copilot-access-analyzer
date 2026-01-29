# ============================================================================
# Complete Access Export - Full export including VBA, Forms, Reports
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDatabase,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = ""
)

$exportModulePath = "$PSScriptRoot\..\references\ExportTodoSimple.bas"

# Validate files exist
if (-not (Test-Path $TargetDatabase)) {
    Write-Error "Target database not found: $TargetDatabase"
    exit 1
}

if (-not (Test-Path $exportModulePath)) {
    Write-Error "Export module not found: $exportModulePath"
    exit 1
}

# Determine output folder
if ([string]::IsNullOrEmpty($OutputFolder)) {
    $targetFolder = Split-Path $TargetDatabase -Parent
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFolder = Join-Path $targetFolder "Exportacion_$timestamp"
}

Write-Host ""
Write-Host "Complete Access Database Export" -ForegroundColor Cyan
Write-Host "  Target:   $TargetDatabase" -ForegroundColor White
Write-Host "  Output:   $OutputFolder" -ForegroundColor White
Write-Host ""

$access = $null
try {
    Write-Host "Opening target database..." -ForegroundColor Yellow
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    
    # Open target database as CurrentDatabase (exclusive mode for import)
    $access.OpenCurrentDatabase($TargetDatabase, $true)
    
    Write-Host "Checking VBA access..." -ForegroundColor Yellow
    $vbProject = $null
    $hasVBAAccess = $true
    try {
        $vbProject = $access.VBE.ActiveVBProject
        Write-Host "  VBA access: OK" -ForegroundColor Green
    }
    catch {
        Write-Host "  VBA access: Limited" -ForegroundColor Yellow
        $hasVBAAccess = $false
    }
    
    # Check if ExportTodoSimple module exists
    $moduleName = "ExportTodoSimple"
    $moduleExists = $false
    $needsCleanup = $false
    
    if ($hasVBAAccess) {
        foreach ($comp in $vbProject.VBComponents) {
            if ($comp.Name -eq $moduleName) {
                $moduleExists = $true
                break
            }
        }
        
        if (-not $moduleExists) {
            Write-Host "Importing export module..." -ForegroundColor Yellow
            $vbProject.VBComponents.Import($exportModulePath)
            Write-Host "  Module imported" -ForegroundColor Green
            $needsCleanup = $true
        } else {
            Write-Host "Using existing export module..." -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "Executing full export..." -ForegroundColor Yellow
    
    # Call ExportAll
    $access.Run("ExportAll", $OutputFolder)
    
    Write-Host ""
    Write-Host "Export completed!" -ForegroundColor Green
    
    # Cleanup
    if ($needsCleanup -and $hasVBAAccess) {
        Write-Host "Removing temporary module..." -ForegroundColor Yellow
        $vbProject.VBComponents.Remove($vbProject.VBComponents.Item($moduleName))
    }
    
    $access.CloseCurrentDatabase()
    $access.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    
    Write-Host ""
    Write-Host "COMPLETE EXPORT SUCCESSFUL" -ForegroundColor Green
    Write-Host "  Location: $OutputFolder" -ForegroundColor White
    Write-Host ""
    Write-Host "Exported:" -ForegroundColor Cyan
    Write-Host "  - Tables structure" -ForegroundColor Gray
    Write-Host "  - Queries" -ForegroundColor Gray
    Write-Host "  - Forms" -ForegroundColor Gray
    Write-Host "  - Reports" -ForegroundColor Gray
    Write-Host "  - Macros" -ForegroundColor Gray
    Write-Host "  - VBA Modules" -ForegroundColor Gray
    Write-Host ""
    
    return $OutputFolder
}
catch {
    Write-Host ""
    $errMsg = $_.Exception.Message
    Write-Error "Export failed: $errMsg"
    
    if ($access -ne $null) {
        try {
            $access.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        } catch { }
    }
    
    exit 1
}
