# ============================================================================
# Access Database Export Script
# Exports all Access objects using COM automation and ExportTodoSimple module
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "",
    
    [Parameter(Mandatory=$false)]
    [string]$ExportModulePath = "$PSScriptRoot\..\references\ExportTodoSimple.bas"
)

# Validate inputs
if (-not (Test-Path $DatabasePath)) {
    Write-Error "Database file not found: $DatabasePath"
    exit 1
}

if (-not (Test-Path $ExportModulePath)) {
    Write-Error "Export module not found: $ExportModulePath"
    exit 1
}

# Determine output folder
if ([string]::IsNullOrEmpty($OutputFolder)) {
    $dbFolder = Split-Path $DatabasePath -Parent
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFolder = Join-Path $dbFolder "Exportacion_$timestamp"
}

Write-Host "Starting Access database export..." -ForegroundColor Cyan
Write-Host "  Database: $DatabasePath"
Write-Host "  Output:   $OutputFolder"
Write-Host ""

try {
    # Create Access Application COM object
    Write-Host "Opening Access application..." -ForegroundColor Yellow
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    
    # Open database
    Write-Host "Opening database..." -ForegroundColor Yellow
    $access.OpenCurrentDatabase($DatabasePath, $false)
    
    # Get VBA project
    $vbProject = $access.VBE.ActiveVBProject
    
    # Check if export module already exists
    $moduleName = "ExportTodoSimple"
    $moduleExists = $false
    
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq $moduleName) {
            $moduleExists = $true
            Write-Host "Export module already exists in database" -ForegroundColor Yellow
            break
        }
    }
    
    # Import export module if needed
    if (-not $moduleExists) {
        Write-Host "Importing export module..." -ForegroundColor Yellow
        Write-Host "  Module path: $ExportModulePath" -ForegroundColor Gray
        $importedComponent = $vbProject.VBComponents.Import($ExportModulePath)
        if ($importedComponent) {
            Write-Host "Export module imported" -ForegroundColor Green
        } else {
            throw "Failed to import export module"
        }
    }
    
    # Execute export function
    Write-Host "Executing export..." -ForegroundColor Yellow
    $access.Run("ExportAll", $OutputFolder)
    
    Write-Host "Export completed successfully!" -ForegroundColor Green
    
    # Clean up: remove temporary module if we added it
    if (-not $moduleExists) {
        Write-Host "Removing temporary export module..." -ForegroundColor Yellow
        $vbProject.VBComponents.Remove($vbProject.VBComponents.Item($moduleName))
        Write-Host "Cleanup completed" -ForegroundColor Green
    }
    
    # Close database
    $access.CloseCurrentDatabase()
    $access.Quit()
    
    # Release COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  Export Complete!" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  Output folder: $OutputFolder"
    Write-Host "  Next steps:"
    Write-Host "    1. Review 00_RESUMEN_APLICACION.txt"
    Write-Host "    2. Open in VS Code"
    Write-Host "    3. Explore VBA code in 06_Codigo_VBA/"
    Write-Host ""
    
    return $OutputFolder
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Error "Export failed: $errorMsg"
    
    # Cleanup on error
    if ($access) {
        try {
            $access.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        } catch {}
    }
    
    exit 1
}
