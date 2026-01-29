# ============================================================================
# Simple Access Export - Direct VBA Execution
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = ""
)

if (-not (Test-Path $DatabasePath)) {
    Write-Error "Database not found: $DatabasePath"
    exit 1
}

# Determine output folder
if ([string]::IsNullOrEmpty($OutputFolder)) {
    $dbFolder = Split-Path $DatabasePath -Parent
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFolder = Join-Path $dbFolder "Exportacion_$timestamp"
}

Write-Host ""
Write-Host "Access Database Export" -ForegroundColor Cyan
Write-Host "  Database: $DatabasePath" -ForegroundColor White
Write-Host "  Output:   $OutputFolder" -ForegroundColor White
Write-Host ""

$access = $null
try {
    Write-Host "Opening Access..." -ForegroundColor Yellow
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($DatabasePath, $false)
    
    Write-Host "Checking VBA access..." -ForegroundColor Yellow
    
    $hasVBAAccess = $true
    $vbProject = $null
    try {
        $vbProject = $access.VBE.ActiveVBProject
        Write-Host "  VBA access: OK" -ForegroundColor Green
    }
    catch {
        $hasVBAAccess = $false
        Write-Host "  VBA access: DENIED" -ForegroundColor Red
        Write-Host ""
        Write-Host "REQUIRED ACTION:" -ForegroundColor Yellow
        Write-Host "Enable VBA programmatic access in Access:" -ForegroundColor White
        Write-Host "  1. Open Access" -ForegroundColor Gray
        Write-Host "  2. File -> Options -> Trust Center" -ForegroundColor Gray
        Write-Host "  3. Trust Center Settings" -ForegroundColor Gray
        Write-Host "  4. Enable: 'Trust access to the VBA project object model'" -ForegroundColor Gray
        throw "VBA programmatic access not enabled"
    }
    
    # Check if ExportTodoSimple exists
    $moduleName = "ExportTodoSimple"
    $moduleExists = $false
    
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Name -eq $moduleName) {
            $moduleExists = $true
            break
        }
    }
    
    if (-not $moduleExists) {
        Write-Host "Importing export module..." -ForegroundColor Yellow
        $modPath = "$PSScriptRoot\..\references\ExportTodoSimple.bas"
        $vbProject.VBComponents.Import($modPath)
        Write-Host "  Module imported" -ForegroundColor Green
        $needsCleanup = $true
    } else {
        Write-Host "Using existing export module" -ForegroundColor Yellow
        $needsCleanup = $false
    }
    
    Write-Host ""
    Write-Host "Executing export..." -ForegroundColor Yellow
    $procName = $moduleName + ".ExportAll"
    $access.Run($procName, $OutputFolder)
    
    Write-Host ""
    Write-Host "Export completed!" -ForegroundColor Green
    
    # Cleanup
    if ($needsCleanup) {
        Write-Host "Removing temporary module..." -ForegroundColor Yellow
        $vbProject.VBComponents.Remove($vbProject.VBComponents.Item($moduleName))
    }
    
    $access.CloseCurrentDatabase()
    $access.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    
    Write-Host ""
    Write-Host "EXPORT SUCCESSFUL" -ForegroundColor Green
    Write-Host "  Location: $OutputFolder" -ForegroundColor White
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
