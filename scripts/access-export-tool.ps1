# ============================================================================
# Access Export Tool - Uses AccessAnalyzer.accdb to export other databases
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDatabase,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = ""
)

$analyzerPath = "$PSScriptRoot\..\AccessAnalyzer.accdb"

# Validate analyzer exists
if (-not (Test-Path $analyzerPath)) {
    Write-Error "AccessAnalyzer.accdb not found at: $analyzerPath"
    Write-Host "`nPlease create it following instructions in SETUP.md" -ForegroundColor Yellow
    exit 1
}

# Validate target exists
if (-not (Test-Path $TargetDatabase)) {
    Write-Error "Target database not found: $TargetDatabase"
    exit 1
}

# Determine output folder
if ([string]::IsNullOrEmpty($OutputFolder)) {
    $targetFolder = Split-Path $TargetDatabase -Parent
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFolder = Join-Path $targetFolder "Exportacion_$timestamp"
}

Write-Host ""
Write-Host "Access Database Export Tool" -ForegroundColor Cyan
Write-Host "  Analyzer: $analyzerPath" -ForegroundColor Gray
Write-Host "  Target:   $TargetDatabase" -ForegroundColor White
Write-Host "  Output:   $OutputFolder" -ForegroundColor White
Write-Host ""

$access = $null
try {
    Write-Host "Opening AccessAnalyzer..." -ForegroundColor Yellow
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($analyzerPath, $false)
    
    Write-Host "Executing export..." -ForegroundColor Yellow
    
    # Use Eval with wrapper function that returns a value
    $evalString = "RunExport(`"$TargetDatabase`",`"$OutputFolder`")"
    $result = $access.Eval($evalString)
    
    Write-Host ""
    Write-Host "Export completed!" -ForegroundColor Green
    
    $access.CloseCurrentDatabase()
    $access.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    
    Write-Host ""
    Write-Host "EXPORT SUCCESSFUL" -ForegroundColor Green
    Write-Host "  Location: $OutputFolder" -ForegroundColor White
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Review 00_RESUMEN.txt" -ForegroundColor Gray
    Write-Host "  2. Open in VS Code: code '$OutputFolder'" -ForegroundColor Gray
    Write-Host "  3. Note: VBA code requires separate export (see 06_Codigo_VBA/00_NOTA.txt)" -ForegroundColor Gray
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
