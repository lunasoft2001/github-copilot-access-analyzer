# Importar archivos modificados de vuelta a Access

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDbPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ImportFolder,
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\AccessAnalyzer.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "IMPORTACION A BASE DE DATOS ACCESS" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar archivos
if (-not (Test-Path $AnalyzerPath)) {
    Write-Error "No se encuentra AccessAnalyzer.accdb: $AnalyzerPath"
    exit 1
}

if (-not (Test-Path $TargetDbPath)) {
    Write-Error "No se encuentra base de datos destino: $TargetDbPath"
    exit 1
}

if (-not (Test-Path $ImportFolder)) {
    Write-Error "No se encuentra carpeta de importación: $ImportFolder"
    exit 1
}

$access = $null

try {
    Write-Host "1. Abriendo AccessAnalyzer..." -ForegroundColor Yellow
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AnalyzerPath, $false)
    
    Write-Host "   OK" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "2. Creando copia de seguridad..." -ForegroundColor Yellow
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = $TargetDbPath -replace '\.accdb$', "_BACKUP_BEFORE_IMPORT_$timestamp.accdb"
    Copy-Item $TargetDbPath $backupPath -Force
    
    Write-Host "   Backup: $backupPath" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "3. Ejecutando importación..." -ForegroundColor Yellow
    Write-Host "   Base destino: $TargetDbPath" -ForegroundColor Cyan
    Write-Host "   Carpeta fuente: $ImportFolder" -ForegroundColor Cyan
    Write-Host "   Idioma: $Language" -ForegroundColor Cyan
    Write-Host ""
    
    # Construir comando
    $targetEscaped = $TargetDbPath.Replace('\', '\\')
    $folderEscaped = $ImportFolder.Replace('\', '\\')
    $cmd = 'RunCompleteImport("' + $targetEscaped + '","' + $folderEscaped + '","' + $Language + '")'
    
    $result = $access.Eval($cmd)
    
    if ($result) {
        Write-Host ""
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host "IMPORTACION EXITOSA!" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Base de datos actualizada: $TargetDbPath" -ForegroundColor White
        Write-Host "Copia de seguridad: $backupPath" -ForegroundColor Gray
    }
    else {
        Write-Host ""
        Write-Host "ERROR en la importación" -ForegroundColor Red
        Write-Host "Backup disponible en: $backupPath" -ForegroundColor Yellow
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    if ($access) {
        Write-Host ""
        Write-Host "Cerrando Access..." -ForegroundColor Yellow
        $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveAll)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Finalizado" -ForegroundColor Green
}
