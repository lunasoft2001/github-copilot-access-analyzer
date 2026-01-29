# Ejecutar exportación completa (sin importar módulo)

param(
    [string]$AnalyzerPath = "$PSScriptRoot\..\AccessAnalyzer.accdb",
    [string]$TestDbPath = "C:\export\test\appGraz.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "EXPORTACION COMPLETA DE ACCESS" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$access = $null

try {
    Write-Host "1. Abriendo AccessAnalyzer..." -ForegroundColor Yellow
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AnalyzerPath, $false)
    
    Write-Host "   OK" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "2. Ejecutando exportación completa..." -ForegroundColor Yellow
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFolder = "C:\export\test\Exportacion_COMPLETA_$timestamp"
    
    Write-Host "   Base: $TestDbPath" -ForegroundColor Cyan
    Write-Host "   Salida: $outputFolder" -ForegroundColor Cyan
    Write-Host ""
    
    # Construir comando
    $dbEscaped = $TestDbPath.Replace('\', '\\')
    $outEscaped = $outputFolder.Replace('\', '\\')
    $cmd = 'RunCompleteExport("' + $dbEscaped + '","' + $outEscaped + '")'
    
    Write-Host "   Ejecutando..." -ForegroundColor Yellow
    
    $result = $access.Eval($cmd)
    
    if ($result) {
        Write-Host ""
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host "EXPORTACION EXITOSA!" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Carpeta: $outputFolder" -ForegroundColor White
        
        if (Test-Path "$outputFolder\00_RESUMEN.txt") {
            Write-Host ""
            Write-Host "RESUMEN:" -ForegroundColor Cyan
            Get-Content "$outputFolder\00_RESUMEN.txt"
        }
        
        Write-Host ""
        Start-Process explorer.exe $outputFolder
    }
    else {
        Write-Host ""
        Write-Host "ERROR en la exportación" -ForegroundColor Red
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
