# Importar ModExportComplete.bas y probar exportación completa

param(
    [string]$AnalyzerPath = "$PSScriptRoot\..\AccessAnalyzer.accdb",
    [string]$TestDbPath = "C:\export\test\appGraz.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "IMPORTAR MODULO Y PROBAR EXPORTACION COMPLETA" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $AnalyzerPath)) {
    Write-Error "No se encuentra AccessAnalyzer.accdb"
    exit 1
}

if (-not (Test-Path $TestDbPath)) {
    Write-Error "No se encuentra base de datos de prueba"
    exit 1
}

$ModulePath = "$PSScriptRoot\..\modules\ModExportComplete.bas"

$access = $null

try {
    Write-Host "1. Abriendo AccessAnalyzer..." -ForegroundColor Yellow
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AnalyzerPath, $false)
    
    Write-Host "   OK" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "2. Verificando módulo existente..." -ForegroundColor Yellow
    
    for ($i = 1; $i -le $access.VBE.ActiveVBProject.VBComponents.Count; $i++) {
        $comp = $access.VBE.ActiveVBProject.VBComponents.Item($i)
        if ($comp.Name -eq "ModExportComplete") {
            Write-Host "   Eliminando módulo anterior..." -ForegroundColor Yellow
            $access.VBE.ActiveVBProject.VBComponents.Remove($comp)
            break
        }
    }
    
    Write-Host ""
    Write-Host "3. Importando ModExportComplete.bas..." -ForegroundColor Yellow
    
    try {
        $vbProj = $access.VBE.ActiveVBProject
        if ($null -eq $vbProj) {
            Write-Host "   ERROR: No se puede acceder al proyecto VBA" -ForegroundColor Red
            Write-Host "   Verifica que esté habilitado:" -ForegroundColor Yellow
            Write-Host "   'Confiar en el acceso al modelo de objetos de proyectos de VBA'" -ForegroundColor Yellow
            exit 1
        }
        
        $vbProj.VBComponents.Import($ModulePath) | Out-Null
    }
    catch {
        Write-Host "   ERROR importando módulo: $_" -ForegroundColor Red
        throw
    }
    
    Write-Host "   OK - Módulo importado" -ForegroundColor Green
    
    $access.DoCmd.Save()
    
    Write-Host ""
    Write-Host "4. Ejecutando exportación completa..." -ForegroundColor Yellow
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFolder = "C:\export\test\Exportacion_COMPLETA_$timestamp"
    
    Write-Host "   Base: $TestDbPath" -ForegroundColor Cyan
    Write-Host "   Salida: $outputFolder" -ForegroundColor Cyan
    
    # Usar comillas simples para construir el comando
    $dbEscaped = $TestDbPath.Replace('\', '\\')
    $outEscaped = $outputFolder.Replace('\', '\\')
    $cmd = 'RunCompleteExport("' + $dbEscaped + '","' + $outEscaped + '")'
    
    Write-Host ""
    Write-Host "   Comando: $cmd" -ForegroundColor Gray
    
    $result = $access.Eval($cmd)
    
    if ($result) {
        Write-Host ""
        Write-Host "EXPORTACION EXITOSA!" -ForegroundColor Green
        Write-Host $outputFolder -ForegroundColor White
        
        if (Test-Path "$outputFolder\00_RESUMEN.txt") {
            Write-Host ""
            Get-Content "$outputFolder\00_RESUMEN.txt" | Select-Object -First 20
        }
        
        Start-Process explorer.exe $outputFolder
    }
    else {
        Write-Host "ERROR en la exportación" -ForegroundColor Red
    }
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($access) {
        $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveAll)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
