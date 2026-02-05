# Importar archivos modificados de vuelta a Access

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDbPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ImportFolder,
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb"
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
    
    # Determinar ruta del log
    $logPath = Join-Path $ImportFolder "00_LOG_IMPORTACION.txt"
    
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
    
    # Iniciar importación en background
    $job = Start-Job -ScriptBlock {
        param($AnalyzerPath, $Command)
        $access = New-Object -ComObject Access.Application
        $access.Visible = $false
        $access.OpenCurrentDatabase($AnalyzerPath, $false)
        $result = $access.Eval($Command)
        $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveNone)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        return $result
    } -ArgumentList $AnalyzerPath, $cmd
    
    # Esperar a que se cree el archivo de log
    Write-Host "   Iniciando importación..." -ForegroundColor Yellow
    $waitCount = 0
    while (-not (Test-Path $logPath) -and $waitCount -lt 20) {
        Start-Sleep -Milliseconds 500
        $waitCount++
    }
    
    # Mostrar progreso en tiempo real
    if (Test-Path $logPath) {
        Write-Host "   ────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        $lastLine = ""
        do {
            Start-Sleep -Milliseconds 300
            if (Test-Path $logPath) {
                $lines = Get-Content $logPath -Tail 5 -ErrorAction SilentlyContinue
                if ($lines) {
                    $currentLine = $lines[-1]
                    if ($currentLine -ne $lastLine) {
                        # Formatear línea de log
                        if ($currentLine -match '\[ERROR') {
                            Write-Host "   $currentLine" -ForegroundColor Red
                        } elseif ($currentLine -match '\[(\d+:\d+)\]') {
                            Write-Host "   $currentLine" -ForegroundColor Cyan
                        } elseif ($currentLine -match 'OK:') {
                            Write-Host "   $currentLine" -ForegroundColor Green
                        } elseif ($currentLine -match '=====') {
                            Write-Host "   $currentLine" -ForegroundColor Yellow
                        } else {
                            Write-Host "   $currentLine" -ForegroundColor Gray
                        }
                        $lastLine = $currentLine
                    }
                }
            }
        } while ($job.State -eq 'Running')
        Write-Host "   ────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    }
    
    # Esperar a que termine el job
    $result = Wait-Job $job | Receive-Job
    Remove-Job $job
    
    if ($result) {
        Write-Host ""
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host "IMPORTACION EXITOSA!" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Base de datos actualizada: $TargetDbPath" -ForegroundColor White
        Write-Host "Copia de seguridad: $backupPath" -ForegroundColor Gray
        Write-Host "Log completo: $logPath" -ForegroundColor Gray
    }
    else {
        Write-Host ""
        Write-Host "ERROR en la importación" -ForegroundColor Red
        Write-Host "Backup disponible en: $backupPath" -ForegroundColor Yellow
        Write-Host "Revisa el log: $logPath" -ForegroundColor Yellow
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Finalizado" -ForegroundColor Green
}
