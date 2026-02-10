# Script para importar archivos exportados usando ModImportComplete.bas via AccessAnalyzer.accdb

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDbPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ExportFolder,
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "IMPORTACION SELECTIVA VIA ACCESSANALYZER (VBA)" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar rutas
Write-Host "1. Validando archivos..." -ForegroundColor Yellow

if (-not (Test-Path $TargetDbPath)) {
    Write-Host "ERROR: Base de datos destino no encontrada: $TargetDbPath" -ForegroundColor Red
    exit 1
}
Write-Host "   BD destino: OK" -ForegroundColor Green

if (-not (Test-Path $ExportFolder)) {
    Write-Host "ERROR: Carpeta de exportación no encontrada: $ExportFolder" -ForegroundColor Red
    exit 1
}
Write-Host "   Carpeta exportación: OK" -ForegroundColor Green

if (-not (Test-Path $AnalyzerPath)) {
    Write-Host "ERROR: AccessAnalyzer.accdb no encontrado: $AnalyzerPath" -ForegroundColor Red
    exit 1
}
Write-Host "   AccessAnalyzer.accdb: OK" -ForegroundColor Green
Write-Host ""
# Detectar cambios con Git
Write-Host "2. Detectando cambios con Git..." -ForegroundColor Yellow
if (-not (Test-Path (Join-Path $ExportFolder ".git"))) {
    Write-Host "ERROR: No se encontró .git en la carpeta de exportación." -ForegroundColor Red
    exit 1
}

Push-Location $ExportFolder
$commitCount = (git rev-list --count HEAD 2>$null)
if ($commitCount -and [int]$commitCount -gt 1) {
    $changedFiles = git diff --name-only HEAD~1 HEAD 2>$null
} else {
    $changedFiles = git status --porcelain 2>$null | ForEach-Object { $_.Substring(3).Trim() }
}
Pop-Location

if (-not $changedFiles -or $changedFiles.Count -eq 0) {
    Write-Host "   Sin cambios detectados." -ForegroundColor Green
    exit 0
}

$queries = @()
$forms = @()
$reports = @()
$macros = @()
$modules = @()

foreach ($file in $changedFiles) {
    $normalizedFile = $file -replace '/', '\\'
    if ($normalizedFile -match '^02_Consultas\\(.+)\.(sql|txt)$') {
        $queries += $Matches[1]
    }
    elseif ($normalizedFile -match '^03_Formularios\\(.+)\.txt$') {
        $forms += $Matches[1]
    }
    elseif ($normalizedFile -match '^04_Informes\\(.+)\.txt$') {
        $reports += $Matches[1]
    }
    elseif ($normalizedFile -match '^05_Macros\\(.+)\.txt$') {
        $macros += $Matches[1]
    }
    elseif ($normalizedFile -match '^06_Codigo_VBA\\(.+)\.(bas|cls)$') {
        $modules += $Matches[1]
    }
}

$queries = $queries | Sort-Object -Unique
$forms = $forms | Sort-Object -Unique
$reports = $reports | Sort-Object -Unique
$macros = $macros | Sort-Object -Unique
$modules = $modules | Sort-Object -Unique

Write-Host "   Cambios detectados:" -ForegroundColor Cyan
Write-Host "     Consultas: $($queries.Count)" -ForegroundColor White
Write-Host "     Formularios: $($forms.Count)" -ForegroundColor White
Write-Host "     Informes: $($reports.Count)" -ForegroundColor White
Write-Host "     Macros: $($macros.Count)" -ForegroundColor White
Write-Host "     Modulos VBA: $($modules.Count)" -ForegroundColor White
Write-Host ""

function Escape-VbaString {
    param([string]$value)
    return $value.Replace('"', '""')
}

$queryCsv = ($queries -join ",")
$formCsv = ($forms -join ",")
$reportCsv = ($reports -join ",")
$macroCsv = ($macros -join ",")
$moduleCsv = ($modules -join ",")

# Crear backup de la BD destino
Write-Host "3. Creando backup..." -ForegroundColor Yellow
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = $TargetDbPath -replace '\.accdb$', "_BACKUP_$timestamp.accdb"
Copy-Item $TargetDbPath $backupPath -Force
Write-Host "   OK: $backupPath" -ForegroundColor Green
Write-Host ""

# Ejecutar importación via AccessAnalyzer.accdb usando ModImportComplete.bas
Write-Host "4. Ejecutando importación via AccessAnalyzer..." -ForegroundColor Yellow

$shouldOpen = $false
$job = $null

try {
    Write-Host "   Abriendo AccessAnalyzer.accdb..." -NoNewline
    Write-Host " OK" -ForegroundColor Green

    $logPath = Join-Path $ExportFolder "00_LOG_IMPORTACION.txt"
    $targetEscaped = Escape-VbaString($TargetDbPath.Replace('\', '\\'))
    $folderEscaped = Escape-VbaString($ExportFolder.Replace('\', '\\'))
    $queryEscaped = Escape-VbaString($queryCsv)
    $formEscaped = Escape-VbaString($formCsv)
    $reportEscaped = Escape-VbaString($reportCsv)
    $macroEscaped = Escape-VbaString($macroCsv)
    $moduleEscaped = Escape-VbaString($moduleCsv)
    $cmd = 'RunSelectedImport("' + $targetEscaped + '","' + $folderEscaped + '","' + $Language + '","' + $queryEscaped + '","' + $formEscaped + '","' + $reportEscaped + '","' + $macroEscaped + '","' + $moduleEscaped + '")'

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

    Write-Host "   Ejecutando ModImportComplete.RunSelectedImport()..." -ForegroundColor Yellow

    $waitCount = 0
    while (-not (Test-Path $logPath) -and $waitCount -lt 20) {
        Start-Sleep -Milliseconds 500
        $waitCount++
    }

    if (Test-Path $logPath) {
        Write-Host "   ---------------------------------------------------------------" -ForegroundColor DarkGray
        $lastLine = ""
        do {
            Start-Sleep -Milliseconds 300
            if (Test-Path $logPath) {
                $lines = Get-Content $logPath -Tail 5 -ErrorAction SilentlyContinue
                if ($lines) {
                    $currentLine = $lines[-1]
                    if ($currentLine -ne $lastLine) {
                        if ($currentLine -match '\[ERROR') {
                            Write-Host "   $currentLine" -ForegroundColor Red
                        } elseif ($currentLine -match '\[(\d+:\d+)\]') {
                            Write-Host "   $currentLine" -ForegroundColor Cyan
                        } elseif ($currentLine -match 'OK:') {
                            Write-Host "   $currentLine" -ForegroundColor Green
                        } else {
                            Write-Host "   $currentLine" -ForegroundColor Gray
                        }
                        $lastLine = $currentLine
                    }
                }
            }
        } while ($job.State -eq 'Running')
        Write-Host "   ---------------------------------------------------------------" -ForegroundColor DarkGray
    }

    $jobOutput = Receive-Job -Job $job -Wait -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "IMPORTACION COMPLETADA" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Backup: $backupPath" -ForegroundColor Gray

    if (Test-Path $logPath) {
        Write-Host "Log: $logPath" -ForegroundColor Gray
    }

    Write-Host ""
    $openAccess = Read-Host "¿Abrir base de datos importada en Access? (S/N)"

    if ($openAccess -eq 'S' -or $openAccess -eq 's' -or $openAccess -eq 'Y' -or $openAccess -eq 'y') {
        $shouldOpen = $true
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($job) {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }

    if ($shouldOpen) {
        Write-Host "Abriendo Access..." -ForegroundColor Yellow
        Start-Process "$TargetDbPath"
    }
}
