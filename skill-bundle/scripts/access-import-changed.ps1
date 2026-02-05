# Script para importar solo archivos modificados (detectados con Git)

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDbPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ExportFolder,
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb",
    
    [switch]$DryRun = $false
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "IMPORTACION INTELIGENTE (SOLO CAMBIOS)" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar Git
if (-not (Test-Path (Join-Path $ExportFolder ".git"))) {
    Write-Host "ERROR: $ExportFolder no es un repositorio Git" -ForegroundColor Red
    Write-Host "Ejecuta primero: access-export-git.ps1" -ForegroundColor Yellow
    exit 1
}

# Detectar archivos modificados
Write-Host "1. Detectando archivos modificados..." -ForegroundColor Yellow

Push-Location $ExportFolder

# Ver archivos modificados desde el último commit
$modifiedFiles = git diff --name-only HEAD~1 HEAD 2>$null

if (-not $modifiedFiles) {
    Write-Host "   No hay cambios desde la última exportación" -ForegroundColor Yellow
    Pop-Location
    exit 0
}

Write-Host "   Archivos modificados: $($modifiedFiles.Count)" -ForegroundColor Green
Write-Host ""

# Clasificar archivos por tipo
$queries = @()
$forms = @()
$reports = @()
$macros = @()
$modules = @()

foreach ($file in $modifiedFiles) {
    # Git usa / en vez de \, normalizar
    $normalizedFile = $file -replace '/', '\'
    
    if ($normalizedFile -match '^02_Consultas\\(.+)\.txt$') {
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
        $modules += @{Name = $Matches[1]; Ext = $Matches[2]}
    }
}

Write-Host "Cambios detectados:" -ForegroundColor Cyan
Write-Host "  Consultas: $($queries.Count)" -ForegroundColor White
Write-Host "  Formularios: $($forms.Count)" -ForegroundColor White
Write-Host "  Informes: $($reports.Count)" -ForegroundColor White
Write-Host "  Macros: $($macros.Count)" -ForegroundColor White
Write-Host "  Módulos VBA: $($modules.Count)" -ForegroundColor White
Write-Host ""

Pop-Location

if ($DryRun) {
    Write-Host "=== DRY RUN ===" -ForegroundColor Yellow
    Write-Host "Se importarían:" -ForegroundColor Yellow
    $queries | ForEach-Object { Write-Host "  [Query] $_" -ForegroundColor Cyan }
    $forms | ForEach-Object { Write-Host "  [Form] $_" -ForegroundColor Cyan }
    $reports | ForEach-Object { Write-Host "  [Report] $_" -ForegroundColor Cyan }
    $macros | ForEach-Object { Write-Host "  [Macro] $_" -ForegroundColor Cyan }
    $modules | ForEach-Object { Write-Host "  [Module] $_" -ForegroundColor Cyan }
    exit 0
}

# Crear backup
Write-Host "2. Creando backup..." -ForegroundColor Yellow
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = $TargetDbPath -replace '\.accdb$', "_BACKUP_$timestamp.accdb"
Copy-Item $TargetDbPath $backupPath -Force
Write-Host "   OK: $backupPath" -ForegroundColor Green

# Importar cambios
Write-Host ""
Write-Host "3. Importando cambios..." -ForegroundColor Yellow

$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($TargetDbPath, $false)
    
    $imported = 0
    $errors = 0
    
    # Importar consultas
    foreach ($queryName in $queries) {
        $filePath = Join-Path $ExportFolder "02_Consultas\$queryName.txt"
        if (Test-Path $filePath) {
            Write-Host "   [Query] $queryName..." -NoNewline
            try {
                $access.DoCmd.DeleteObject([Microsoft.Office.Interop.Access.AcObjectType]::acQuery, $queryName)
                $access.LoadFromText([Microsoft.Office.Interop.Access.AcObjectType]::acQuery, $queryName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                Write-Host " ERROR" -ForegroundColor Red
                $errors++
            }
        }
    }
    
    # Importar formularios
    foreach ($formName in $forms) {
        $filePath = Join-Path $ExportFolder "03_Formularios\$formName.txt"
        if (Test-Path $filePath) {
            Write-Host "   [Form] $formName..." -NoNewline
            try {
                $access.DoCmd.DeleteObject([Microsoft.Office.Interop.Access.AcObjectType]::acForm, $formName)
                $access.LoadFromText([Microsoft.Office.Interop.Access.AcObjectType]::acForm, $formName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                Write-Host " ERROR" -ForegroundColor Red
                $errors++
            }
        }
    }
    
    # Importar informes
    foreach ($reportName in $reports) {
        $filePath = Join-Path $ExportFolder "04_Informes\$reportName.txt"
        if (Test-Path $filePath) {
            Write-Host "   [Report] $reportName..." -NoNewline
            try {
                $access.DoCmd.DeleteObject([Microsoft.Office.Interop.Access.AcObjectType]::acReport, $reportName)
                $access.LoadFromText([Microsoft.Office.Interop.Access.AcObjectType]::acReport, $reportName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                Write-Host " ERROR" -ForegroundColor Red
                $errors++
            }
        }
    }
    
    # Importar macros
    foreach ($macroName in $macros) {
        $filePath = Join-Path $ExportFolder "05_Macros\$macroName.txt"
        if (Test-Path $filePath) {
            Write-Host "   [Macro] $macroName..." -NoNewline
            try {
                $access.DoCmd.DeleteObject([Microsoft.Office.Interop.Access.AcObjectType]::acMacro, $macroName)
                $access.LoadFromText([Microsoft.Office.Interop.Access.AcObjectType]::acMacro, $macroName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                Write-Host " ERROR" -ForegroundColor Red
                $errors++
            }
        }
    }
    
    # Importar módulos VBA (respetando .cls y .bas)
    foreach ($module in $modules) {
        $moduleName = $module.Name
        $moduleExt = $module.Ext
        $filePath = Join-Path $ExportFolder "06_Codigo_VBA\$moduleName.$moduleExt"
        
        if (Test-Path $filePath) {
            $moduleType = if ($moduleExt -eq 'cls') { 'Class' } else { 'Module' }
            Write-Host "   [$moduleType] $moduleName..." -NoNewline
            try {
                $access.DoCmd.DeleteObject([Microsoft.Office.Interop.Access.AcObjectType]::acModule, $moduleName)
                $access.LoadFromText([Microsoft.Office.Interop.Access.AcObjectType]::acModule, $moduleName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                Write-Host " ERROR: $_" -ForegroundColor Red
                $errors++
            }
        }
    }
    
    Write-Host ""
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "IMPORTACION COMPLETADA" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Objetos importados: $imported" -ForegroundColor Green
    if ($errors -gt 0) {
        Write-Host "Errores: $errors" -ForegroundColor Red
    }
    Write-Host "Backup: $backupPath" -ForegroundColor Gray
    
    # Preguntar si quiere abrir la base de datos refactorizada
    Write-Host ""
    $openAccess = Read-Host "¿Abrir base de datos refactorizada en Access? (S/N)"
    
    if ($openAccess -eq 'S' -or $openAccess -eq 's' -or $openAccess -eq 'Y' -or $openAccess -eq 'y') {
        Write-Host "Abriendo Access..." -ForegroundColor Yellow
        Start-Process "$TargetDbPath"
    }
}
catch {
    Write-Host ""
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
