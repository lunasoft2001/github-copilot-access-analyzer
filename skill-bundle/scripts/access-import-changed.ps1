# Script para importar solo archivos modificados (detectados con Git)

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDbPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ExportFolder,
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb",
    
    [switch]$DryRun = $false,

    [string[]]$QueryNames,
    [string[]]$FormNames,
    [string[]]$ReportNames,
    [string[]]$MacroNames,
    [string[]]$ModuleNames,

    [switch]$Interactive = $false
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "IMPORTACION INTELIGENTE (SOLO CAMBIOS)" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

function Resolve-AccessFile {
    param(
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string]$BaseName,
        [Parameter(Mandatory=$true)][string[]]$Extensions
    )

    foreach ($ext in $Extensions) {
        $candidate = Join-Path $Folder ("$BaseName.$ext")
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Prompt-Names {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string[]]$Extensions
    )

    Write-Host "" 
    Write-Host $Title -ForegroundColor Cyan
    $items = @()

    foreach ($ext in $Extensions) {
        $items += Get-ChildItem -Path $Folder -Filter "*.$ext" -ErrorAction SilentlyContinue
    }

    $items = $items | Sort-Object Name -Unique
    if (-not $items) {
        Write-Host "  (sin archivos)" -ForegroundColor DarkGray
        return @()
    }

    $items | ForEach-Object { Write-Host "  - $($_.BaseName)" -ForegroundColor White }
    $raw = Read-Host "Escribe uno o varios nombres (separados por coma) o Enter para omitir"
    if (-not $raw) {
        return @()
    }

    return $raw.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
}

$queries = @()
$forms = @()
$reports = @()
$macros = @()
$modules = @()

$useManual = $Interactive -or $QueryNames -or $FormNames -or $ReportNames -or $MacroNames -or $ModuleNames

if ($useManual) {
    Write-Host "1. Seleccion manual de objetos..." -ForegroundColor Yellow

    if ($Interactive) {
        $queries = Prompt-Names "Consultas disponibles" (Join-Path $ExportFolder "02_Consultas") @("sql", "txt")
        $modules = Prompt-Names "Modulos VBA disponibles" (Join-Path $ExportFolder "06_Codigo_VBA") @("bas", "cls") |
            ForEach-Object { @{ Name = $_; Ext = $null } }
        $forms = Prompt-Names "Formularios disponibles" (Join-Path $ExportFolder "03_Formularios") @("txt")
        $reports = Prompt-Names "Informes disponibles" (Join-Path $ExportFolder "04_Informes") @("txt")
        $macros = Prompt-Names "Macros disponibles" (Join-Path $ExportFolder "05_Macros") @("txt")
    }

    if ($QueryNames) { $queries += $QueryNames }
    if ($FormNames) { $forms += $FormNames }
    if ($ReportNames) { $reports += $ReportNames }
    if ($MacroNames) { $macros += $MacroNames }
    if ($ModuleNames) { $modules += $ModuleNames | ForEach-Object { @{ Name = $_; Ext = $null } } }
}
else {
    # Validar Git
    if (-not (Test-Path (Join-Path $ExportFolder ".git"))) {
        Write-Host "ERROR: $ExportFolder no es un repositorio Git" -ForegroundColor Red
        Write-Host "Ejecuta primero: access-export-git.ps1 o usa -Interactive" -ForegroundColor Yellow
        exit 1
    }

    # Detectar archivos modificados
    Write-Host "1. Detectando archivos modificados..." -ForegroundColor Yellow

    Push-Location $ExportFolder

    # Ver archivos modificados desde el ultimo commit
    $modifiedFiles = git diff --name-only HEAD~1 HEAD 2>$null

    if (-not $modifiedFiles) {
        Write-Host "   No hay cambios desde la ultima exportacion" -ForegroundColor Yellow
        Pop-Location
        exit 0
    }

    Write-Host "   Archivos modificados: $($modifiedFiles.Count)" -ForegroundColor Green
    Write-Host ""

    # Clasificar archivos por tipo
    foreach ($file in $modifiedFiles) {
        # Git usa / en vez de \, normalizar
        $normalizedFile = $file -replace '/', '\\'
        
        if ($normalizedFile -match '^02_Consultas\\(.+)\.(txt|sql)$') {
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
    Write-Host "  Modulos VBA: $($modules.Count)" -ForegroundColor White
    Write-Host ""

    Pop-Location
}

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
$shouldOpen = $false

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($TargetDbPath, $false)

    $acQuery = 1
    $acForm = 2
    $acReport = 3
    $acMacro = 4
    $acModule = 5
    
    $imported = 0
    $errors = 0
    
    # Importar consultas
    foreach ($queryName in $queries) {
        $filePath = Resolve-AccessFile (Join-Path $ExportFolder "02_Consultas") $queryName @("sql", "txt")
        if (Test-Path $filePath) {
            Write-Host "   [Query] $queryName..." -NoNewline
            try {
                $ext = [System.IO.Path]::GetExtension($filePath).TrimStart('.')

                if ($ext -eq 'sql') {
                    $sqlText = Get-Content -Path $filePath -Raw -Encoding UTF8
                    # Remove line comments (Access SQL does not accept -- comments)
                    $sqlText = ($sqlText -split "`r?`n") | Where-Object { $_ -notmatch '^\s*--' } | ForEach-Object { $_.TrimEnd() }
                    $sqlText = ($sqlText -join "`r`n").Trim()
                    $db = $null
                    try {
                        $db = $access.CurrentDb()
                        try { $db.QueryDefs.Delete($queryName) } catch { }
                        $null = $db.CreateQueryDef($queryName, $sqlText)
                    }
                    finally {
                        if ($db) {
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($db) | Out-Null
                        }
                    }
                }
                else {
                    try { $access.DoCmd.DeleteObject($acQuery, $queryName) } catch { }
                    $access.LoadFromText($acQuery, $queryName, $filePath)
                }
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Host " ERROR: $errMsg" -ForegroundColor Red
                Write-Host "     Archivo: $filePath" -ForegroundColor DarkGray
                $errors++
            }
        }
    }
    
    # Importar formularios
    foreach ($formName in $forms) {
        $filePath = Resolve-AccessFile (Join-Path $ExportFolder "03_Formularios") $formName @("txt")
        if (Test-Path $filePath) {
            Write-Host "   [Form] $formName..." -NoNewline
            try {
                try { $access.DoCmd.DeleteObject($acForm, $formName) } catch { }
                $access.LoadFromText($acForm, $formName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Host " ERROR: $errMsg" -ForegroundColor Red
                Write-Host "     Archivo: $filePath" -ForegroundColor DarkGray
                $errors++
            }
        }
    }
    
    # Importar informes
    foreach ($reportName in $reports) {
        $filePath = Resolve-AccessFile (Join-Path $ExportFolder "04_Informes") $reportName @("txt")
        if (Test-Path $filePath) {
            Write-Host "   [Report] $reportName..." -NoNewline
            try {
                try { $access.DoCmd.DeleteObject($acReport, $reportName) } catch { }
                $access.LoadFromText($acReport, $reportName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Host " ERROR: $errMsg" -ForegroundColor Red
                Write-Host "     Archivo: $filePath" -ForegroundColor DarkGray
                $errors++
            }
        }
    }
    
    # Importar macros
    foreach ($macroName in $macros) {
        $filePath = Resolve-AccessFile (Join-Path $ExportFolder "05_Macros") $macroName @("txt")
        if (Test-Path $filePath) {
            Write-Host "   [Macro] $macroName..." -NoNewline
            try {
                try { $access.DoCmd.DeleteObject($acMacro, $macroName) } catch { }
                $access.LoadFromText($acMacro, $macroName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Host " ERROR: $errMsg" -ForegroundColor Red
                Write-Host "     Archivo: $filePath" -ForegroundColor DarkGray
                $errors++
            }
        }
    }
    
    # Importar módulos VBA (respetando .cls y .bas)
    foreach ($module in $modules) {
        $moduleName = $module.Name
        $moduleExt = $module.Ext
        if (-not $moduleExt) {
            $moduleExt = if (Test-Path (Join-Path $ExportFolder "06_Codigo_VBA\$moduleName.cls")) { "cls" } else { "bas" }
        }
        $filePath = Join-Path $ExportFolder "06_Codigo_VBA\$moduleName.$moduleExt"
        
        if (Test-Path $filePath) {
            $moduleType = if ($moduleExt -eq 'cls') { 'Class' } else { 'Module' }
            Write-Host "   [$moduleType] $moduleName..." -NoNewline
            try {
                try { $access.DoCmd.DeleteObject($acModule, $moduleName) } catch { }
                $access.LoadFromText($acModule, $moduleName, $filePath)
                Write-Host " OK" -ForegroundColor Green
                $imported++
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Host " ERROR: $errMsg" -ForegroundColor Red
                Write-Host "     Archivo: $filePath" -ForegroundColor DarkGray
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
        $shouldOpen = $true
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

    if ($shouldOpen) {
        Write-Host "Abriendo Access..." -ForegroundColor Yellow
        Start-Process "$TargetDbPath"
    }
}
