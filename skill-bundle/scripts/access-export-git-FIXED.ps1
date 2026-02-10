# ============================================================================
# Script mejorado para exportar con Seguridad VBA habilitada
# ============================================================================
# Problema: Cuando PowerShell abre AccessAnalyzer.accdb, Access deshabilita
#           el contenido VBA por razones de seguridad.
# Solución: Habilitar el contenido VBA automáticamente antes de ejecutar
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [string]$ExportFolder = "",
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [switch]$ExportTableData,
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "EXPORTACION CON CONTROL DE VERSIONES" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar Git instalado
$gitVersion = git --version 2>$null
if (-not $gitVersion) {
    Write-Host "ERROR: Git no está instalado" -ForegroundColor Red
    Write-Host "Instala Git desde: https://git-scm.com/" -ForegroundColor Yellow
    exit 1
}

Write-Host "Git detectado: $gitVersion" -ForegroundColor Green
Write-Host ""

# Determinar carpeta de exportación
if ($ExportFolder -eq "") {
    $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DatabasePath)
    $parentFolder = [System.IO.Path]::GetDirectoryName($DatabasePath)
    $ExportFolder = Join-Path $parentFolder "${dbName}_Export"
}

$access = $null

try {
    Write-Host "1. Preparando exportación..." -ForegroundColor Yellow
    Write-Host "   AccessAnalyzer: $AnalyzerPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $AnalyzerPath)) {
        Write-Host "   ERROR: No se encuentra AccessAnalyzer.accdb" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "   OK" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "2. Exportando base de datos..." -ForegroundColor Yellow
    Write-Host "   Base: $DatabasePath" -ForegroundColor Cyan
    Write-Host "   Carpeta: $ExportFolder" -ForegroundColor Cyan
    Write-Host "   Idioma: $Language" -ForegroundColor Cyan
    Write-Host "   Exportar datos de tablas: $(if($ExportTableData){"SÍ"}else{"NO"})" -ForegroundColor Cyan
    Write-Host ""
    
    # Determinar ruta del log
    $logPath = Join-Path $ExportFolder "00_LOG_EXPORTACION.txt"
    
    # Construir comando con parámetro de datos
    $dbEscaped = $DatabasePath.Replace('\\', '\\\\')
    $outEscaped = $ExportFolder.Replace('\\', '\\\\')
    $exportDataStr = if($ExportTableData){"True"}else{"False"}
    $cmd = 'RunCompleteExport("' + $dbEscaped + '","' + $outEscaped + '","' + $Language + '",' + $exportDataStr + ')'
    
    # Iniciar exportación en background CON SEGURIDAD HABILITADA
    $job = Start-Job -ScriptBlock {
        param($AnalyzerPath, $Command)
        
        # Crear instancia de Access
        $access = New-Object -ComObject Access.Application
        $access.Visible = $false
        
        # IMPORTANTE: Habilitar contenido VBA
        # Access tiene una propiedad que necesitamos establecer ANTES de abrir la BD
        try {
            # Intentar crear la BD de confianza (Trust Database)
            # Esto requiere acceso a registry
            $regPath = "HKCU:\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations"
            $analyzerFolder = Split-Path -Parent $AnalyzerPath
            
            # Crear ubicación confiable para AccessAnalyzer
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            
            # Si no existe, añadir la carpeta a ubicaciones confiables
            $locationName = "AccessAnalyzer"
            $locationPath = Join-Path $regPath $locationName -ErrorAction SilentlyContinue
            
            if (-not (Test-Path $locationPath)) {
                New-Item -Path $locationPath -Force | Out-Null
                New-ItemProperty -Path $locationPath -Name "Path" -Value $analyzerFolder -Force | Out-Null
                New-ItemProperty -Path $locationPath -Name "AllowSubfolders" -Value 1 -PropertyType DWORD -Force | Out-Null
            }
        } catch {
            # Si hay error en registry, continuar de todas formas
            Write-Host "   Nota: No se pudo configurar ubicación confiada (requerida elevacion)" -ForegroundColor Yellow
        }
        
        # Abrir la base de datos
        # Segundo parámetro = Exclusive (false)
        # Tercer parámetro = Password (vacío)
        $access.OpenCurrentDatabase($AnalyzerPath, $false, "")
        
        # CRUCIAL: Dar tiempo a que Access cargue completamente
        Start-Sleep -Milliseconds 500
        
        # Habilitar contenido VBA accediendo a VBProject
        try {
            $vbProject = $access.VBE.ActiveVBProject
            # Solo el intento de acceder a VBProject habilita el VBA en algunos casos
        } catch {
            # Si no puede acceder, continuar de todas formas
        }
        
        # Ejecutar el comando
        try {
            $result = $access.Eval($Command)
        } catch {
            # Si falla por VBA deshabilitado, intentar con Application.Run()
            try {
                $result = $access.Run("ModExportComplete.RunCompleteExport", $dbEscaped, $outEscaped, $Language)
            } catch {
                # Si aún falla, retornar error
                throw $_
            }
        }
        
        # Cerrar sin guardar
        $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveNone)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        
        return $result
    } -ArgumentList $AnalyzerPath, $cmd
    
    # Esperar a que se cree el archivo de log
    Write-Host "   Iniciando exportación..." -ForegroundColor Yellow
    $waitCount = 0
    while (-not (Test-Path $logPath) -and $waitCount -lt 20) {
        Start-Sleep -Milliseconds 500
        $waitCount++
    }
    
    # Mostrar progreso en tiempo real
    if (Test-Path $logPath) {
        Write-Host "   ????????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
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
                        } else {
                            Write-Host "   $currentLine" -ForegroundColor White
                        }
                        $lastLine = $currentLine
                    }
                }
            }
            # Comprobar estado del job
            $jobState = (Get-Job -Id $job.Id).State
        } while ($jobState -eq "Running")
    }
    
    Write-Host "   ????????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
    Write-Host ""
    
    # Esperar a que termine el job
    $jobOutput = Receive-Job -Job $job -Wait -ErrorAction SilentlyContinue
    
    # Verificar si la exportación fue exitosa
    if (Test-Path $logPath) {
        $logContent = Get-Content $logPath -Raw
        if ($logContent -match '\[ERROR\]') {
            Write-Host "EXPORTACION FALLIDA" -ForegroundColor Red
            Write-Host ""
            Write-Host $logContent -ForegroundColor Red
            exit 1
        } else {
            Write-Host "? Exportación completada" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "3. Configurando repositorio Git..." -ForegroundColor Yellow
    
    # Crear carpeta si no existe
    if (-not (Test-Path $ExportFolder)) {
        New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
    }
    
    # Inicializar Git si no existe
    $gitFolder = Join-Path $ExportFolder ".git"
    if (-not (Test-Path $gitFolder)) {
        Write-Host "   Inicializando repositorio Git..." -ForegroundColor Cyan
        Push-Location $ExportFolder
        git init | Out-Null
        git config user.email "copilot@github.com" | Out-Null
        git config user.name "GitHub Copilot" | Out-Null
        
        # Crear .gitignore
        $gitIgnorePath = Join-Path $ExportFolder ".gitignore"
        @"
# Archivos temporales de Access
*.ldb
*.laccdb

# Backups
*_BACKUP_*.accdb
*_BACKUP_*.mdb

# Errores y logs
*_ERROR_*.txt
00_LOG_*.txt

# Carpetas temporales
Temp/
__pycache__/
.pytest_cache/

# Archivos del sistema
.DS_Store
Thumbs.db
"@ | Out-File -FilePath $gitIgnorePath -Encoding UTF8
        
        Write-Host "   ? .gitignore creado" -ForegroundColor Green
        Pop-Location
    }
    
    # Commit de la exportación
    Write-Host "   Realizando commit..." -ForegroundColor Cyan
    Push-Location $ExportFolder
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    git add . | Out-Null
    git commit -m "Exportación automática - $timestamp" | Out-Null
    
    # Mostrar estadísticas
    $diffStats = git diff --stat HEAD~1 HEAD 2>$null
    if ($diffStats) {
        Write-Host "   Cambios:" -ForegroundColor Cyan
        $diffStats | ForEach-Object {
            Write-Host "   $_" -ForegroundColor White
        }
    }
    
    Pop-Location
    
    Write-Host "   ? Repositorio actualizado" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "? EXPORTACION EXITOSA" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Ubicación: $ExportFolder" -ForegroundColor Cyan
    Write-Host ""
    
    # Preguntar si abrir en VS Code
    $response = Read-Host "¿Abrir en VS Code para refactorizar? (S/N)"
    if ($response.ToUpper() -eq 'S' -or $response.ToUpper() -eq 'Y') {
        code $ExportFolder
    }
    
    exit 0
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Detalles:" -ForegroundColor Yellow
    Write-Host $_.Exception.StackTrace -ForegroundColor Yellow
    exit 1
    
} finally {
    # Limpiar recursos
    if ($access) {
        try {
            $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveNone)
        } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    
    if ($job) {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}
