# ============================================================================
# Script para agregar AccessAnalyzer a ubicaciones confiables de Access
# ============================================================================
# Problema: Access deshabilita VBA cuando se abre desde PowerShell/COM
# Solución: Agregar la carpeta a ubicaciones confiables del registro
# ============================================================================

param(
    [string]$AccessAnalyzerPath = "$PSScriptRoot\..\assets\AccessAnalyzer.accdb"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "CONFIGURAR UBICACION CONFIABLE EN ACCESS" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar que el archivo existe
if (-not (Test-Path $AccessAnalyzerPath)) {
    Write-Host "ERROR: No se encuentra AccessAnalyzer.accdb" -ForegroundColor Red
    Write-Host "Ruta esperada: $AccessAnalyzerPath" -ForegroundColor Yellow
    exit 1
}

$analyzerFolder = Split-Path -Parent $AccessAnalyzerPath
Write-Host "Carpeta a agregar como confiable:" -ForegroundColor Cyan
Write-Host "$analyzerFolder" -ForegroundColor White
Write-Host ""

# Rutas del registro para diferentes versiones de Access
$regPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations",  # Office 2016/2019/Microsoft 365
    "HKCU:\Software\Microsoft\Office\15.0\Access\Security\Trusted Locations",  # Office 2013
    "HKCU:\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations"   # Office 2010
)

# Buscar la versión de Access instalada
$accessVersion = $null
foreach ($regPath in $regPaths) {
    if (Test-Path $regPath) {
        $accessVersion = $regPath
        Write-Host "? Versión de Access detectada" -ForegroundColor Green
        break
    }
}

if (-not $accessVersion) {
    Write-Host "ERROR: No se encontró Access instalado o registry accesible" -ForegroundColor Red
    Write-Host ""
    Write-Host "Solución manual:" -ForegroundColor Yellow
    Write-Host "1. Abre Access" -ForegroundColor Yellow
    Write-Host "2. Ve a: Archivo ? Opciones ? Centro de confianza ? Configuración del Centro de confianza" -ForegroundColor Yellow
    Write-Host "3. En 'Ubicaciones de confianza', haz clic en 'Agregar nueva ubicación'" -ForegroundColor Yellow
    Write-Host "4. Selecciona la carpeta: $analyzerFolder" -ForegroundColor Yellow
    Write-Host "5. Marca 'Confiar en todas las subcarpetas de esta ubicación'" -ForegroundColor Yellow
    exit 1
}

Write-Host "Ruta del registro: $accessVersion" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "Agregando ubicación confiable..." -ForegroundColor Yellow
    Write-Host ""
    
    # Crear la clave si no existe
    if (-not (Test-Path $accessVersion)) {
        Write-Host "Creando clave del registro..." -ForegroundColor Cyan
        New-Item -Path $accessVersion -Force | Out-Null
    }
    
    # Crear entrada para AccessAnalyzer
    $locationName = "AccessAnalyzer"
    $locationPath = Join-Path $accessVersion $locationName
    
    # Si ya existe, remover primero para evitar conflictos
    if (Test-Path $locationPath) {
        Write-Host "Removiendo entrada anterior..." -ForegroundColor Cyan
        Remove-Item -Path $locationPath -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }
    
    # Crear entrada nueva
    Write-Host "Creando nueva entrada en el registro..." -ForegroundColor Cyan
    New-Item -Path $locationPath -Force | Out-Null
    
    # Establecer propiedades
    Write-Host "Configurando propiedades..." -ForegroundColor Cyan
    
    # Path: Ruta de la carpeta
    New-ItemProperty `
        -Path $locationPath `
        -Name "Path" `
        -Value $analyzerFolder `
        -PropertyType String `
        -Force | Out-Null
    
    # AllowSubfolders: Permitir subcarpetas (1 = true)
    New-ItemProperty `
        -Path $locationPath `
        -Name "AllowSubfolders" `
        -Value 1 `
        -PropertyType DWORD `
        -Force | Out-Null
    
    # Description: Descripción
    New-ItemProperty `
        -Path $locationPath `
        -Name "Description" `
        -Value "GitHub Copilot Access Analyzer - Ubicación para scripts automatizados" `
        -PropertyType String `
        -Force | Out-Null
    
    Write-Host ""
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "? UBICACION CONFIABLE AGREGADA" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Detalles:" -ForegroundColor Cyan
    Write-Host "  Carpeta: $analyzerFolder" -ForegroundColor White
    Write-Host "  Subcarpetas: Habilitadas" -ForegroundColor White
    Write-Host "  Descripción: GitHub Copilot Access Analyzer" -ForegroundColor White
    Write-Host ""
    Write-Host "Ahora puede ejecutar los scripts sin advertencias de seguridad:" -ForegroundColor Cyan
    Write-Host "  .\access-export-git.ps1 -DatabasePath <ruta-bd>" -ForegroundColor White
    Write-Host ""
    
    exit 0
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Detalles:" -ForegroundColor Yellow
    Write-Host $_.Exception.StackTrace -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Si el error es de permisos, ejecuta este script como Administrador" -ForegroundColor Yellow
    exit 1
}
