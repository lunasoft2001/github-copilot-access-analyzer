<#
.SYNOPSIS
    Instala, actualiza o desinstala el skill access-analyzer para GitHub Copilot

.DESCRIPTION
    Este script automatiza la instalación del skill access-analyzer en la carpeta
    de skills de GitHub Copilot (~\.copilot\skills\access-analyzer).
    
    Incluye opciones para instalar, actualizar, desinstalar y verificar el skill.

.PARAMETER Action
    Acción a realizar: Install, Update, Uninstall, Verify
    Por defecto: Install

.EXAMPLE
    .\install-skill.ps1
    Instala el skill access-analyzer

.EXAMPLE
    .\install-skill.ps1 -Action Update
    Actualiza el skill access-analyzer existente

.EXAMPLE
    .\install-skill.ps1 -Action Uninstall
    Desinstala el skill access-analyzer

.EXAMPLE
    .\install-skill.ps1 -Action Verify
    Verifica que el skill esté correctamente instalado
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Install", "Update", "Uninstall", "Verify")]
    [string]$Action = "Install"
)

# Configuración
$SkillName = "access-analyzer"
$SkillBundlePath = Join-Path $PSScriptRoot "skill-bundle"
$CopilotSkillsPath = Join-Path $env:USERPROFILE ".copilot\skills"
$SkillInstallPath = Join-Path $CopilotSkillsPath $SkillName

# Colores para output
function Write-Success { param([string]$Message) Write-Host "? $Message" -ForegroundColor Green }
function Write-Info { param([string]$Message) Write-Host "? $Message" -ForegroundColor Cyan }
function Write-Warning { param([string]$Message) Write-Host "? $Message" -ForegroundColor Yellow }
function Write-Error-Custom { param([string]$Message) Write-Host "? $Message" -ForegroundColor Red }

# Banner
Write-Host ""
Write-Host "???????????????????????????????????????????????????????" -ForegroundColor Magenta
Write-Host "  GitHub Copilot Skill Installer - Access Analyzer" -ForegroundColor Magenta
Write-Host "???????????????????????????????????????????????????????" -ForegroundColor Magenta
Write-Host ""

# Función para verificar prerrequisitos
function Test-Prerequisites {
    Write-Info "Verificando prerrequisitos..."
    
    # Verificar que existe skill-bundle
    if (-not (Test-Path $SkillBundlePath)) {
        Write-Error-Custom "No se encontró la carpeta skill-bundle en: $SkillBundlePath"
        Write-Info "Asegúrate de ejecutar este script desde la raíz del repositorio."
        return $false
    }
    
    # Verificar que existe SKILL.md en el bundle
    $skillMdPath = Join-Path $SkillBundlePath "SKILL.md"
    if (-not (Test-Path $skillMdPath)) {
        Write-Error-Custom "No se encontró SKILL.md en el bundle: $skillMdPath"
        return $false
    }
    
    # Crear carpeta de skills si no existe
    if (-not (Test-Path $CopilotSkillsPath)) {
        Write-Warning "La carpeta de skills no existe, creándola: $CopilotSkillsPath"
        try {
            New-Item -ItemType Directory -Path $CopilotSkillsPath -Force | Out-Null
            Write-Success "Carpeta de skills creada"
        } catch {
            Write-Error-Custom "Error al crear carpeta de skills: $_"
            return $false
        }
    }
    
    Write-Success "Prerrequisitos verificados"
    return $true
}

# Función para instalar el skill
function Install-Skill {
    Write-Info "Instalando skill '$SkillName'..."
    
    # Verificar si ya existe
    if (Test-Path $SkillInstallPath) {
        Write-Warning "El skill '$SkillName' ya está instalado en: $SkillInstallPath"
        $response = Read-Host "¿Deseas sobrescribirlo? (S/N)"
        if ($response -notmatch '^[Ss]$') {
            Write-Info "Instalación cancelada"
            return $false
        }
        # Eliminar instalación anterior
        Remove-Item $SkillInstallPath -Recurse -Force
    }
    
    # Copiar skill-bundle completo
    try {
        Copy-Item -Path $SkillBundlePath -Destination $SkillInstallPath -Recurse -Force
        Write-Success "Skill instalado correctamente en: $SkillInstallPath"
        
        # Mostrar estructura instalada
        Write-Info "Estructura instalada:"
        Get-ChildItem $SkillInstallPath -Recurse | ForEach-Object {
            $relativePath = $_.FullName.Substring($SkillInstallPath.Length)
            Write-Host "  $relativePath" -ForegroundColor Gray
        }
        
        return $true
    } catch {
        Write-Error-Custom "Error al instalar el skill: $_"
        return $false
    }
}

# Función para actualizar el skill
function Update-Skill {
    Write-Info "Actualizando skill '$SkillName'..."
    
    if (-not (Test-Path $SkillInstallPath)) {
        Write-Warning "El skill '$SkillName' no está instalado"
        Write-Info "Ejecuta con -Action Install para instalarlo"
        return $false
    }
    
    # Crear backup de la instalación actual
    $backupPath = "$SkillInstallPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    try {
        Copy-Item -Path $SkillInstallPath -Destination $backupPath -Recurse -Force
        Write-Success "Backup creado en: $backupPath"
    } catch {
        Write-Warning "No se pudo crear backup: $_"
    }
    
    # Eliminar instalación actual
    Remove-Item $SkillInstallPath -Recurse -Force
    
    # Instalar nueva versión
    try {
        Copy-Item -Path $SkillBundlePath -Destination $SkillInstallPath -Recurse -Force
        Write-Success "Skill actualizado correctamente"
        return $true
    } catch {
        Write-Error-Custom "Error al actualizar: $_"
        # Restaurar backup si existe
        if (Test-Path $backupPath) {
            Write-Info "Restaurando backup..."
            Copy-Item -Path $backupPath -Destination $SkillInstallPath -Recurse -Force
            Write-Warning "Se restauró la versión anterior del backup"
        }
        return $false
    }
}

# Función para desinstalar el skill
function Uninstall-Skill {
    Write-Info "Desinstalando skill '$SkillName'..."
    
    if (-not (Test-Path $SkillInstallPath)) {
        Write-Warning "El skill '$SkillName' no está instalado"
        return $false
    }
    
    $response = Read-Host "¿Estás seguro de desinstalar el skill? (S/N)"
    if ($response -notmatch '^[Ss]$') {
        Write-Info "Desinstalación cancelada"
        return $false
    }
    
    try {
        Remove-Item $SkillInstallPath -Recurse -Force
        Write-Success "Skill desinstalado correctamente"
        return $true
    } catch {
        Write-Error-Custom "Error al desinstalar: $_"
        return $false
    }
}

# Función para verificar instalación
function Verify-Skill {
    Write-Info "Verificando instalación del skill '$SkillName'..."
    
    $issues = @()
    
    # Verificar carpeta principal
    if (-not (Test-Path $SkillInstallPath)) {
        $issues += "El skill no está instalado en: $SkillInstallPath"
    } else {
        Write-Success "Carpeta del skill encontrada"
        
        # Verificar SKILL.md
        $skillMd = Join-Path $SkillInstallPath "SKILL.md"
        if (-not (Test-Path $skillMd)) {
            $issues += "Falta SKILL.md"
        } else {
            Write-Success "SKILL.md encontrado"
            
            # Verificar frontmatter
            $content = Get-Content $skillMd -Raw
            if ($content -match '(?s)^---\s*\nname:\s*access-analyzer\s*\ndescription:.*?---') {
                Write-Success "Frontmatter correcto"
            } else {
                $issues += "Frontmatter de SKILL.md incorrecto o incompleto"
            }
        }
        
        # Verificar scripts/
        $scriptsPath = Join-Path $SkillInstallPath "scripts"
        if (-not (Test-Path $scriptsPath)) {
            $issues += "Falta carpeta scripts/"
        } else {
            $scriptCount = (Get-ChildItem $scriptsPath -Filter "*.ps1").Count
            Write-Success "Carpeta scripts/ encontrada ($scriptCount scripts)"
        }
        
        # Verificar references/
        $referencesPath = Join-Path $SkillInstallPath "references"
        if (-not (Test-Path $referencesPath)) {
            $issues += "Falta carpeta references/"
        } else {
            $refCount = (Get-ChildItem $referencesPath).Count
            Write-Success "Carpeta references/ encontrada ($refCount archivos)"
        }
        
        # Verificar assets/
        $assetsPath = Join-Path $SkillInstallPath "assets"
        if (-not (Test-Path $assetsPath)) {
            $issues += "Falta carpeta assets/"
        } else {
            $accdbPath = Join-Path $assetsPath "AccessAnalyzer.accdb"
            if (Test-Path $accdbPath) {
                Write-Success "Base de datos AccessAnalyzer.accdb encontrada"
            } else {
                $issues += "Falta AccessAnalyzer.accdb en assets/"
            }
        }
    }
    
    # Resumen
    Write-Host ""
    if ($issues.Count -eq 0) {
        Write-Success "? Skill correctamente instalado y verificado"
        Write-Host ""
        Write-Info "Reinicia VS Code para que Copilot detecte el skill"
        return $true
    } else {
        Write-Error-Custom "Se encontraron $($issues.Count) problema(s):"
        foreach ($issue in $issues) {
            Write-Host "  • $issue" -ForegroundColor Red
        }
        return $false
    }
}

# Main execution
try {
    if (-not (Test-Prerequisites)) {
        exit 1
    }
    
    Write-Host ""
    
    switch ($Action) {
        "Install" {
            if (Install-Skill) {
                Write-Host ""
                Write-Success "???????????????????????????????????????????????????????"
                Write-Success "  Instalación completada exitosamente"
                Write-Success "???????????????????????????????????????????????????????"
                Write-Host ""
                Write-Info "Próximos pasos:"
                Write-Host "  1. Cierra y reinicia VS Code" -ForegroundColor Yellow
                Write-Host "  2. El skill 'access-analyzer' estará disponible en Copilot" -ForegroundColor Yellow
                Write-Host "  3. Prueba preguntando: 'Exporta mi base de datos Access'" -ForegroundColor Yellow
                Write-Host ""
            } else {
                exit 1
            }
        }
        "Update" {
            if (Update-Skill) {
                Write-Host ""
                Write-Success "Actualización completada. Reinicia VS Code."
                Write-Host ""
            } else {
                exit 1
            }
        }
        "Uninstall" {
            if (Uninstall-Skill) {
                Write-Host ""
                Write-Success "Desinstalación completada"
                Write-Host ""
            } else {
                exit 1
            }
        }
        "Verify" {
            if (Verify-Skill) {
                exit 0
            } else {
                exit 1
            }
        }
    }
} catch {
    Write-Error-Custom "Error inesperado: $_"
    exit 1
}
