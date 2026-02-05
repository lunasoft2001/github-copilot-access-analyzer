# ============================================================================
# Script para actualizar skill access-analyzer instalado localmente
# ============================================================================

param(
    [string]$RepositoryPath = "E:\datos\GitHub\github-copilot-access-analyzer",
    [string]$SkillPath = "C:\Users\jjAzure\.copilot\skills\access-analyzer"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "ACTUALIZAR SKILL access-analyzer" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar directorios
if (-not (Test-Path $RepositoryPath)) {
    Write-Host "ERROR: No se encuentra repositorio" -ForegroundColor Red
    Write-Host "Ruta: $RepositoryPath" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $SkillPath)) {
    Write-Host "ERROR: No se encuentra skill instalado" -ForegroundColor Red
    Write-Host "Ruta: $SkillPath" -ForegroundColor Yellow
    exit 1
}

Write-Host "?? Repositorio: $RepositoryPath" -ForegroundColor Green
Write-Host "?? Skill local:  $SkillPath" -ForegroundColor Green
Write-Host ""

try {
    Write-Host "Copiando archivos..." -ForegroundColor Yellow
    Write-Host ""
    
    # Copiar skill-bundle a skill
    $source = Join-Path $RepositoryPath "skill-bundle"
    
    # Copiar scripts
    Write-Host "  • Scripts PowerShell..." -ForegroundColor Cyan
    Copy-Item "$source/scripts/*.ps1" "$SkillPath/scripts/" -Force
    Write-Host "    ? $(Get-ChildItem "$SkillPath/scripts/*.ps1" | Measure-Object | Select-Object -ExpandProperty Count) scripts actualizados" -ForegroundColor Green
    
    # Copiar SKILL.md
    Write-Host "  • Documentación del skill..." -ForegroundColor Cyan
    Copy-Item "$source/SKILL.md" "$SkillPath/SKILL.md" -Force
    Write-Host "    ? SKILL.md actualizado" -ForegroundColor Green
    
    # Copiar referencias
    Write-Host "  • Referencias y módulos..." -ForegroundColor Cyan
    Copy-Item "$source/references/*.*" "$SkillPath/references/" -Force -Recurse
    Write-Host "    ? Referencias actualizadas" -ForegroundColor Green
    
    # Copiar assets
    Write-Host "  • Assets..." -ForegroundColor Cyan
    Copy-Item "$source/assets/*.*" "$SkillPath/assets/" -Force -Recurse
    Write-Host "    ? Assets actualizados" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "? SKILL ACTUALIZADO CORRECTAMENTE" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Cambios aplicados:" -ForegroundColor Cyan
    Write-Host "  • Rutas de AnalyzerPath corregidas" -ForegroundColor White
    Write-Host "  • Nuevos scripts de configuración (setup-trusted-location.ps1)" -ForegroundColor White
    Write-Host "  • Script mejorado de exportación (access-export-git-FIXED.ps1)" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Próximos pasos:" -ForegroundColor Yellow
    Write-Host "  1. Ejecutar como Administrador:" -ForegroundColor White
    Write-Host "     & 'C:\Users\jjAzure\.copilot\skills\access-analyzer\scripts\setup-trusted-location.ps1'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  2. Usar normalmente:" -ForegroundColor White
    Write-Host "     & '.\scripts\access-export-git.ps1' -DatabasePath 'ruta\a\tu\base.accdb'" -ForegroundColor Cyan
    Write-Host ""
    
    exit 0
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Detalles:" -ForegroundColor Yellow
    Write-Host $_.Exception.StackTrace -ForegroundColor Yellow
    exit 1
}
