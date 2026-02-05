# ============================================================================
# Script para inyectar ModExportComplete.bas en AccessAnalyzer.accdb
# ============================================================================
# Propósito: Corregir el error "RunCompleteExport function not found"
# Causa: ModExportComplete.bas no está importado en AccessAnalyzer.accdb
# Solución: Inyectar el módulo automáticamente
# ============================================================================

param(
    [string]$AccessAnalyzerPath = "$PSScriptRoot\skill-bundle\assets\AccessAnalyzer.accdb",
    [string]$ModuleSourcePath = "$PSScriptRoot\modules\ModExportComplete.bas"
)

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "INYECTANDO ModExportComplete EN AccessAnalyzer" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Validar archivos existen
if (-not (Test-Path $AccessAnalyzerPath)) {
    Write-Host "ERROR: No se encuentra AccessAnalyzer.accdb" -ForegroundColor Red
    Write-Host "Ruta esperada: $AccessAnalyzerPath" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $ModuleSourcePath)) {
    Write-Host "ERROR: No se encuentra ModExportComplete.bas" -ForegroundColor Red
    Write-Host "Ruta esperada: $ModuleSourcePath" -ForegroundColor Yellow
    exit 1
}

Write-Host "? AccessAnalyzer encontrado: $AccessAnalyzerPath" -ForegroundColor Green
Write-Host "? Módulo fuente encontrado: $ModuleSourcePath" -ForegroundColor Green
Write-Host ""

$access = $null
$vbProject = $null

try {
    Write-Host "1. Abriendo AccessAnalyzer.accdb..." -ForegroundColor Yellow
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AccessAnalyzerPath, $false)
    
    Write-Host "   ? Base de datos abierta" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "2. Preparando inyección del módulo..." -ForegroundColor Yellow
    
    # Obtener el VBProject
    $vbProject = $access.VBE.ActiveVBProject
    
    # Verificar si el módulo ya existe
    $moduleExists = $false
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq "ModExportComplete") {
            $moduleExists = $true
            Write-Host "   ! Módulo ModExportComplete ya existe" -ForegroundColor Yellow
            Write-Host "   ! Se removerá para reimportar la versión más reciente..." -ForegroundColor Yellow
            
            # Remover el módulo existente
            $vbProject.VBComponents.Remove($component)
            Write-Host "   ? Módulo antiguo removido" -ForegroundColor Green
            break
        }
    }
    
    Write-Host ""
    Write-Host "3. Importando módulo ModExportComplete.bas..." -ForegroundColor Yellow
    
    # Importar el nuevo módulo
    $vbProject.VBComponents.Import($ModuleSourcePath)
    
    Write-Host "   ? Módulo importado correctamente" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "4. Validando la función RunCompleteExport..." -ForegroundColor Yellow
    
    # Verificar que la función existe
    $functionFound = $false
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq "ModExportComplete") {
            $codeModule = $component.CodeModule
            $codeText = $codeModule.Lines(1, $codeModule.CountOfLines)
            
            if ($codeText -match "Public Function RunCompleteExport") {
                $functionFound = $true
                Write-Host "   ? Función RunCompleteExport encontrada" -ForegroundColor Green
            }
            break
        }
    }
    
    if (-not $functionFound) {
        Write-Host "   ! Advertencia: Función RunCompleteExport no encontrada" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "5. Guardando cambios..." -ForegroundColor Yellow
    
    $access.CloseCurrentDatabase()
    $access.CurrentDb().Close()
    
    Write-Host "   ? Cambios guardados" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "? ÉXITO: ModExportComplete ha sido inyectado correctamente" -ForegroundColor Green
    Write-Host ""
    Write-Host "Ahora puedes ejecutar access-export-git.ps1 sin problemas" -ForegroundColor Cyan
    Write-Host ""
    
    exit 0
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    exit 1
    
} finally {
    # Limpiar recursos
    if ($access) {
        try {
            $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveAll)
        } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    }
}
