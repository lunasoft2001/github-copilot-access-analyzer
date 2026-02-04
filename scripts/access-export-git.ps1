# Script para exportar con control de versiones Git

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [string]$ExportFolder = "",
    
    [ValidateSet("ES", "EN", "DE", "FR", "IT")]
    [string]$Language = "ES",
    
    [string]$AnalyzerPath = "$PSScriptRoot\..\AccessAnalyzer.accdb"
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
    Write-Host "1. Abriendo AccessAnalyzer..." -ForegroundColor Yellow
    
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AnalyzerPath, $false)
    
    Write-Host "   OK" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "2. Exportando base de datos..." -ForegroundColor Yellow
    Write-Host "   Base: $DatabasePath" -ForegroundColor Cyan
    Write-Host "   Carpeta: $ExportFolder" -ForegroundColor Cyan
    Write-Host "   Idioma: $Language" -ForegroundColor Cyan
    
    # Construir comando
    $dbEscaped = $DatabasePath.Replace('\', '\\')
    $outEscaped = $ExportFolder.Replace('\', '\\')
    $cmd = 'RunCompleteExport("' + $dbEscaped + '","' + $outEscaped + '","' + $Language + '")'
    
    $result = $access.Eval($cmd)
    
    if (-not $result) {
        Write-Host "   ERROR en la exportación" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "   OK - Exportación completada" -ForegroundColor Green
    
    # Inicializar Git si no existe
    Write-Host ""
    Write-Host "3. Configurando control de versiones..." -ForegroundColor Yellow
    
    Push-Location $ExportFolder
    
    if (-not (Test-Path ".git")) {
        Write-Host "   Inicializando repositorio Git..." -ForegroundColor Yellow
        git init | Out-Null
        
        # Crear .gitignore
        @"
# Access temporary files
*.ldb
*.laccdb

# Backup files
*_BACKUP_*.accdb

# Error files
errors*.txt
ERROR_*.txt

# Resumen (no versionamos por timestamp)
00_RESUMEN.txt
"@ | Set-Content ".gitignore" -Encoding UTF8
        
        git add .gitignore | Out-Null
        git commit -m "Initial commit: .gitignore" | Out-Null
        
        Write-Host "   OK - Repositorio Git creado" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "4. Registrando cambios en Git..." -ForegroundColor Yellow
    
    # Agregar todos los cambios
    git add -A | Out-Null
    
    # Ver si hay cambios
    $status = git status --porcelain
    
    if ($status) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $commitMsg = "Export from $DatabasePath at $timestamp"
        
        git commit -m $commitMsg | Out-Null
        
        Write-Host "   OK - Commit creado" -ForegroundColor Green
        Write-Host "   Mensaje: $commitMsg" -ForegroundColor Gray
        
        # Mostrar estadísticas
        Write-Host ""
        Write-Host "   Archivos modificados:" -ForegroundColor Cyan
        git diff --stat HEAD~1 HEAD | ForEach-Object { Write-Host "     $_" -ForegroundColor White }
    }
    else {
        Write-Host "   No hay cambios desde la última exportación" -ForegroundColor Yellow
    }
    
    Pop-Location
    
    # Crear plan de refactorización
    Write-Host ""
    Write-Host "5. Creando plan de refactorización..." -ForegroundColor Yellow
    
    Push-Location $ExportFolder
    
    # Contar objetos exportados
    $queryCount = (Get-ChildItem "02_Consultas\*.txt" -ErrorAction SilentlyContinue).Count
    $formCount = (Get-ChildItem "03_Formularios\*.txt" -ErrorAction SilentlyContinue).Count
    $reportCount = (Get-ChildItem "04_Informes\*.txt" -ErrorAction SilentlyContinue).Count
    $macroCount = (Get-ChildItem "05_Macros\*.txt" -ErrorAction SilentlyContinue).Count
    $moduleCount = (Get-ChildItem "06_Codigo_VBA\*.bas" -ErrorAction SilentlyContinue).Count
    
    $refactoringPlan = @"
# 📋 Plan de Refactorización
**Base de datos:** ``$($DatabasePath | Split-Path -Leaf)``  
**Fecha de exportación:** $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")  
**Carpeta:** ``$ExportFolder``

---

## 📊 Inventario de Objetos Exportados

| Tipo | Cantidad | Ubicación |
|------|----------|-----------|
| Consultas | $queryCount | ``02_Consultas/`` |
| Formularios | $formCount | ``03_Formularios/`` |
| Informes | $reportCount | ``04_Informes/`` |
| Macros | $macroCount | ``05_Macros/`` |
| Módulos VBA | $moduleCount | ``06_Codigo_VBA/`` |

**Total objetos:** $($queryCount + $formCount + $reportCount + $macroCount + $moduleCount)

---

## 🎯 Objetivos de Refactorización

<!-- Describe qué quieres lograr con esta refactorización -->

- [ ] **Objetivo 1:** Mejorar rendimiento de consultas
- [ ] **Objetivo 2:** Simplificar lógica de formularios
- [ ] **Objetivo 3:** Refactorizar módulos VBA duplicados
- [ ] **Objetivo 4:** Documentar código sin comentarios
- [ ] **Objetivo 5:** Eliminar código muerto

---

## ✅ Checklist de Refactorización

### Fase 1: Análisis
- [ ] Revisar ``00_RESUMEN_APLICACION.txt``
- [ ] Identificar módulos principales en ``06_Codigo_VBA/``
- [ ] Listar consultas más complejas en ``02_Consultas/``
- [ ] Encontrar formularios con mucho código
- [ ] Buscar código duplicado

### Fase 2: Planificación
- [ ] Priorizar archivos a refactorizar
- [ ] Definir estándares de código
- [ ] Planificar tests de regresión
- [ ] Crear backup adicional si es crítico

### Fase 3: Ejecución
- [ ] Refactorizar módulos VBA
- [ ] Optimizar consultas SQL
- [ ] Simplificar formularios
- [ ] Mejorar nombres de variables
- [ ] Agregar comentarios

### Fase 4: Validación
- [ ] Dry-run de importación
- [ ] Importar cambios
- [ ] Probar funcionalidad en Access
- [ ] Verificar que no se rompió nada
- [ ] Documentar cambios realizados

---

## 📝 Registro de Cambios

<!-- Documenta aquí cada cambio que hagas -->

### Cambio 1
**Fecha:** $(Get-Date -Format "dd/MM/yyyy")  
**Archivos modificados:**  
- ``06_Codigo_VBA/Modul_General.bas``

**Descripción:**  
<!-- Qué cambiaste y por qué -->

**Resultado:**  
<!-- ✅ OK | ❌ Error | ⚠️ Revisar -->

---

### Cambio 2
**Fecha:** _pendiente_  
**Archivos modificados:**  
- 

**Descripción:**  


**Resultado:**  


---

## 🔍 Notas de Refactorización

### Código Problemático Encontrado
<!-- Lista aquí código que necesita atención especial -->

- **Archivo:** ``06_Codigo_VBA/basExcel.bas``  
  **Problema:** Función con 500+ líneas, difícil de mantener  
  **Solución propuesta:** Dividir en funciones más pequeñas

### Dependencias Identificadas
<!-- Documenta relaciones entre módulos -->

- ``Modul_General.bas`` depende de ``basExcel.bas``
- ``frmKUNDEN`` usa funciones de ``Modul_Funciones_globales.bas``

### Queries que Necesitan Optimización
<!-- Lista consultas lentas o complejas -->

1. ``abKUNDEN`` - JOIN múltiple, optimizar índices
2. ``abRECHNUNG_TOTAL`` - Subconsultas anidadas

---

## 🚀 Comandos Git Útiles

\`\`\`powershell
# Ver qué modificaste
git status

# Ver cambios en detalle
git diff

# Confirmar cambios
git add -A
git commit -m "Descripción de cambios"

# Ver historial
git log --oneline

# Deshacer si algo sale mal
git reset --hard HEAD
\`\`\`

---

## 📌 Próximos Pasos

1. Revisar este plan y ajustar objetivos
2. Empezar por archivos más críticos
3. Hacer commits frecuentes (después de cada cambio funcional)
4. Probar con ``access-import-changed.ps1 -DryRun`` antes de importar
5. Importar y validar en Access

---

**Generado automáticamente por:** ``access-export-git.ps1``  
**Última actualización:** $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
"@
    
    $refactoringPlan | Set-Content "REFACTORING_PLAN.md" -Encoding UTF8
    Write-Host "   OK - Plan creado: REFACTORING_PLAN.md" -ForegroundColor Green
    
    Pop-Location
    
    Write-Host ""
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "EXPORTACION EXITOSA CON VERSION CONTROL" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Carpeta: $ExportFolder" -ForegroundColor White
    Write-Host ""
    Write-Host "Comandos útiles:" -ForegroundColor Cyan
    Write-Host "  cd $ExportFolder" -ForegroundColor White
    Write-Host "  git log --oneline     # Ver historial" -ForegroundColor White
    Write-Host "  git diff HEAD~1       # Ver cambios" -ForegroundColor White
    Write-Host "  git status            # Ver estado" -ForegroundColor White
    
    # Preguntar si quiere abrir en VS Code
    Write-Host ""
    $openVSCode = Read-Host "¿Abrir en VS Code para refactorizar? (S/N)"
    
    if ($openVSCode -eq 'S' -or $openVSCode -eq 's' -or $openVSCode -eq 'Y' -or $openVSCode -eq 'y') {
        Write-Host "Abriendo VS Code..." -ForegroundColor Yellow
        code "$ExportFolder"
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
