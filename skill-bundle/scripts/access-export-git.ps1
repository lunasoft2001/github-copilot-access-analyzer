# Script para exportar con control de versiones Git

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

# Preguntar sobre exportación de datos de tablas (si no se especificó el parámetro)
if (-not $PSBoundParameters.ContainsKey('ExportTableData')) {
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host " EXPORTACIÓN DE DATOS DE TABLAS" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "¿Deseas exportar los DATOS de las tablas junto con la estructura?" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [S] SÍ  - Exporta estructura + datos (archivos .table + .tabledata)" -ForegroundColor White
    Write-Host "          Útil para: backups completos, migración de datos" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [N] NO  - Exporta solo estructura (archivos .table únicamente)" -ForegroundColor White
    Write-Host "          Útil para: control de versiones Git, solo esquema" -ForegroundColor Gray
    Write-Host ""
    
    do {
        $response = Read-Host "Exportar datos? [S/N]"
        $response = $response.ToUpper()
    } while ($response -ne 'S' -and $response -ne 'N')
    
    $ExportTableData = ($response -eq 'S')
    Write-Host ""
}

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
    
    # Iniciar exportación en background
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
    Write-Host "   Iniciando exportación..." -ForegroundColor Yellow
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
    
    # Cerrar Access (ya no está abierto en el proceso principal)
    $access = $null
    
    if (-not $result) {
        Write-Host "   ERROR en la exportación" -ForegroundColor Red
        exit 1
    }
    
    Write-Host ""
    Write-Host "   ✓ Exportación completada exitosamente" -ForegroundColor Green
    
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
00_RESUMEN_TABLAS.txt
00_LOG_*.txt
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
    Write-Host ""
    Write-Host "Para refactorizar, abre la carpeta en VS Code:" -ForegroundColor Cyan
    Write-Host "  code $ExportFolder" -ForegroundColor White
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    # Limpiar solo si Access sigue abierto en el proceso principal
    # (El job maneja su propio cierre)
    if ($access) {
        try {
            $access.Quit([Microsoft.Office.Interop.Access.AcQuitOption]::acQuitSaveNone)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
        } catch {
            # Ignorar errores al cerrar
        }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
