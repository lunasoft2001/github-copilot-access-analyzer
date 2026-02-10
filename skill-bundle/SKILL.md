---
name: access-analyzer
description: "Analyze, export, refactor, and re-import Microsoft Access database applications. Use when working with .accdb/.mdb files to: (1) Create backups, (2) Export all objects (tables, queries, forms, reports, macros, VBA) to text files for version control and analysis, (3) Optionally export table data or structure-only, (4) Refactor exported code in VS Code, (5) Import changes back into Access. Includes automated export module based on ExportTodoSimple.bas."
license: MIT
---

# Microsoft Access Analyzer & Refactoring Tool

Comprehensive workflow for analyzing, exporting, refactoring, and re-importing Microsoft Access database applications without manually adding export modules to each project.

## When to Use This Skill

Use this skill when the user asks to:
- Analyze an Access database (.accdb or .mdb file)
- Export Access objects for version control or code review
- Refactor Access VBA code in a modern editor (VS Code)
- Apply changes back to the original Access database
- Create backups before making changes
- Understand the structure of an Access application

## Core Workflow

### Phase 1: Backup
Always create a timestamped backup before any operation:
```
DatabaseName_BACKUP_yyyyMMdd_HHmmss.accdb
```

### Phase 2: Export All Objects
Export the complete Access application to a structured folder:

```
Exportacion_yyyyMMdd_HHmmss/
├── 00_RESUMEN_APLICACION.txt    # Overview and inventory
├── 01_Tablas/                    # Table structures (DDL)
│   ├── Access/                   # Access DDL per table
│   │   ├── CLIENTES.txt
│   │   ├── PEDIDOS.txt
│   │   └── ...
│   └── SQLServer/                # SQL Server DDL per table
│       ├── CLIENTES.txt
│       ├── PEDIDOS.txt
│       └── ...
├── 02_Consultas/                 # Queries as .sql files
├── 03_Formularios/               # Forms + VBA code
├── 04_Informes/                  # Reports + VBA code
├── 05_Macros/                    # Macro definitions
└── 06_Codigo_VBA/                # All VBA modules (.bas/.cls)
```

**Key Features:**
- UTF-8 encoding for perfect VS Code compatibility
- Spanish characters (á, é, í, ó, ú, ñ) preserved correctly
- **Each table exported as individual DDL file** (Access AND SQL Server formats)- **Optional table data export** (estructura sola o estructura + datos)- Individual .sql files for queries
- Complete VBA module extraction
- Table structure documentation with types and properties

### Phase 3: Analysis & Refactoring
Once exported, analyze and refactor in VS Code:
- Review code quality and patterns
- Identify dependencies and relationships
- Apply refactoring patterns
- Add comments and documentation
- Fix bugs or improve logic
- Modernize code structure

### Phase 4: Re-import (When User Requests)
Import modified code back into Access:
1. Validate exported files haven't been corrupted
2. Create new backup before import
3. Import VBA modules
4. Import query definitions
5. Import form/report code
6. Verify all imports successful

## How to Execute

### Step 1: Locate Access File
Ask user for the path to the .accdb or .mdb file, or search workspace.

### Step 2: Create Backup
```powershell
# Example
Copy-Item "C:\path\to\database.accdb" "C:\path\to\database_BACKUP_20260128_143000.accdb"
```

### Step 3: Export Using VBA Module
The skill includes [ExportTodoSimple.bas](./references/ExportTodoSimple.bas) module that handles the complete export process.

**Two approaches:**

**A) Temporary Injection (Recommended)**
1. Open Access database programmatically
2. Temporarily import ExportTodoSimple module
3. Execute `ExportAll` subroutine
4. Remove temporary module
5. Close database

**B) PowerShell Automation with Git (RECOMENDADO)**
Use the [access-export-git.ps1](./scripts/access-export-git.ps1) script which:
- Opens AccessAnalyzer.accdb tool database
- Calls `RunCompleteExport` from ModExportComplete
- Exports to persistent folder: `{DatabaseName}_Export`
- Initializes Git repository in export folder (if not exists)
- Creates .gitignore (excludes .ldb, backups, errors)
- **Crea automáticamente `REFACTORING_PLAN.md`** con checklist y tracking
- Commits all changes with timestamp
- Shows git diff stats
- **Pregunta si quiere abrir VS Code automáticamente**

**Ventajas del workflow Git:**
- ✅ Historial completo de cambios
- ✅ Detección automática de archivos modificados
- ✅ Rollback a versiones anteriores
- ✅ Solo importa lo que cambió (más rápido y seguro)
- ✅ Plan de refactorización incluido

### Step 4: Open in VS Code
Al exportar con `access-export-git.ps1`, se pregunta automáticamente:
```
¿Abrir en VS Code para refactorizar? (S/N)
```

O manualmente:
```powershell
code "path\to\Exportacion_folder"
```

### Step 5: Analyze & Refactor
- **Start with `REFACTORING_PLAN.md`** - Plan generado automáticamente con:
  - Inventario de objetos exportados
  - Checklist de refactorización por fases
  - Template para documentar cambios
  - Notas para código problemático
- Review `00_RESUMEN_APLICACION.txt` for database overview
- Explore `06_Codigo_VBA/` for main business logic
- Review queries in `02_Consultas/`
- Check forms/reports for UI logic

**Usando control de versiones Git:**
1. Modifica los archivos .bas o .txt que necesites refactorizar
2. Guarda cambios en VS Code
3. Git detectará automáticamente los archivos modificados

### Step 6: Re-import Inteligente (Solo Cambios)
Use [access-import-changed.ps1](./scripts/access-import-changed.ps1) para importar **solo lo modificado**:

```powershell
# Modo dry-run (ver qué se importaría)
.\access-import-changed.ps1 -TargetDbPath "C:\path\to\db.accdb" `
                             -ExportFolder "C:\path\to\db_Export" `
                             -DryRun

# Importación real
.\access-import-changed.ps1 -TargetDbPath "C:\path\to\db.accdb" `
                             -ExportFolder "C:\path\to\db_Export"
```

**Cómo funciona:**
1. Detecta archivos modificados con `git diff --name-only HEAD~1 HEAD`
2. Clasifica por tipo: Consultas, Formularios, Informes, Macros, Módulos VBA
3. Crea backup automático antes de importar
4. Importa SOLO los objetos que cambiaron
5. Muestra progreso con OK/ERROR por cada objeto
6. Genera reporte final con count de importados y errores
7. **Pregunta si quiere abrir Access automáticamente** para verificar cambios

**Ventajas:**
- ✅ Solo importa lo modificado (evita errores en objetos que funcionan)
- ✅ Mucho más rápido (3 objetos vs 621 totales)
- ✅ Menor riesgo de romper objetos que no tocaste
- ✅ Backup automático por seguridad
- ✅ Apertura automática de Access para validar

## REFACTORING_PLAN.md - Plan Automático

Al exportar con `access-export-git.ps1`, se genera automáticamente un **Plan de Refactorización** que incluye:

### Contenido del Plan
1. **📊 Inventario** - Conteo automático de todos los objetos exportados
2. **🎯 Objetivos** - Checklist editable de metas de refactorización
3. **✅ Checklist por Fases:**
   - Análisis (revisar resumen, identificar módulos principales)
   - Planificación (priorizar, definir estándares)
   - Ejecución (refactorizar módulos, optimizar queries)
   - Validación (dry-run, importar, probar)
4. **📝 Registro de Cambios** - Template para documentar cada modificación
5. **🔝 Notas de Refactorización:**
   - Código problemático encontrado
   - Dependencias identificadas
   - Queries que necesitan optimización
6. **🚀 Comandos Git** - Referencia rápida de comandos útiles
7. **📌 Próximos Pasos** - Guía de workflow

### Uso del Plan
- Editar objetivos según necesidades
- Marcar checkboxes conforme avanza (`- [x]`)
- Documentar cada cambio en registro
- Añadir notas sobre problemas encontrados
- Versionar con Git (historial de planificación)

## Important Notes

### Cu�ndo Exportar Datos vs Solo Estructura

**Exportar SOLO estructura** (`-ExportTableData:$false`):
- ? Control de versiones Git (archivos m�s peque�os)
- ? An�lisis de esquema de base de datos
- ? Refactorizaci�n de c�digo VBA y queries
- ? Documentaci�n de estructura
- ? Comparaci�n de versiones de esquema
- ?? Los archivos `.tabledata` no se versionan bien en Git (grandes, binarios)

**Exportar estructura + datos** (`-ExportTableData:$true`):
- ? Backup completo antes de migraciones
- ? Transferir datos entre entornos
- ? Archivar estado completo de la aplicaci�n
- ? An�lisis de datos reales
- ? Testing con datos de producci�n
- ?? Archivos m�s grandes, commit Git pesado

**Recomendaci�n general:** Para workflows de refactorizaci�n con Git, usa `-ExportTableData:$false`. Solo incluye datos para backups espec�ficos fuera de Git.

### UTF-8 Encoding
All exports use UTF-8 encoding. VS Code will display Spanish characters perfectly:
- ✅ Acentos: á, é, í, ó, ú
- ✅ Ñ and other special characters
- ✅ No manual encoding changes needed

### VBA Access Requirements
Access must allow programmatic access to VBA project:
- File → Options → Trust Center → Trust Center Settings
- Enable "Trust access to the VBA project object model"

### Limitations
- Cannot export binary objects (images, OLE objects)
- Form/report visual layout exported as text (not editable visually)
- Macros exported as definitions (not executable outside Access)
- Linked tables show connection info only

### Safety
- Always creates backup before operations
- Non-destructive exports (original database unchanged)
- Import creates new backup before applying changes
- Logs all operations for traceability

## Scripts Reference

### [access-export-git.ps1](./scripts/access-export-git.ps1) ⭝ RECOMENDADO
Exportación automatizada con control de versiones Git integrado.

**Usage:**
```powershell
# Pregunta interactivamente sobre exportar datos
.\access-export-git.ps1 -DatabasePath "C:\export\test\appGraz.accdb"

# Solo estructura (sin datos) - ideal para Git
.\access-export-git.ps1 -DatabasePath "C:\export\test\appGraz.accdb" -ExportTableData:$false

# Estructura + datos - backup completo
.\access-export-git.ps1 -DatabasePath "C:\export\test\appGraz.accdb" -ExportTableData:$true

# Con idioma espec�fico
.\access-export-git.ps1 -DatabasePath "C:\export\test\appGraz.accdb" -Language "EN"
```

**Par�metros:**
- `-DatabasePath`: Ruta al archivo .accdb/.mdb (obligatorio)
- `-ExportTableData`: `$true` = estructura + datos, `$false` = solo estructura (pregunta si se omite)
- `-Language`: Idioma de exportaci�n (ES/EN/DE/FR/IT, default: ES)
- `-ExportFolder`: Carpeta destino (default: `{DatabaseName}_Export`)

**Caracter�sticas:**
- **Pregunta interactivamente** si exportar datos de tablas (si no se especifica par�metro)
- Exporta a carpeta persistente: `{DatabaseName}_Export`
- Inicializa repositorio Git automáticamente
- Crea .gitignore para archivos temporales (.ldb, backups, errors)
- Commit automático con timestamp
- Muestra estadísticas de archivos exportados

**Returns:** Carpeta de exportación con Git inicializado

### [access-import-changed.ps1](./scripts/access-import-changed.ps1) ⭝ RECOMENDADO
Importación inteligente que **solo importa archivos modificados** detectados por Git.

**Notas recientes (importacion robusta):**
- Soporta consultas en `.sql` (crea QueryDef directamente).
- Elimina comentarios de linea `--` antes de crear la consulta.
- Permite seleccion manual con `-Interactive` o por nombre (`-QueryNames`, `-ModuleNames`, `-FormNames`, `-ReportNames`, `-MacroNames`).
- Abre Access **despues** de liberar COM para evitar ventana en blanco o bloqueo.
- Para modulos/clases usa VBE (requiere "Trust access to the VBA project object model").

**Usage:**
```powershell
# Dry-run (ver qué se importaría)
.\access-import-changed.ps1 -TargetDbPath "C:\path\to\db.accdb" `
                             -ExportFolder "C:\path\to\db_Export" `
                             -DryRun

# Importación real
.\access-import-changed.ps1 -TargetDbPath "C:\path\to\db.accdb" `
                             -ExportFolder "C:\path\to\db_Export"
```

**Características:**
- Detecta cambios con `git diff HEAD~1 HEAD`
- Solo importa objetos modificados (no todo)
- Backup automático antes de importar
- Progreso detallado por objeto (OK/ERROR)
- Reporte final de importados/errores

**Returns:** Count de objetos importados y errores

### [access-import.ps1](./scripts/access-import.ps1)
Re-import all exported objects back to Access (importación completa).

**Usage:**
```powershell
.\access-import.ps1 -TargetDbPath "C:\path\to\db.accdb" -ImportFolder "C:\export"

# Con idioma específico
.\access-import.ps1 -TargetDbPath "C:\path\to\db.accdb" -ImportFolder "C:\export" -Language "EN"
```

**Características:**
- Backup automático antes de importar
- Compatible con carpetas multiidioma
- Sin interrupciones (sin MsgBox)
- Importa completo: queries, forms, reports, macros, VBA

**Returns:** Import log with success/failure details

### [access-backup.ps1](./scripts/access-backup.ps1)
Create timestamped backup of Access file.

**Usage:**
```powershell
.\access-backup.ps1 -DatabasePath "C:\path\to\db.accdb"
```

**Returns:** Path to backup file

## Common Patterns

### Pattern 1: Workflow Completo con Git (RECOMENDADO)
```powershell
# 1. Primera exportaci�n con Git (solo estructura - ideal para control de versiones)
cd C:\Users\juanjo_admin\.copilot\skills\access-analyzer\scripts
.\access-export-git.ps1 -DatabasePath "C:\export\test\appGraz.accdb" -ExportTableData:$false

# 2. Abrir en VS Code
cd C:\export\test\appGraz_Export
code .

# 3. Refactorizar archivos en VS Code
# - Edita módulos .bas en 06_Codigo_VBA/
# - Modifica consultas .txt en 02_Consultas/
# - Guarda cambios

# 4. Ver cambios detectados por Git
git status
git diff

# 5. Crear commit con cambios
git add -A
git commit -m "Refactorización: mejorar lógica de X"

# 6. Ver qué se importaría (dry-run)
cd C:\Users\juanjo_admin\.copilot\skills\access-analyzer\scripts
.\access-import-changed.ps1 -TargetDbPath "C:\export\test\appGraz.accdb" `
                             -ExportFolder "C:\export\test\appGraz_Export" `
                             -DryRun

# 7. Importar SOLO lo modificado
.\access-import-changed.ps1 -TargetDbPath "C:\export\test\appGraz.accdb" `
                             -ExportFolder "C:\export\test\appGraz_Export"

# 8. Si algo salió mal, rollback con Git
cd C:\export\test\appGraz_Export
git log --oneline              # Ver historial
git revert HEAD                # Deshacer último commit
# O restaurar backup: appGraz_BACKUP_timestamp.accdb
```

### Pattern 2: Quick Analysis (Sin Git)
```powershell
User: "Analyze this Access database: C:\project\inventory.accdb"

1. Create backup
2. Export all objects (pregunta si incluir datos)
3. Open in VS Code
4. Show 00_RESUMEN_APLICACION.txt
5. Provide high-level overview
```

### Pattern 3: Backup Completo con Datos
```powershell
User: "Create a complete backup with data for migration"

1. Create timestamped backup of .accdb file
2. Export with -ExportTableData:$true
3. Archive export folder for migration
4. Document table structures and data formats
```

### Pattern 4: Refactor Specific Module
```
User: "Refactor the invoice calculation module"

1. Export if not already done
2. Locate module in 06_Codigo_VBA/
3. Analyze code
4. Suggest improvements
5. Apply refactoring when approved
6. Re-import when user confirms
```

### Pattern 5: Version Control Setup
```
User: "Set up version control for this Access app"

1. Export complete structure
2. Initialize git repository in export folder
3. Create .gitignore for Access-specific files
4. Create initial commit
5. Document workflow for team
```

### Pattern 6: Database Documentation
```
User: "Document this Access database"

1. Export all objects
2. Analyze table relationships
3. Map query dependencies
4. Document form workflows
5. Generate comprehensive markdown documentation
```

## Reference Files

- [ExportTodoSimple.bas](./references/ExportTodoSimple.bas) - Complete VBA export module (original)
- [AccessObjectTypes.md](./references/AccessObjectTypes.md) - Reference for Access object types and properties
- [VBA-Patterns.md](./references/VBA-Patterns.md) - Common VBA patterns and refactoring guidelines

## Troubleshooting

### "Cannot access VBA project"
Solution: Enable "Trust access to the VBA project object model" in Access Trust Center settings.

### "Export creates garbled characters"
Solution: Verify UTF-8 encoding in exported files. Re-export if necessary.

### "Import fails for specific objects"
Solution: Check import log for specific errors. Some objects may need manual intervention.

### "Database is locked/in use"
Solution: Close all Access instances and try again. Check for .laccdb lock files.

## Security Considerations

- Never commit database files (.accdb) to version control
- Backup before any destructive operation
- Review imported code before executing in production
- Keep backups in separate location from source
- Use source control for exported text files only
