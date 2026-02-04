# Changelog

All notable changes to this project will be documented in this file.

---

## [Unreleased] - 2026-02-04

### Added
- 🌍 **Soporte multiidioma completo** (ES, EN, DE, FR, IT)
- 📊 **Exportación de tablas individual** con DDL en Access y SQL Server
- 🔧 **Funciones de conversión de tipos** DAO → Access/SQL Server
- 📁 **Carpetas localizadas** según idioma seleccionado

### Changed
- ✨ **Eliminados todos los MsgBox** (12 total) → Reemplazados por Debug.Print
- 🧹 **Limpieza de scripts** 13 → 4 scripts esenciales (-69%)
- 📝 **Consolidación de documentación** 24 → 10 archivos MD (-58%)
- 🔄 **Actualización de módulos VBA** con multiidioma y sin bloqueos

### Removed
- ❌ 9 scripts PowerShell obsoletos
- ❌ 14 archivos .md redundantes consolidados en CHANGELOG.md

### Fixed
- ✅ Automatización sin pausas (sin MsgBox)
- ✅ Compatibilidad con Task Scheduler y CI/CD
- ✅ Estructura de carpetas consistente por idioma

---

## [1.0.0] - 2026-01-29

### Added
- Complete Access database export functionality
- VBA module export with UTF-8 encoding
- Git-based version control workflow
- Selective import based on Git diff
- Auto-generated REFACTORING_PLAN.md
- Comprehensive documentation
- PowerShell automation scripts
- GitHub Copilot skill integration

---

## Detalle de Mejoras - 2026-02-04

### 🌍 Multiidioma
**Idiomas**: Español • English • Deutsch • Français • Italiano

**Archivos modificados**:
- `modules/ModExportComplete.bas` - Función GetFolderName()
- `modules/ModImportComplete.bas` - Función GetFolderName()
- `scripts/access-export-git.ps1` - Parámetro Language
- `scripts/access-import.ps1` - Parámetro Language

**Mapeo de carpetas**:
```
ES: 02_CONSULTAS, 03_FORMULARIOS, 06_CODIGO_VBA
EN: 02_QUERIES, 03_FORMS, 06_VBA_CODE
DE: 02_ABFRAGEN, 03_FORMULARE, 06_VBA_CODE
FR: 02_REQUÊTES, 03_FORMULAIRES, 06_CODE_VBA
IT: 02_QUERY, 03_FORM, 06_VBA_CODE
```

### 📊 Tablas Mejoradas
**Antes**: Un único archivo con todas las tablas  
**Ahora**: DDL individual por tabla en 2 formatos

```
01_Tablas/
├── Access/
│   ├── CLIENTES_DDL.txt
│   └── PEDIDOS_DDL.txt
└── SQLServer/
    ├── CLIENTES_DDL.sql
    └── PEDIDOS_DDL.sql
```

**Nuevas funciones**:
- `ExportTableAccessDDL()` - DDL compatible Access
- `ExportTableSQLServerDDL()` - DDL compatible SQL Server
- `GetAccessFieldType()` - Conversión DAO → Access
- `GetSQLServerFieldType()` - Conversión DAO → SQL Server

### 🚫 Eliminación MsgBox
**Total eliminados**: 12 MsgBox → 0 MsgBox

| Módulo | MsgBox eliminados |
|--------|------------------|
| ModExportComplete.bas | 5 |
| ModImportComplete.bas | 5 |
| ExportTodoSimple.bas | 2 |

**Beneficios**:
- ✅ Ejecución sin pausas
- ✅ Compatible con Task Scheduler
- ✅ Compatible con CI/CD
- ✅ Logs en VBA Immediate Window

### 🧹 Limpieza Scripts
**Eliminados (9)**:
- access-export.ps1
- access-export-complete.ps1
- access-export-simple.ps1
- access-export-tool.ps1
- access-import-old.ps1
- test-export.ps1
- test-import.ps1
- import-module-and-test.ps1
- check-modules.ps1

**Mantenidos (4)**:
- ✅ access-backup.ps1
- ✅ access-export-git.ps1 ⭐
- ✅ access-import.ps1 ⭐
- ✅ access-import-changed.ps1

### 📚 Scripts Actuales

#### access-export-git.ps1 ⭐ PRINCIPAL
```powershell
.\access-export-git.ps1 -DatabasePath "DB.accdb" -Language "ES"
```
- Export completo con Git
- Genera REFACTORING_PLAN.md
- Carpetas localizadas
- DDL individual por tabla

#### access-import.ps1 ⭐ PRINCIPAL
```powershell
.\access-import.ps1 -TargetDbPath "DB.accdb" -ImportFolder "export" -Language "ES"
```
- Import completo
- Backup automático
- Sin MsgBox

#### access-import-changed.ps1
```powershell
.\access-import-changed.ps1 -TargetDbPath "DB.accdb" -ExportFolder "export" -DryRun
```
- Import inteligente (solo cambios Git)
- Dry run disponible

#### access-backup.ps1
```powershell
.\access-backup.ps1 -DatabasePath "DB.accdb"
```
- Backup timestamped

---

## 📊 Estadísticas

| Métrica | v1.0.0 | Unreleased | Cambio |
|---------|--------|------------|--------|
| Scripts PS | 13 | 4 | -69% |
| MsgBox | 12 | 0 | -100% |
| Idiomas | 1 | 5 | +400% |
| Archivos MD | 24 | 10 | -58% |

---

## 🔄 Migración Requerida

**⚠️ AccessAnalyzer.accdb debe actualizarse**

1. Abrir AccessAnalyzer.accdb
2. Alt+F11 (VBA Editor)
3. Eliminar módulos antiguos
4. Importar desde `modules/`:
   - ModExportComplete.bas
   - ModImportComplete.bas
5. Ctrl+S (Guardar)

---

## 🐛 Problemas Conocidos

**access-import-changed.ps1**:
- Usa paths hardcoded (solo español)
- Workaround: Usar access-import.ps1

---

## 📖 Referencias

- `SCRIPTS_REFERENCIA.md` - Guía completa de scripts
- `SETUP.md` - Instalación
- `README_GIT_WORKFLOW.md` - Workflow Git
- `SKILL.md` - Definición del skill

---

[1.0.0]: https://github.com/lunasoft2001/github-copilot-access-analyzer/releases/tag/v1.0.0