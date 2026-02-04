# 📋 Scripts PowerShell - Referencia Rápida

## Estado Actual: ✅ Limpio y Optimizado

Total de scripts: **4 esenciales** (eliminados 9 obsoletos)

---

## Scripts Disponibles

### 1. 🔐 `access-backup.ps1`
**Propósito**: Crear copia de seguridad automática

```powershell
.\access-backup.ps1 -DatabasePath "C:\MiBD.accdb"
```

**Genera**: `MiBD_BACKUP_<timestamp>.accdb`

---

### 2. 📤 `access-export-git.ps1` ⭐ **PRINCIPAL**
**Propósito**: Exportación completa con Git + Refactoring Plan

```powershell
# Español (default)
.\access-export-git.ps1 -DatabasePath "C:\MiBD.accdb"

# Otros idiomas
.\access-export-git.ps1 -DatabasePath "C:\MiBD.accdb" -Language "EN"
.\access-export-git.ps1 -DatabasePath "C:\MiBD.accdb" -Language "DE"
```

**Características**:
- ✅ Export multiidioma (ES/EN/DE/FR/IT)
- ✅ Git integration automático
- ✅ Genera REFACTORING_PLAN.md
- ✅ Carpetas localizadas
- ✅ Tablas en DDL individual (Access + SQL Server)

**Genera**:
```
export/
├── .git/
├── REFACTORING_PLAN.md
├── 00_RESUMEN.txt
├── 01_TABLAS/
│   ├── Access/
│   └── SQLServer/
├── 02_CONSULTAS/ (o QUERIES, ABFRAGEN, etc)
└── ... más carpetas
```

---

### 3. 📥 `access-import.ps1` ⭐ **PRINCIPAL**
**Propósito**: Importación completa desde export

```powershell
# Español (default)
.\access-import.ps1 -TargetDbPath "C:\MiBD.accdb" -ImportFolder "export"

# Con idioma específico
.\access-import.ps1 -TargetDbPath "C:\MiBD.accdb" -ImportFolder "export" -Language "EN"
```

**Características**:
- ✅ Import multiidioma (detecta automáticamente)
- ✅ Crea backup antes de importar
- ✅ Reimporta todos los objetos
- ✅ Sin interrupciones (sin MsgBox)

**Genera**:
```
MiBD_BACKUP_BEFORE_IMPORT_<timestamp>.accdb
```

---

### 4. 🎯 `access-import-changed.ps1` ⭐ **ALTERNATIVA**
**Propósito**: Importación inteligente (solo cambios detectados por Git)

```powershell
# Detectar y importar solo lo modificado
.\access-import-changed.ps1 -TargetDbPath "C:\MiBD.accdb" -ExportFolder "export"

# Modo "dry run" para ver qué se importaría
.\access-import-changed.ps1 -TargetDbPath "C:\MiBD.accdb" -ExportFolder "export" -DryRun
```

**Características**:
- ✅ Lee Git diff (HEAD~1 HEAD)
- ✅ Importa solo cambios recientes
- ✅ Más rápido para grandes bases de datos
- ✅ Dry run para previsualizar

**Nota**: Requiere que export esté en repositorio Git

---

## 🗑️ Scripts Eliminados

| Script | Razón |
|--------|-------|
| access-export.ps1 | Usa ExportTodoSimple (antiguo) |
| access-export-complete.ps1 | Versión antigua, superada por access-export-git.ps1 |
| access-export-simple.ps1 | Versión simple, no mantiene Git |
| access-export-tool.ps1 | Herramienta alternativa/experimental |
| access-import-old.ps1 | Claramente obsoleto |
| test-export.ps1 | Scripts de testing internos |
| test-import.ps1 | Scripts de testing internos |
| import-module-and-test.ps1 | Script de desarrollo |
| check-modules.ps1 | Utilidad de debugging |

---

## 📊 Flujo de Trabajo Recomendado

### Scenario 1: Exportar base de datos
```powershell
cd e:\datos\GitHub\github-copilot-access-analyzer\scripts
.\access-export-git.ps1 -DatabasePath "C:\MiBD.accdb"
```
✅ Crea carpeta con todo exportado
✅ Git initialized automáticamente
✅ REFACTORING_PLAN.md generado

### Scenario 2: Hacer cambios en VS Code
```
1. Edita archivos en export/
2. Git commit tus cambios
3. cd e:\datos\GitHub\github-copilot-access-analyzer\scripts
```

### Scenario 3: Reimportar cambios a BD
```powershell
.\access-import.ps1 -TargetDbPath "C:\MiBD.accdb" -ImportFolder "export"
```
✅ Crea backup automático
✅ Reimporta todo limpiamente
✅ Sin pausas de MsgBox

### Scenario 4: Importar solo cambios recientes
```powershell
.\access-import-changed.ps1 -TargetDbPath "C:\MiBD.accdb" -ExportFolder "export" -DryRun
.\access-import-changed.ps1 -TargetDbPath "C:\MiBD.accdb" -ExportFolder "export"
```
✅ Más rápido
✅ Solo cambios desde último commit

---

## ⚙️ Parámetros Comunes

### Todos los scripts aceptan:

| Parámetro | Valores | Defecto |
|-----------|---------|---------|
| `-DatabasePath` | Ruta completa a .accdb | Requerido |
| `-Language` | ES, EN, DE, FR, IT | ES |
| `-OutputFolder` | Ruta de salida | Auto (con timestamp) |

### Ejemplos de Lenguajes:

```powershell
# Español
.\access-export-git.ps1 -DatabasePath "app.accdb" -Language "ES"
# Genera: 02_CONSULTAS, 03_FORMULARIOS, 06_CODIGO_VBA

# Inglés
.\access-export-git.ps1 -DatabasePath "app.accdb" -Language "EN"
# Genera: 02_QUERIES, 03_FORMS, 06_VBA_CODE

# Alemán
.\access-export-git.ps1 -DatabasePath "app.accdb" -Language "DE"
# Genera: 02_ABFRAGEN, 03_FORMULARE, 06_VBA_CODE

# Francés
.\access-export-git.ps1 -DatabasePath "app.accdb" -Language "FR"
# Genera: 02_REQUÊTES, 03_FORMULAIRES, 06_CODE_VBA

# Italiano
.\access-export-git.ps1 -DatabasePath "app.accdb" -Language "IT"
# Genera: 02_QUERY, 03_FORM, 06_VBA_CODE
```

---

## 🔍 Troubleshooting

### ❌ "No se encuentra AccessAnalyzer.accdb"
```powershell
# Solución: Copia AccessAnalyzer.accdb a la raíz del proyecto
# Debe estar en: e:\datos\GitHub\github-copilot-access-analyzer\AccessAnalyzer.accdb
```

### ❌ "Error en ModExportComplete"
```powershell
# Solución: Actualiza AccessAnalyzer.accdb con nuevos módulos
# Ver: ACTUALIZAR_ACCESSANALYZER.md
```

### ❌ "La carpeta de importación no se encuentra"
```powershell
# Solución: Verifica la ruta
.\access-import.ps1 -TargetDbPath "C:\BD.accdb" -ImportFolder "C:\ruta\export"
```

---

## 📌 Notas Importantes

✅ **Todos los scripts creados recientemente** NO tienen MsgBox (sin bloqueos)

✅ **Multiidioma completo** - carpetas se crean en idioma seleccionado

✅ **Backups automáticos** - import siempre crea backup antes

✅ **Git integration** - export-git.ps1 maneja versión control automáticamente

✅ **Compatible PowerShell** - tested en Windows PowerShell 5.1+

---

## 🚀 Próximas Mejoras (Futuro)

- Logging a archivo en lugar de solo Debug.Print
- Soporte para logging centralizado
- Integración con Azure DevOps o GitHub Actions
- CLI mejorada con más opciones

