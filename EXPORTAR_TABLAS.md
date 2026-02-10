# Exportación de Tablas

## Opciones de Exportación XML

El sistema ahora permite controlar si se exportan los **datos** de las tablas o solo la **estructura**.

### Exportar CON datos (por defecto)

```powershell
.\access-export-git-FIXED.ps1 -DatabasePath "C:\ruta\database.accdb" -ExportTableData
```

Genera:
- `01_Tablas/XML/NombreTabla.table` (estructura)
- `01_Tablas/XML/NombreTabla.tabledata` (datos)
- `01_Tablas/Access/NombreTabla.txt` (DDL Access)
- `01_Tablas/SQLServer/NombreTabla.txt` (DDL SQL Server)

### Exportar SIN datos (solo estructura)

```powershell
.\access-export-git-FIXED.ps1 -DatabasePath "C:\ruta\database.accdb"
```

Genera:
- `01_Tablas/XML/NombreTabla.table` (solo estructura)
- `01_Tablas/Access/NombreTabla.txt` (DDL Access)
- `01_Tablas/SQLServer/NombreTabla.txt` (DDL SQL Server)

?? **Sin el switch `-ExportTableData`, NO se exportan datos**

## Casos de Uso

### 1. Control de versiones con Git

**Recomendación:** NO exportar datos

```powershell
.\access-export-git-FIXED.ps1 -DatabasePath "E:\datos\GitHub\test260210\appGraz3264.accdb"
```

**Razón:** Los datos cambian constantemente, generando commits innecesarios. Solo versionar estructura.

### 2. Migración completa de base de datos

**Recomendación:** Exportar datos

```powershell
.\access-export-git-FIXED.ps1 -DatabasePath "C:\migrar\database.accdb" -ExportTableData
```

**Razón:** Necesitas mover datos + estructura a otro entorno.

### 3. Backup completo

**Recomendación:** Exportar datos

```powershell
.\access-export-git-FIXED.ps1 -DatabasePath "C:\produccion\app.accdb" -ExportTableData
```

**Razón:** Backup completo recuperable mediante `ImportXML`.

## Ventajas de Solo Estructura

? Commits más limpios en Git  
? Repositorio más pequeño  
? Foco en cambios de código/diseño  
? Exportación más rápida

## Ventajas de Estructura + Datos

? Backup completo recuperable  
? Migración de datos incluida  
? Testing con datos reales  
? Snapshots completos

## Importación

Al importar con `access-import-changed.ps1`, se importarán automáticamente:
- `.table` (estructura)
- `.tabledata` (datos, si existe)

El sistema detecta automáticamente qué archivos están disponibles.
