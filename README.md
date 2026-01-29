# Access Analyzer Skill

Personal skill for comprehensive Microsoft Access database analysis, export, refactoring, and re-import.

## Structure

```
access-analyzer/
├── SKILL.md                          # Main skill definition
├── scripts/
│   ├── access-backup.ps1             # Create timestamped backups
│   ├── access-export.ps1             # Export all Access objects
│   └── access-import.ps1             # Re-import modified code
└── references/
    ├── ExportTodoSimple.bas          # VBA export module (original)
    ├── AccessObjectTypes.md          # Object type reference
    └── VBA-Patterns.md               # Refactoring patterns
```

## Quick Start

### Export an Access Database

```powershell
.\scripts\access-export.ps1 -DatabasePath "C:\path\to\database.accdb"
```

### Create Backup

```powershell
.\scripts\access-backup.ps1 -DatabasePath "C:\path\to\database.accdb"
```

### Re-import Changes

```powershell
.\scripts\access-import.ps1 -DatabasePath "C:\path\to\database.accdb" -SourceFolder "C:\path\to\Exportacion_xxx"
```

## Features

- ✅ Automated backup with timestamps
- ✅ Complete object export (tables, queries, forms, reports, macros, VBA)
- ✅ UTF-8 encoding for perfect VS Code compatibility
- ✅ Spanish character support (á, é, í, ó, ú, ñ)
- ✅ Refactoring guidelines and patterns
- ✅ Re-import modified VBA and queries
- ✅ Detailed logging

## Requirements

- Microsoft Access installed
- PowerShell 5.1 or higher
- "Trust access to VBA project object model" enabled in Access Trust Center

## License

MIT
