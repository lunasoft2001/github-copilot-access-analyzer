# Quick Start Example

Complete walkthrough of exporting, refactoring, and importing an Access database.

## Step 1: Export

```powershell
.\scripts\access-export-git.ps1 -DatabasePath "C:\MyApp.accdb"
```

This creates `MyApp_Export/` folder with Git initialized.

## Step 2: Refactor in VS Code

1. Open VS Code in export folder
2. Review `REFACTORING_PLAN.md`
3. Edit VBA files in `06_Codigo_VBA/`
4. Save changes

## Step 3: Import Changes

```powershell
.\scripts\access-import-changed.ps1 -TargetDbPath "C:\MyApp.accdb" `
                                     -ExportFolder "C:\MyApp_Export"
```

See [WORKFLOW.md](../docs/WORKFLOW.md) for complete guide.
