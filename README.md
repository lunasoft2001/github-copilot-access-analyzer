#  GitHub Copilot Access Analyzer Skill

A comprehensive GitHub Copilot skill for analyzing, exporting, refactoring, and re-importing Microsoft Access database applications with intelligent Git-based version control.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/en-us/powershell/)
[![Access](https://img.shields.io/badge/MS%20Access-2010%2B-red)](https://www.microsoft.com/en-us/microsoft-365/access)

##  Overview

This skill enables GitHub Copilot to work seamlessly with Microsoft Access databases (.accdb/.mdb) by providing automated workflows for:

- ** Intelligent Export**: Extract all database objects (tables, queries, forms, reports, macros, VBA) to version-controlled text files
- ** Refactoring**: Analyze and improve VBA code in VS Code with modern development tools
- ** Selective Import**: Re-import only modified objects back to Access using Git diff detection
- ** Safety**: Automated backups, dry-run modes, and detailed logging
- ** Planning**: Auto-generated refactoring plans with checklists and progress tracking

##  Key Features

-  **Git Integration**: Full version control workflow with automatic commit tracking
-  **UTF-8 Support**: Perfect Spanish character handling (á, é, í, ó, ú, ñ)
-  **Selective Import**: Only import changed files (detected via Git diff)
-  **Refactoring Plans**: Auto-generated `REFACTORING_PLAN.md` with checklists
-  **Structured Export**: Organized folder hierarchy for all object types
-  **Automated Backups**: Timestamped backups before destructive operations
-  **Dry-Run Mode**: Preview import operations before execution
-  **VS Code Integration**: Seamless workflow with modern code editor

##  Quick Start

### Prerequisites

- Microsoft Access 2010+ installed
- PowerShell 5.1 or higher
- Git (optional, but recommended for best workflow)
- VS Code with GitHub Copilot (for AI-assisted refactoring)

### Installation

1. **Enable VBA Project Access** (required):
   - Open Access  File  Options  Trust Center  Trust Center Settings
   - Check "Trust access to the VBA project object model"

2. **Install the Skill**:
   ```bash
   # Clone or download this repository to your Copilot skills folder
   cd ~/.copilot/skills  # or %USERPROFILE%\.copilot\skills on Windows
   git clone https://github.com/lunasoft2001/github-copilot-access-analyzer.git access-analyzer
   ```

3. **Restart VS Code** to load the skill

### Basic Usage with GitHub Copilot

Simply ask GitHub Copilot in VS Code:

```
"Export my Access database at C:\projects\inventory.accdb"
"Refactor the VBA code in my Access database"
"Import changes back to Access from the export folder"
```

Copilot will automatically use this skill to handle Access database operations.

##  Workflow Examples

### Recommended Workflow (with Git)

```powershell
# 1. Export database with Git version control
cd $env:USERPROFILE\.copilot\skills\access-analyzer\scripts
.\access-export-git.ps1 -DatabasePath "C:\projects\myapp.accdb"

# 2. Open in VS Code (prompted automatically)
# Review REFACTORING_PLAN.md for guided workflow

# 3. Refactor code in VS Code

# 4. Preview changes before import (dry-run)
.\access-import-changed.ps1 -TargetDbPath "C:\projects\myapp.accdb" `
                             -ExportFolder "C:\projects\myapp_Export" `
                             -DryRun

# 5. Import only modified files
.\access-import-changed.ps1 -TargetDbPath "C:\projects\myapp.accdb" `
                             -ExportFolder "C:\projects\myapp_Export"
```

##  Documentation

- [Installation Guide](./docs/INSTALLATION.md) - Detailed setup instructions
- [Workflow Guide](./docs/WORKFLOW.md) - Step-by-step workflows
- [Quick Start Example](./examples/QUICK_START.md) - Complete tutorial
- [Scripts Reference](./SKILL.md) - Complete script documentation
- [VBA Patterns](./references/VBA-Patterns.md) - Refactoring guidelines
- [Publishing Guide](./PUBLISH_GUIDE.md) - How to share your own skills

##  Contributing

Contributions are welcome! See [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

##  Documentation

| Document | Description |
|----------|-------------|
| [SKILL.md](./SKILL.md) | GitHub Copilot skill definition and usage |
| [SETUP.md](./SETUP.md) | Installation and initial setup instructions |
| [SCRIPTS_REFERENCIA.md](./SCRIPTS_REFERENCIA.md) | Complete PowerShell scripts guide |
| [README_GIT_WORKFLOW.md](./README_GIT_WORKFLOW.md) | Git workflow and best practices |
| [CHANGELOG.md](./CHANGELOG.md) | Detailed changelog and upgrade guide |
| [CONTRIBUTING.md](./CONTRIBUTING.md) | Contribution guidelines |
| [docs/](./docs/) | Additional technical documentation |
| [examples/](./examples/) | Usage examples and tutorials |

##  License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

##  Author

**Juanjo Luna**
- Email: [Juanjo@luna-soft.es](mailto:Juanjo@luna-soft.es)
- GitHub: [@lunasoft2001](https://github.com/lunasoft2001)

##  Support

-  [Report Issues](https://github.com/lunasoft2001/github-copilot-access-analyzer/issues)
-  [Discussions](https://github.com/lunasoft2001/github-copilot-access-analyzer/discussions)
-  [Documentation](./docs/)

---

**Made with  for developers working with Microsoft Access in the age of AI**