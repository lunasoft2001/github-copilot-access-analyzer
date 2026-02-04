# 📚 Índice de Documentación

Guía rápida para navegar la documentación del proyecto.

---

## 🚀 Quick Start

1. **[README.md](README.md)** - Comienza aquí
2. **[SETUP.md](SETUP.md)** - Instalación inicial
3. **[SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md)** - Guía de scripts PowerShell
4. **[examples/QUICK_START.md](examples/QUICK_START.md)** - Ejemplo práctico

---

## 📖 Documentación Principal

### ⭐ Esencial
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [README.md](README.md) | Documentación principal del proyecto | Primero |
| [SKILL.md](SKILL.md) | Definición del skill para GitHub Copilot | Para entender el skill |
| [SETUP.md](SETUP.md) | Instalación y configuración inicial | Antes de empezar |

### 📋 Standard
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [CHANGELOG.md](CHANGELOG.md) | Registro de cambios detallado | Ver mejoras/actualizaciones |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Guía para contribuir | Si vas a contribuir |
| [CONTRIBUTORS.md](CONTRIBUTORS.md) | Lista de contribuyentes | Conocer autores |

### 📖 Guías
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) | Guía completa de scripts PowerShell | Usar scripts |
| [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) | Workflow con Git | Trabajar con Git |

---

## 🎯 Por Caso de Uso

### 💻 "Quiero empezar a usar el skill"
1. [README.md](README.md) - Overview
2. [SETUP.md](SETUP.md) - Instalación
3. [examples/QUICK_START.md](examples/QUICK_START.md) - Primer export/import
4. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - Referencia de scripts

### 🔧 "Necesito configurar el entorno"
1. [SETUP.md](SETUP.md) - Instrucciones completas
2. [CHANGELOG.md](CHANGELOG.md) - Sección "Migración Requerida"

### 📊 "Quiero exportar mi base de datos"
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - access-export-git.ps1
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) - Workflow completo

### 📥 "Quiero importar cambios"
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - access-import.ps1
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) - Import workflow

### 🌍 "Quiero usar en otro idioma"
1. [CHANGELOG.md](CHANGELOG.md) - Sección "Multiidioma"
2. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - Parámetro Language

### 🐛 "Tengo un problema"
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - Sección Troubleshooting
2. [CHANGELOG.md](CHANGELOG.md) - Problemas Conocidos

### 🔄 "¿Qué ha cambiado?"
1. [CHANGELOG.md](CHANGELOG.md) - Registro completo de cambios

---

## 📁 Estructura del Proyecto

```
github-copilot-access-analyzer/
├── 📄 README.md                    ⭐ Documentación principal
├── 📄 SKILL.md                     ⭐ Definición del skill
├── 📄 SETUP.md                     ⭐ Instalación
├── 📄 CHANGELOG.md                 📋 Registro de cambios
├── 📄 CONTRIBUTING.md              📋 Guía contribución
├── 📄 CONTRIBUTORS.md              📋 Contribuyentes
├── 📄 SCRIPTS_REFERENCIA.md        📖 Guía scripts
├── 📄 README_GIT_WORKFLOW.md       📖 Workflow Git
├── 📄 INDEX.md                     📚 Este archivo
│
├── 📁 modules/                     Módulos VBA
│   ├── ModExportComplete.bas       Export con multiidioma
│   └── ModImportComplete.bas       Import con multiidioma
│
├── 📁 scripts/                     Scripts PowerShell
│   ├── access-backup.ps1           Backups
│   ├── access-export-git.ps1       Export principal ⭐
│   ├── access-import.ps1           Import completo ⭐
│   └── access-import-changed.ps1   Import inteligente
│
├── 📁 docs/                        Documentación adicional
│   ├── INSTALLATION.md
│   └── WORKFLOW.md
│
├── 📁 examples/                    Ejemplos y tutoriales
│   └── QUICK_START.md
│
└── 📁 references/                  Referencias técnicas
    ├── AccessObjectTypes.md
    ├── VBA-Patterns.md
    └── ExportTodoSimple.bas
```

---

## 🎓 Niveles de Conocimiento

### 👶 Principiante
**Nunca he usado este skill**
1. [README.md](README.md)
2. [SETUP.md](SETUP.md)
3. [examples/QUICK_START.md](examples/QUICK_START.md)

### 🧑 Intermedio
**Ya exporté/importé algunas veces**
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md)
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md)
3. [CHANGELOG.md](CHANGELOG.md) - Multiidioma

### 👨‍💻 Avanzado
**Quiero contribuir o personalizar**
1. [CONTRIBUTING.md](CONTRIBUTING.md)
2. [CHANGELOG.md](CHANGELOG.md) - Detalle técnico
3. [references/VBA-Patterns.md](references/VBA-Patterns.md)

---

## 🔍 Búsqueda Rápida

### Temas Comunes

| Busco... | Ver... |
|----------|--------|
| Instalación | [SETUP.md](SETUP.md) |
| Primer uso | [examples/QUICK_START.md](examples/QUICK_START.md) |
| Scripts PowerShell | [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) |
| Multiidioma | [CHANGELOG.md](CHANGELOG.md#-multiidioma) |
| Git workflow | [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) |
| Troubleshooting | [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md#-troubleshooting) |
| Cambios recientes | [CHANGELOG.md](CHANGELOG.md) |
| Contribuir | [CONTRIBUTING.md](CONTRIBUTING.md) |

---

## 📞 Soporte

- **Problemas**: [GitHub Issues](https://github.com/lunasoft2001/github-copilot-access-analyzer/issues)
- **Preguntas**: Ver primero [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) → Troubleshooting
- **Email**: Juanjo@luna-soft.es

---

## ✨ Última Actualización

**Fecha**: 2026-02-04  
**Cambios**: Consolidación de documentación, eliminación de archivos redundantes  
**Ver**: [CHANGELOG.md](CHANGELOG.md)

---

**Made with ❤️ for developers working with Microsoft Access**

