# ?? Índice de Documentación

Guía rápida para navegar la documentación del proyecto.

---

## ?? Quick Start

1. **[README.md](README.md)** - Comienza aquí
2. **[SKILL_INSTALLATION.md](SKILL_INSTALLATION.md)** - Instalación del skill
3. **[SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md)** - Guía de scripts PowerShell
4. **[examples/QUICK_START.md](examples/QUICK_START.md)** - Ejemplo práctico

---

## ?? Documentación Principal

### ? Esencial
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [README.md](README.md) | Documentación principal del proyecto | Primero |
| [skill-bundle/SKILL.md](skill-bundle/SKILL.md) | Definición del skill para GitHub Copilot | Para entender el skill |
| [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md) | Guía de instalación completa | Antes de instalar |
| [SETUP.md](SETUP.md) | Configuración inicial de Access | Después de instalar |

### ?? Standard
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [CHANGELOG.md](CHANGELOG.md) | Registro de cambios detallado | Ver mejoras/actualizaciones |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Guía para contribuir | Si vas a contribuir |
| [CONTRIBUTORS.md](CONTRIBUTORS.md) | Lista de contribuyentes | Conocer autores |

### ?? Guías
| Documento | Descripción | Cuándo Leer |
|-----------|-------------|-------------|
| [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) | Guía completa de scripts PowerShell | Usar scripts |
| [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) | Workflow con Git | Trabajar con Git |
| [CLEAN_STRUCTURE.md](CLEAN_STRUCTURE.md) | Estructura del proyecto | Entender organización |

---

## ?? Por Caso de Uso

### ?? "Quiero instalar el skill"
1. [README.md](README.md) - Overview
2. [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md) - Instalación automática
3. Ejecutar `.\install-skill.ps1`
4. Reiniciar VS Code

### ?? "Necesito configurar Access"
1. [SETUP.md](SETUP.md) - Instrucciones completas
2. Habilitar "Trust access to the VBA project object model"

### ?? "Quiero exportar mi base de datos"
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - access-export-git.ps1
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) - Workflow completo

### ?? "Quiero importar cambios"
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - access-import.ps1
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) - Import workflow

### ?? "Quiero usar en otro idioma"
1. [CHANGELOG.md](CHANGELOG.md) - Sección "Multiidioma"
2. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - Parámetro Language

### ?? "Tengo un problema"
1. [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md) - Troubleshooting
2. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) - Sección Troubleshooting
3. [CHANGELOG.md](CHANGELOG.md) - Problemas Conocidos

### ?? "¿Qué ha cambiado?"
1. [CHANGELOG.md](CHANGELOG.md) - Registro completo de cambios

---

## ?? Estructura del Proyecto

```
github-copilot-access-analyzer/
??? ?? README.md                    ? Documentación principal
??? ?? SKILL_INSTALLATION.md        ? Guía de instalación
??? ?? install-skill.ps1            ?? Instalador automatizado
??? ?? SETUP.md                     ? Configuración Access
??? ?? CHANGELOG.md                 ?? Registro de cambios
??? ?? CONTRIBUTING.md              ?? Guía contribución
??? ?? CONTRIBUTORS.md              ?? Contribuyentes
??? ?? SCRIPTS_REFERENCIA.md        ?? Guía scripts
??? ?? README_GIT_WORKFLOW.md       ?? Workflow Git
??? ?? CLEAN_STRUCTURE.md           ?? Estructura del proyecto
??? ?? INDEX.md                     ?? Este archivo
?
??? ?? skill-bundle/                ? Skill limpio para instalar
?   ??? SKILL.md                   Metadata del skill
?   ??? scripts/                   Scripts PowerShell
?   ?   ??? access-backup.ps1
?   ?   ??? access-export-git.ps1  ? Export principal
?   ?   ??? access-import.ps1      ? Import completo
?   ?   ??? access-import-changed.ps1
?   ??? references/                Referencias técnicas
?   ?   ??? AccessObjectTypes.md
?   ?   ??? VBA-Patterns.md
?   ?   ??? ExportTodoSimple.bas
?   ??? assets/                    Recursos
?       ??? AccessAnalyzer.accdb
?
??? ?? modules/                     Módulos VBA de desarrollo
?   ??? ModExportComplete_v2.bas
?   ??? ModExportComplete.bas
?   ??? ModImportComplete.bas
?
??? ?? docs/                        Documentación adicional
?   ??? INSTALLATION.md
?   ??? WORKFLOW.md
?
??? ?? examples/                    Ejemplos y tutoriales
    ??? QUICK_START.md
```

---

## ?? Niveles de Conocimiento

### ?? Principiante
**Nunca he usado este skill**
1. [README.md](README.md)
2. [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md)
3. [SETUP.md](SETUP.md)
4. [examples/QUICK_START.md](examples/QUICK_START.md)

### ?? Intermedio
**Ya exporté/importé algunas veces**
1. [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md)
2. [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md)
3. [CHANGELOG.md](CHANGELOG.md) - Multiidioma

### ????? Avanzado
**Quiero contribuir o personalizar**
1. [CONTRIBUTING.md](CONTRIBUTING.md)
2. [CLEAN_STRUCTURE.md](CLEAN_STRUCTURE.md)
3. [CHANGELOG.md](CHANGELOG.md) - Detalle técnico
4. [skill-bundle/references/VBA-Patterns.md](skill-bundle/references/VBA-Patterns.md)

---

## ?? Búsqueda Rápida

### Temas Comunes

| Busco... | Ver... |
|----------|--------|
| Instalación | [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md) |
| Primer uso | [examples/QUICK_START.md](examples/QUICK_START.md) |
| Scripts PowerShell | [SCRIPTS_REFERENCIA.md](SCRIPTS_REFERENCIA.md) |
| Multiidioma | [CHANGELOG.md](CHANGELOG.md#-multiidioma) |
| Git workflow | [README_GIT_WORKFLOW.md](README_GIT_WORKFLOW.md) |
| Troubleshooting | [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md#-troubleshooting) |
| Cambios recientes | [CHANGELOG.md](CHANGELOG.md) |
| Contribuir | [CONTRIBUTING.md](CONTRIBUTING.md) |
| Estructura proyecto | [CLEAN_STRUCTURE.md](CLEAN_STRUCTURE.md) |

---

## ?? Soporte

- **Problemas**: [GitHub Issues](https://github.com/lunasoft2001/github-copilot-access-analyzer/issues)
- **Preguntas**: Ver primero [SKILL_INSTALLATION.md](SKILL_INSTALLATION.md) ? Troubleshooting
- **Email**: Juanjo@luna-soft.es

---

## ? Última Actualización

**Fecha**: 2026-02-05  
**Cambios**: Skill bundle limpio, script de instalación automatizada, corrección de codificación  
**Ver**: [CHANGELOG.md](CHANGELOG.md)

---

**Made with ?? for developers working with Microsoft Access**

