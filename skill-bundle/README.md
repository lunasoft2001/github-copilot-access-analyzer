# Access Analyzer Skill Bundle

Este directorio contiene el **skill bundle limpio** listo para instalar en GitHub Copilot.

## Estructura

```
skill-bundle/
??? SKILL.md                      # Metadata del skill
??? scripts/                      # Scripts PowerShell de automatización
?   ??? access-backup.ps1        # Crear backups
?   ??? access-export-git.ps1    # Exportar con integración Git
?   ??? access-import-changed.ps1 # Importar solo objetos modificados
?   ??? access-import.ps1        # Importar todos los objetos
??? references/                   # Documentación de referencia
?   ??? AccessObjectTypes.md     # Tipos de objetos en Access
?   ??? ExportTodoSimple.bas     # Módulo VBA de exportación
?   ??? VBA-Patterns.md          # Patrones de código VBA
??? assets/                       # Recursos del skill
    ??? AccessAnalyzer.accdb     # Base de datos de ejemplo/template
```

## Instalación

Para instalar este skill en GitHub Copilot, usa el script de instalación automatizado en la raíz del repositorio:

```powershell
cd ..
.\install-skill.ps1
```

Para instrucciones detalladas, consulta [SKILL_INSTALLATION.md](../SKILL_INSTALLATION.md).

## ¿Por qué un bundle separado?

Este repositorio contiene archivos adicionales para desarrollo (docs/, examples/, README, CHANGELOG, etc.) que **no deben incluirse** en el skill instalado según las guías de skill-creator.

El `skill-bundle/` contiene **solo** los archivos esenciales que Copilot necesita:

- **SKILL.md**: Metadata y documentación del skill
- **scripts/**: Herramientas ejecutables
- **references/**: Documentación de soporte
- **assets/**: Recursos necesarios (bases de datos, templates)

## Instalación Manual

Si prefieres instalar manualmente sin el script:

```powershell
# Copiar el bundle completo a la carpeta de skills de Copilot
Copy-Item -Path "skill-bundle" -Destination "$env:USERPROFILE\.copilot\skills\access-analyzer" -Recurse
```

Luego reinicia VS Code.

## Notas

- Este bundle está optimizado para GitHub Copilot
- Cumple con los requisitos de skill-creator
- Contiene solo archivos esenciales (sin README, sin docs adicionales)
- La base de datos AccessAnalyzer.accdb está en `assets/` como recomienda skill-creator

---

**Instalación completa:** [SKILL_INSTALLATION.md](../SKILL_INSTALLATION.md)
