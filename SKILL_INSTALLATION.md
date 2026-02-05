# ?? Instalación del Skill Access Analyzer

Guía completa para instalar el skill **access-analyzer** en GitHub Copilot.

## ?? ¿Qué es este skill?

El skill **access-analyzer** permite a GitHub Copilot trabajar con bases de datos Microsoft Access:

- ? Crear backups automáticos
- ? Exportar todos los objetos (tablas, consultas, formularios, informes, macros, VBA) a archivos de texto
- ? Analizar y refactorizar código VBA en VS Code
- ? Re-importar cambios a la base de datos Access
- ? Control de versiones para aplicaciones Access

## ?? Requisitos Previos

1. **VS Code** con la extensión **GitHub Copilot** instalada y activa
2. **PowerShell** 5.1 o superior (incluido en Windows 10/11)
3. **Microsoft Access** instalado (para trabajar con bases de datos .accdb/.mdb)

## ?? Instalación Automática (Recomendado)

### Paso 1: Clonar o descargar el repositorio

```powershell
# Clonar con Git
git clone https://github.com/your-username/github-copilot-access-analyzer.git
cd github-copilot-access-analyzer

# O descargar ZIP y extraer
```

### Paso 2: Ejecutar el script de instalación

Abre PowerShell en la carpeta del repositorio y ejecuta:

```powershell
.\install-skill.ps1
```

El script automáticamente:
- Verifica que existe `skill-bundle/`
- Copia todo a `%USERPROFILE%\.copilot\skills\access-analyzer`
- Muestra la estructura instalada
- Te pide confirmar si el skill ya está instalado

**Salida esperada:**
```
???????????????????????????????????????????????????????
  GitHub Copilot Skill Installer - Access Analyzer
???????????????????????????????????????????????????????

? Verificando prerrequisitos...
? Prerrequisitos verificados

? Instalando skill 'access-analyzer'...
? Skill instalado correctamente en: C:\Users\Usuario\.copilot\skills\access-analyzer

???????????????????????????????????????????????????????
  Instalación completada exitosamente
???????????????????????????????????????????????????????

? Próximos pasos:
  1. Cierra y reinicia VS Code
  2. El skill 'access-analyzer' estará disponible en Copilot
  3. Prueba preguntando: 'Exporta mi base de datos Access'
```

### Paso 3: Reiniciar VS Code

**Importante:** Cierra completamente VS Code y vuelve a abrirlo para que Copilot detecte el nuevo skill.

### Paso 4: Verificar instalación

Pregunta a Copilot:

```
Dime qué skills están disponibles
```

Deberías ver **access-analyzer** en la lista.

## ?? Otras Acciones con el Script

### Actualizar el skill

Si ya tienes el skill instalado y quieres actualizar a una nueva versión:

```powershell
.\install-skill.ps1 -Action Update
```

Esto crea un backup automático de la versión anterior antes de actualizar.

### Verificar instalación

Para comprobar que el skill está correctamente instalado:

```powershell
.\install-skill.ps1 -Action Verify
```

**Verifica:**
- Existencia de SKILL.md con frontmatter correcto
- Carpetas `scripts/`, `references/`, `assets/`
- Presencia de AccessAnalyzer.accdb

### Desinstalar el skill

Para eliminar completamente el skill:

```powershell
.\install-skill.ps1 -Action Uninstall
```

Solicita confirmación antes de eliminar.

## ?? Instalación Manual (Alternativa)

Si prefieres no usar el script, puedes instalar manualmente:

### Paso 1: Ubicar la carpeta de skills

La carpeta de skills de Copilot se encuentra en:

```
Windows: C:\Users\<TuUsuario>\.copilot\skills
macOS/Linux: ~/.copilot/skills
```

Si no existe, créala:

```powershell
New-Item -ItemType Directory -Path "$env:USERPROFILE\.copilot\skills" -Force
```

### Paso 2: Copiar el skill bundle

Copia toda la carpeta `skill-bundle/` del repositorio a la carpeta de skills:

```powershell
Copy-Item -Path ".\skill-bundle" -Destination "$env:USERPROFILE\.copilot\skills\access-analyzer" -Recurse
```

### Paso 3: Verificar estructura

La estructura final debe ser:

```
%USERPROFILE%\.copilot\skills\access-analyzer/
??? SKILL.md
??? scripts/
?   ??? access-backup.ps1
?   ??? access-export-git.ps1
?   ??? access-import-changed.ps1
?   ??? access-import.ps1
??? references/
?   ??? AccessObjectTypes.md
?   ??? ExportTodoSimple.bas
?   ??? VBA-Patterns.md
??? assets/
    ??? AccessAnalyzer.accdb
```

### Paso 4: Reiniciar VS Code

Cierra completamente y vuelve a abrir VS Code.

## ? Verificación Post-Instalación

### 1. Listar skills disponibles

Pregunta a Copilot:

```
¿Qué skills tienes disponibles?
```

Deberías ver **access-analyzer** listado.

### 2. Probar el skill

Pregunta a Copilot algo relacionado con Access:

```
Analiza esta base de datos Access y exporta todos sus objetos
```

Copilot debería reconocer que necesita usar el skill **access-analyzer**.

### 3. Verificar con el script

```powershell
.\install-skill.ps1 -Action Verify
```

## ?? Troubleshooting

### El skill no aparece en la lista de skills disponibles

**Causas posibles:**
- VS Code no se reinició completamente
- SKILL.md tiene frontmatter incorrecto
- La carpeta no está en la ruta correcta

**Soluciones:**
1. Cierra **todas** las ventanas de VS Code (incluido el System Tray)
2. Verifica el frontmatter del SKILL.md:
   ```yaml
   ---
   name: access-analyzer
   description: 'Analyze, export, refactor, and re-import Microsoft Access database applications...'
   ---
   ```
3. Verifica la ruta: `%USERPROFILE%\.copilot\skills\access-analyzer\SKILL.md` debe existir
4. Ejecuta: `.\install-skill.ps1 -Action Verify`

### Error "No se encontró la carpeta skill-bundle"

**Causa:** El script se ejecutó desde una carpeta incorrecta.

**Solución:** Ejecuta el script desde la raíz del repositorio:
```powershell
cd e:\datos\GitHub\github-copilot-access-analyzer
.\install-skill.ps1
```

### Error "Access denied" al copiar archivos

**Causa:** Permisos insuficientes o archivos bloqueados.

**Soluciones:**
1. Ejecuta PowerShell como **Administrador**
2. Cierra todas las instancias de VS Code
3. Verifica que no tengas AccessAnalyzer.accdb abierto en Access

### El skill aparece pero Copilot no lo usa

**Causa:** El frontmatter o la descripción del skill no coinciden con tu pregunta.

**Solución:** Sé explícito en tus preguntas:
- ? "Usando el skill access-analyzer, exporta esta base de datos"
- ? "Analiza mi archivo .accdb"
- ? "Ayúdame con esta base de datos" (demasiado genérico)

### AccessAnalyzer.accdb no se puede abrir

**Causa:** El archivo puede estar bloqueado por Windows.

**Solución:**
1. Click derecho en `AccessAnalyzer.accdb` ? Propiedades
2. Si hay un checkbox "Desbloquear", márcalo y aplica
3. O copia el archivo desde otro origen

## ?? Uso del Skill

Una vez instalado, puedes usar el skill preguntando a Copilot cosas como:

### Ejemplos de uso:

1. **Exportar una base de datos:**
   ```
   Exporta todos los objetos de C:\MiBD.accdb a archivos de texto para control de versiones
   ```

2. **Crear backup:**
   ```
   Crea un backup de mi base de datos Access antes de modificarla
   ```

3. **Analizar estructura:**
   ```
   Analiza la estructura de esta base de datos .accdb y dame un resumen
   ```

4. **Refactoring workflow:**
   ```
   Exporta esta base de datos, refactoriza el código VBA y vuelve a importarlo
   ```

5. **Ver diferencias:**
   ```
   Compara la versión exportada con el último commit en Git
   ```

## ?? Actualización del Skill

Cuando descargues una nueva versión del repositorio:

```powershell
git pull
.\install-skill.ps1 -Action Update
```

Esto creará un backup automático de la versión anterior antes de actualizar.

## ??? Desinstalación

Para eliminar completamente el skill:

```powershell
.\install-skill.ps1 -Action Uninstall
```

O manualmente:

```powershell
Remove-Item "$env:USERPROFILE\.copilot\skills\access-analyzer" -Recurse -Force
```

Luego reinicia VS Code.

## ?? Soporte

Si tienes problemas:

1. Ejecuta `.\install-skill.ps1 -Action Verify` y revisa los errores
2. Consulta la sección de Troubleshooting
3. Abre un issue en el repositorio con:
   - Salida del comando `Verify`
   - Versión de VS Code
   - Versión de GitHub Copilot
   - Pasos que seguiste

## ?? Licencia

Este skill está bajo licencia MIT. Ver [LICENSE](LICENSE) para más detalles.

---

**¡Listo!** Una vez instalado, GitHub Copilot podrá ayudarte a trabajar con bases de datos Access de manera profesional. ??
