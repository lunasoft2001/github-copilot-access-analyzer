# INSTRUCCIONES PARA CREAR AccessAnalyzer.accdb

## Ubicación del archivo
Crear el archivo Access aquí:
`C:\Users\juanjo_admin\.copilot\skills\access-analyzer\AccessAnalyzer.accdb`

## Pasos para crear el archivo herramienta:

1. **Crear nuevo archivo Access:**
   - Abre Microsoft Access
   - Crear base de datos en blanco
   - Guardar como: `C:\Users\juanjo_admin\.copilot\skills\access-analyzer\AccessAnalyzer.accdb`

2. **Importar módulos VBA:**
   - Presiona Alt+F11 (abrir editor VBA)
   - Archivo → Importar Archivo
   - Importa los siguientes módulos:
     * `modules\ModExportExternal.bas` - Exportación desde archivo externo
     * `modules\ModExportVBALocal.bas` - Exportación VBA local

3. **Verificar importación:**
   - En el editor VBA, verifica que aparezcan:
     * ModExportExternal
     * ModExportVBALocal

4. **Guardar y cerrar**

## Uso del archivo creado:

### Desde PowerShell (automático):
```powershell
# Exportar otro archivo Access
.\scripts\access-export-tool.ps1 -TargetDatabase "C:\path\to\database.accdb"
```

### Desde Access (manual):
1. Abre `AccessAnalyzer.accdb`
2. Alt+F11 (VBA)
3. Ctrl+G (ventana Inmediato)
4. Ejecuta:
```vba
ExportExternalDatabase "C:\path\to\database.accdb", "C:\path\to\output"
```

## Módulos creados:

### ModExportExternal.bas
- **Función principal:** `ExportExternalDatabase(sourceDbPath, outputFolder)`
- **Propósito:** Exportar tablas, consultas, y metadatos de OTRO archivo Access
- **Limitación:** No puede exportar VBA de archivos externos (Access no lo permite)

### ModExportVBALocal.bas
- **Función principal:** `ExportAllVBALocal(outputFolder)`
- **Propósito:** Exportar código VBA cuando se ejecuta DENTRO del archivo objetivo
- **Uso:** Importar este módulo al archivo que quieres analizar y ejecutarlo desde allí

## Próximo paso:
Una vez hayas creado el archivo, ejecuta:
```powershell
Test-Path "C:\Users\juanjo_admin\.copilot\skills\access-analyzer\AccessAnalyzer.accdb"
```

Esto debería devolver `True` cuando el archivo esté listo.
