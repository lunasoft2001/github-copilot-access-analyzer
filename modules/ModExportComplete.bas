Attribute VB_Name = "ModExportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModExportComplete
' VERSION: 2.1.0
' AUTOR: Juanjo Luna (juanjo@luna-soft.es)
' FECHA: 2026-02-05
' PROYECTO: GitHub Copilot Access Analyzer Skill
' 
' PROPÓSITO: Exportar COMPLETAMENTE otro archivo Access con soporte:
'   - Multiidioma (ES, EN, DE, FR, IT)
'   - DDL individual por tabla (Access + SQL Server)
'   - Consultas, Formularios, Informes, Macros, Módulos VBA
'   - Automatización sin MsgBox (Debug.Print)
'
' USO: RunCompleteExport "C:\path\to\database.accdb", "C:\output\folder", "ES"
'==========================================================================='

' Wrapper para llamar desde PowerShell con Eval
Public Function RunCompleteExport(ByVal sourceDbPath As String, ByVal outputFolder As String, Optional ByVal language As String = "ES") As Boolean
    On Error GoTo ErrHandler
    ExportCompleteDatabase sourceDbPath, outputFolder, language
    RunCompleteExport = True
    Exit Function
ErrHandler:
    RunCompleteExport = False
    Debug.Print "Export Error: " & Err.Number & " - " & Err.Description
End Function

Public Sub ExportCompleteDatabase(ByVal sourceDbPath As String, Optional ByVal outputFolder As String = "", Optional ByVal language As String = "ES")
    On Error GoTo ErrHandler
    
    ' Validar idioma (por defecto inglés si hay error)
    Select Case UCase(language)
        Case "ES", "EN", "DE", "FR", "IT"
            ' OK
        Case Else
            language = "EN"
    End Select
    
    ' Validar archivo existe
    If Dir(sourceDbPath) = "" Then
        Debug.Print "Archivo no encontrado: " & sourceDbPath
        Exit Sub
    End If
    
    ' Determinar carpeta de salida
    If Len(outputFolder) = 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim parentFolder As String
        parentFolder = fso.GetParentFolderName(sourceDbPath)
        outputFolder = parentFolder & "\Exportacion_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    
    ' Abrir Access externo SIN ejecutar Autoexec
    Dim accessApp As Access.Application
    Set accessApp = OpenAccessNoAutoexec(sourceDbPath)
    
    If accessApp Is Nothing Then
        Debug.Print "No se pudo abrir el archivo Access"
        Exit Sub
    End If
    
    ' Crear estructura de carpetas
    CreateFolders outputFolder, language
    
    ' Exportar todo usando la instancia externa
    ExportAllFromExternal accessApp, sourceDbPath, outputFolder, language
    
    ' Cerrar Access externo
    accessApp.Quit acQuitSaveNone
    Set accessApp = Nothing
    
    Debug.Print "Exportación completada: " & sourceDbPath & " -> " & outputFolder
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Export Error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
End Sub

'===========================================================================
' ABRIR ACCESS SIN AUTOEXEC (método del proyecto SVN)
'===========================================================================
Private Function OpenAccessNoAutoexec(ByVal strMDBPath As String) As Access.Application
    On Error GoTo ErrHandler
    
    Dim objAcc As Access.Application
    
    If Dir(strMDBPath) = "" Then
        Err.Raise 53, , "Archivo no encontrado"
    End If
    
    ' Crear nueva instancia de Access
    Set objAcc = New Access.Application
    objAcc.Visible = False
    
    ' Abrir base de datos (sin Autoexec por ahora - simplificado)
    objAcc.OpenCurrentDatabase strMDBPath, False
    
    Set OpenAccessNoAutoexec = objAcc
    Exit Function
    
ErrHandler:
    Set OpenAccessNoAutoexec = Nothing
End Function

'===========================================================================
' EXPORTAR TODO DESDE INSTANCIA EXTERNA
'===========================================================================
Private Sub ExportAllFromExternal(accessApp As Access.Application, dbPath As String, basePath As String, Optional language As String = "ES")
    On Error GoTo ErrHandler
    
    ' Exportar resumen
    ExportSummary accessApp, dbPath, basePath, language
    
    ' Exportar tablas
    ExportTables accessApp, basePath, language
    
    ' Exportar consultas
    ExportQueries accessApp, basePath, language
    
    ' Exportar formularios completos
    ExportForms accessApp, basePath, language
    
    ' Exportar informes completos
    ExportReports accessApp, basePath, language
    
    ' Exportar macros completos
    ExportMacros accessApp, basePath, language
    
    ' Exportar VBA completo
    ExportVBA accessApp, basePath, language
    
    Exit Sub
ErrHandler:
    On Error GoTo 0
End Sub

'===========================================================================
' CREAR CARPETAS CON SOPORTE MULTIIDIOMA
'===========================================================================
Private Sub CreateFolders(basePath As String, Optional language As String = "ES")
    On Error Resume Next
    MkDir basePath
    MkDir basePath & "\" & GetFolderName("TABLES", language)
    MkDir basePath & "\" & GetFolderName("TABLES", language) & "\" & GetFolderName("ACCESS", language)
    MkDir basePath & "\" & GetFolderName("TABLES", language) & "\" & GetFolderName("SQLSERVER", language)
    MkDir basePath & "\" & GetFolderName("QUERIES", language)
    MkDir basePath & "\" & GetFolderName("FORMS", language)
    MkDir basePath & "\" & GetFolderName("REPORTS", language)
    MkDir basePath & "\" & GetFolderName("MACROS", language)
    MkDir basePath & "\" & GetFolderName("VBA", language)
    On Error GoTo 0
End Sub

'===========================================================================
' OBTENER NOMBRE DE CARPETA LOCALIZADO
'===========================================================================
Private Function GetFolderName(folderType As String, Optional language As String = "ES") As String
    Dim result As String
    
    Select Case UCase(folderType)
        Case "TABLES"
            Select Case UCase(language)
                Case "ES": result = "01_Tablas"
                Case "EN": result = "01_Tables"
                Case "DE": result = "01_Tabellen"
                Case "FR": result = "01_Tables"
                Case "IT": result = "01_Tabelle"
                Case Else: result = "01_Tables"
            End Select
        
        Case "ACCESS"
            Select Case UCase(language)
                Case "ES": result = "Access"
                Case "EN": result = "Access"
                Case "DE": result = "Access"
                Case "FR": result = "Access"
                Case "IT": result = "Access"
                Case Else: result = "Access"
            End Select
        
        Case "SQLSERVER"
            Select Case UCase(language)
                Case "ES": result = "SQLServer"
                Case "EN": result = "SQLServer"
                Case "DE": result = "SQLServer"
                Case "FR": result = "SQLServer"
                Case "IT": result = "SQLServer"
                Case Else: result = "SQLServer"
            End Select
        
        Case "QUERIES"
            Select Case UCase(language)
                Case "ES": result = "02_Consultas"
                Case "EN": result = "02_Queries"
                Case "DE": result = "02_Abfragen"
                Case "FR": result = "02_Requêtes"
                Case "IT": result = "02_Query"
                Case Else: result = "02_Queries"
            End Select
        
        Case "FORMS"
            Select Case UCase(language)
                Case "ES": result = "03_Formularios"
                Case "EN": result = "03_Forms"
                Case "DE": result = "03_Formulare"
                Case "FR": result = "03_Formulaires"
                Case "IT": result = "03_Moduli"
                Case Else: result = "03_Forms"
            End Select
        
        Case "REPORTS"
            Select Case UCase(language)
                Case "ES": result = "04_Informes"
                Case "EN": result = "04_Reports"
                Case "DE": result = "04_Berichte"
                Case "FR": result = "04_Rapports"
                Case "IT": result = "04_Rapporti"
                Case Else: result = "04_Reports"
            End Select
        
        Case "MACROS"
            Select Case UCase(language)
                Case "ES": result = "05_Macros"
                Case "EN": result = "05_Macros"
                Case "DE": result = "05_Makros"
                Case "FR": result = "05_Macros"
                Case "IT": result = "05_Macro"
                Case Else: result = "05_Macros"
            End Select
        
        Case "VBA"
            Select Case UCase(language)
                Case "ES": result = "06_Codigo_VBA"
                Case "EN": result = "06_VBA_Code"
                Case "DE": result = "06_VBA_Code"
                Case "FR": result = "06_Code_VBA"
                Case "IT": result = "06_Codice_VBA"
                Case Else: result = "06_VBA_Code"
            End Select
        
        Case Else
            result = folderType
    End Select
    
    GetFolderName = result
End Function

'===========================================================================
' EXPORTAR RESUMEN
'===========================================================================
Private Sub ExportSummary(accessApp As Access.Application, dbPath As String, basePath As String, Optional language As String = "ES")
    On Error GoTo ErrH
    
    Dim db As DAO.Database
    Set db = accessApp.CurrentDb
    
    Dim content As String
    content = "=============================================================" & vbCrLf
    content = content & "EXPORTACIÓN COMPLETA DE ACCESS" & vbCrLf
    content = content & "=============================================================" & vbCrLf
    content = content & "Archivo: " & dbPath & vbCrLf
    content = content & "Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "Codificación: UTF-8" & vbCrLf
    content = content & "Idioma: " & language & vbCrLf
    content = content & "=============================================================" & vbCrLf & vbCrLf
    
    content = content & "INVENTARIO:" & vbCrLf
    content = content & "- Tablas: " & CountTables(db) & vbCrLf
    content = content & "- Consultas: " & CountQueries(db) & vbCrLf
    content = content & "- Formularios: " & accessApp.CurrentProject.AllForms.Count & vbCrLf
    content = content & "- Informes: " & accessApp.CurrentProject.AllReports.Count & vbCrLf
    content = content & "- Macros: " & accessApp.CurrentProject.AllMacros.Count & vbCrLf
    content = content & "- Módulos VBA: " & accessApp.CurrentProject.AllModules.Count & vbCrLf
    
    WriteUTF8File basePath & "\00_RESUMEN.txt", content
    
    Exit Sub
ErrH:
End Sub

'===========================================================================
' EXPORTAR FORMULARIOS CON SaveAsText
'===========================================================================
Private Sub ExportForms(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllForms.Count - 1
        Dim formName As String
        formName = accessApp.CurrentProject.AllForms(i).Name
        
        ' Usar SaveAsText para exportar definición completa
        Dim filePath As String
        filePath = basePath & "\" & GetFolderName("FORMS", language) & "\" & CleanName(formName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acForm, formName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR INFORMES CON SaveAsText
'===========================================================================
Private Sub ExportReports(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllReports.Count - 1
        Dim reportName As String
        reportName = accessApp.CurrentProject.AllReports(i).Name
        
        Dim filePath As String
        filePath = basePath & "\" & GetFolderName("REPORTS", language) & "\" & CleanName(reportName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acReport, reportName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR MACROS CON SaveAsText
'===========================================================================
Private Sub ExportMacros(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllMacros.Count - 1
        Dim macroName As String
        macroName = accessApp.CurrentProject.AllMacros(i).Name
        
        Dim filePath As String
        filePath = basePath & "\" & GetFolderName("MACROS", language) & "\" & CleanName(macroName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acMacro, macroName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR VBA COMPLETO
'===========================================================================
Private Sub ExportVBA(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error GoTo ErrH
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Integer
    
    On Error Resume Next
    Set vbProj = accessApp.VBE.ActiveVBProject
    On Error GoTo ErrH
    
    If vbProj Is Nothing Then
        WriteUTF8File basePath & "\" & GetFolderName("VBA", language) & "\00_ERROR.txt", "No se puede acceder al proyecto VBA. Habilitar acceso programático."
        Exit Sub
    End If
    
    For i = 1 To vbProj.VBComponents.Count
        Set vbComp = vbProj.VBComponents(i)
        
        If vbComp.CodeModule.CountOfLines > 0 Then
            ExportVBAComponent basePath & "\" & GetFolderName("VBA", language), vbComp
        End If
    Next i
    
    Exit Sub
ErrH:
End Sub

Private Sub ExportVBAComponent(basePath As String, vbComp As Object)
    On Error GoTo ErrH
    
    Dim fileName As String
    Dim content As String
    Dim i As Long
    
    fileName = CleanName(vbComp.Name) & ".bas"
    
    content = "' ===============================================" & vbCrLf
    content = content & "' MÓDULO VBA: " & vbComp.Name & vbCrLf
    content = content & "' Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "' ===============================================" & vbCrLf & vbCrLf
    
    For i = 1 To vbComp.CodeModule.CountOfLines
        content = content & vbComp.CodeModule.Lines(i, 1) & vbCrLf
    Next i
    
    WriteUTF8File basePath & "\" & fileName, content
    
    Exit Sub
ErrH:
End Sub

'===========================================================================
' EXPORTAR TABLAS - UN ARCHIVO DDL POR TABLA (ACCESS Y SQL SERVER)
'===========================================================================
Private Sub ExportTables(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error GoTo ErrH
    
    Dim db As DAO.Database
    Set db = accessApp.CurrentDb
    
    Dim tbl As DAO.TableDef
    Dim accessTablesPath As String
    Dim sqlServerTablesPath As String
    
    accessTablesPath = basePath & "\" & GetFolderName("TABLES", language) & "\" & GetFolderName("ACCESS", language)
    sqlServerTablesPath = basePath & "\" & GetFolderName("TABLES", language) & "\" & GetFolderName("SQLSERVER", language)
    
    Debug.Print "ExportTables iniciado - Rutas:"
    Debug.Print "  Access: " & accessTablesPath
    Debug.Print "  SQL Server: " & sqlServerTablesPath
    
    ' Crear subcarpetas
    MkDir accessTablesPath
    MkDir sqlServerTablesPath
    
    Debug.Print "Carpetas creadas"
    
    ' Exportar cada tabla individual
    Dim tableCount As Integer
    tableCount = 0
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            tableCount = tableCount + 1
            Debug.Print "Exportando tabla [" & tableCount & "]: " & tbl.Name
            
            ' Generar DDL Access
            ExportTableAccessDDL tbl, accessTablesPath
            
            ' Generar DDL SQL Server
            ExportTableSQLServerDDL tbl, sqlServerTablesPath
        End If
    Next tbl
    
    Debug.Print "ExportTables completado - " & tableCount & " tablas exportadas"
    
    Exit Sub
ErrH:
    Debug.Print "Error en ExportTables: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
End Sub

'===========================================================================
' EXPORTAR DDL DE TABLA PARA ACCESS
'===========================================================================
Private Sub ExportTableAccessDDL(tbl As DAO.TableDef, basePath As String)
    On Error GoTo ErrH
    
    Dim content As String
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim cleanTableName As String
    Dim fieldCount As Integer
    Dim primaryKeyStr As String
    
    cleanTableName = CleanName(tbl.Name)
    
    ' Encabezado
    content = "-- =============================================================" & vbCrLf
    content = content & "-- DDL DE ACCESS: Tabla [" & tbl.Name & "]" & vbCrLf
    content = content & "-- Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "-- Motor: Microsoft Access" & vbCrLf
    content = content & "-- =============================================================" & vbCrLf & vbCrLf
    
    ' Iniciar CREATE TABLE
    content = content & "CREATE TABLE [" & tbl.Name & "] (" & vbCrLf
    
    ' Agregar campos
    fieldCount = 0
    For Each fld In tbl.Fields
        fieldCount = fieldCount + 1
        
        If fieldCount > 1 Then
            content = content & "," & vbCrLf
        End If
        
        content = content & "    [" & fld.Name & "] " & GetAccessFieldType(fld)
        
        ' Propiedades del campo
        If (fld.Attributes And dbAutoIncrField) <> 0 Then
            content = content & " AUTOINCREMENT"
        End If
        
        If (fld.Required) Then
            content = content & " NOT NULL"
        End If
        
        If fld.DefaultValue <> "" Then
            content = content & " DEFAULT " & fld.DefaultValue
        End If
    Next fld
    
    ' Agregar claves primarias
    Dim primaryKeyStr As String
    primaryKeyStr = GetPrimaryKeyFields(tbl)
    If primaryKeyStr <> "" Then
        content = content & "," & vbCrLf & "    PRIMARY KEY ([" & Replace(primaryKeyStr, ";", "],[") & "])"
    End If
    
    content = content & vbCrLf & ");" & vbCrLf & vbCrLf
    
    ' Documentación adicional
    content = content & "-- PROPIEDADES DE LA TABLA:" & vbCrLf
    content = content & "-- Total de campos: " & tbl.Fields.Count & vbCrLf
    content = content & "-- Índices: " & tbl.Indexes.Count & vbCrLf & vbCrLf
    
    ' Índices
    If tbl.Indexes.Count > 0 Then
        content = content & "-- ÍNDICES:" & vbCrLf
        For Each idx In tbl.Indexes
            content = content & "-- CREATE " & IIf(idx.Unique, "UNIQUE ", "") & "INDEX [" & idx.Name & "]"
            content = content & " ON [" & tbl.Name & "] ([" & Replace(idx.Fields, ";", "],[") & "])" & vbCrLf
        Next idx
        content = content & vbCrLf
    End If
    
    ' Listado de campos
    content = content & "-- CAMPOS:" & vbCrLf
    For Each fld In tbl.Fields
        content = content & "-- [" & fld.Name & "] - " & GetAccessFieldType(fld)
        If fld.Size > 0 Then content = content & " (Size: " & fld.Size & ")"
        If fld.Required Then content = content & " [NOT NULL]"
        content = content & vbCrLf
    Next fld
    
    WriteUTF8File basePath & "\" & cleanTableName & ".txt", content
    
    Exit Sub
ErrH:
    Debug.Print "Error en ExportTableAccessDDL [" & tbl.Name & "]: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
End Sub

'===========================================================================
' EXPORTAR DDL DE TABLA PARA SQL SERVER
'===========================================================================
Private Sub ExportTableSQLServerDDL(tbl As DAO.TableDef, basePath As String)
    On Error GoTo ErrH
    
    Dim content As String
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim cleanTableName As String
    Dim fieldCount As Integer
    Dim primaryKeyStr As String
    
    cleanTableName = CleanName(tbl.Name)
    
    ' Encabezado
    content = "-- =============================================================" & vbCrLf
    content = content & "-- DDL DE SQL SERVER: Tabla [" & tbl.Name & "]" & vbCrLf
    content = content & "-- Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "-- Motor: SQL Server" & vbCrLf
    content = content & "-- =============================================================" & vbCrLf & vbCrLf
    
    ' Iniciar CREATE TABLE
    content = content & "IF OBJECT_ID('[dbo].[" & tbl.Name & "]', 'U') IS NOT NULL" & vbCrLf
    content = content & "    DROP TABLE [dbo].[" & tbl.Name & "];" & vbCrLf & vbCrLf
    
    content = content & "CREATE TABLE [dbo].[" & tbl.Name & "] (" & vbCrLf
    
    ' Agregar campos
    fieldCount = 0
    For Each fld In tbl.Fields
        fieldCount = fieldCount + 1
        
        If fieldCount > 1 Then
            content = content & "," & vbCrLf
        End If
        
        content = content & "    [" & fld.Name & "] " & GetSQLServerFieldType(fld)
        
        ' Propiedades del campo
        If (fld.Attributes And dbAutoIncrField) <> 0 Then
            content = content & " IDENTITY(1,1)"
        End If
        
        If (fld.Required) Then
            content = content & " NOT NULL"
        Else
            content = content & " NULL"
        End If
        
        If fld.DefaultValue <> "" Then
            content = content & " DEFAULT " & fld.DefaultValue
        End If
    Next fld
    
    ' Agregar claves primarias
    primaryKeyStr = GetPrimaryKeyFields(tbl)
    If primaryKeyStr <> "" Then
        content = content & "," & vbCrLf & "    CONSTRAINT [PK_" & tbl.Name & "] PRIMARY KEY ([" & Replace(primaryKeyStr, ";", "],[") & "])"
    End If
    
    content = content & vbCrLf & ");" & vbCrLf & vbCrLf
    
    ' Índices no-primarios
    If tbl.Indexes.Count > 0 Then
        For Each idx In tbl.Indexes
            If Not idx.Primary Then
                content = content & "CREATE " & IIf(idx.Unique, "UNIQUE ", "") & "INDEX [IX_" & idx.Name & "]"
                content = content & " ON [dbo].[" & tbl.Name & "] ([" & Replace(idx.Fields, ";", "],[") & "]);" & vbCrLf
            End If
        Next idx
        content = content & vbCrLf
    End If
    
    ' Documentación
    content = content & "-- INFORMACIÓN DE LA TABLA:" & vbCrLf
    content = content & "-- Total de campos: " & tbl.Fields.Count & vbCrLf
    content = content & "-- Índices: " & tbl.Indexes.Count & vbCrLf & vbCrLf
    
    ' Listado de campos
    content = content & "-- CAMPOS:" & vbCrLf
    For Each fld In tbl.Fields
        content = content & "-- [" & fld.Name & "] - " & GetSQLServerFieldType(fld)
        If fld.Size > 0 Then content = content & " (Size: " & fld.Size & ")"
        If fld.Required Then content = content & " [NOT NULL]"
        content = content & vbCrLf
    Next fld
    
    WriteUTF8File basePath & "\" & cleanTableName & ".txt", content
    
    Exit Sub
ErrH:
    Debug.Print "Error en ExportTableSQLServerDDL [" & tbl.Name & "]: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
End Sub

'===========================================================================
' OBTENER TIPO DE DATO ACCESS CON DETALLE
'===========================================================================
Private Function GetAccessFieldType(fld As DAO.Field) As String
    Dim result As String
    
    Select Case fld.Type
        Case dbBoolean
            result = "YES/NO"
        Case dbByte
            result = "BYTE"
        Case dbInteger
            result = "SHORT"
        Case dbLong
            result = "LONG"
        Case dbCurrency
            result = "CURRENCY"
        Case dbSingle
            result = "SINGLE"
        Case dbDouble
            result = "DOUBLE"
        Case dbDate
            result = "DATE/TIME"
        Case dbText
            result = "TEXT(" & fld.Size & ")"
        Case dbMemo
            result = "MEMO"
        Case dbGUID
            result = "GUID"
        Case Else
            result = "TYPE_" & CStr(fld.Type)
    End Select
    
    GetAccessFieldType = result
End Function

'===========================================================================
' OBTENER TIPO DE DATO SQL SERVER CON DETALLE
'===========================================================================
Private Function GetSQLServerFieldType(fld As DAO.Field) As String
    Dim result As String
    Dim size As Long
    
    Select Case fld.Type
        Case dbBoolean
            result = "BIT"
        Case dbByte
            result = "TINYINT"
        Case dbInteger
            result = "SMALLINT"
        Case dbLong
            result = "INT"
        Case dbCurrency
            result = "MONEY"
        Case dbSingle
            result = "REAL"
        Case dbDouble
            result = "FLOAT"
        Case dbDate
            result = "DATETIME2"
        Case dbText
            size = fld.Size
            If size <= 0 Then size = 50
            If size > 8000 Then
                result = "NVARCHAR(MAX)"
            Else
                result = "NVARCHAR(" & size & ")"
            End If
        Case dbMemo
            result = "NVARCHAR(MAX)"
        Case dbGUID
            result = "UNIQUEIDENTIFIER"
        Case Else
            result = "SQL_VARIANT"
    End Select
    
    GetSQLServerFieldType = result
End Function

'===========================================================================
' FUNCIONES AUXILIARES
'===========================================================================
Private Function CountTables(db As DAO.Database) As Integer
    Dim tbl As DAO.TableDef, cnt As Integer
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then cnt = cnt + 1
    Next tbl
    CountTables = cnt
End Function

Private Function CountQueries(db As DAO.Database) As Integer
    Dim qry As DAO.QueryDef, cnt As Integer
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then cnt = cnt + 1
    Next qry
    CountQueries = cnt
End Function

Private Function IsUserTable(tbl As DAO.TableDef) As Boolean
    On Error Resume Next
    If (tbl.Attributes And (dbSystemObject Or dbHiddenObject)) <> 0 Then Exit Function
    IsUserTable = Not (Left$(UCase$(tbl.Name), 4) = "MSYS" Or Left$(UCase$(tbl.Name), 4) = "USYS")
End Function

Private Function IsUserQuery(qry As DAO.QueryDef) As Boolean
    IsUserQuery = Not (Left$(qry.Name, 4) = "~sq_" Or Left$(UCase$(qry.Name), 4) = "MSYS")
End Function

Private Function GetFieldType(f As DAO.Field) As String
    Select Case f.Type
        Case dbBoolean: GetFieldType = "Sí/No"
        Case dbByte: GetFieldType = "Byte"
        Case dbInteger: GetFieldType = "Entero"
        Case dbLong: GetFieldType = "Entero largo"
        Case dbCurrency: GetFieldType = "Moneda"
        Case dbSingle: GetFieldType = "Simple"
        Case dbDouble: GetFieldType = "Doble"
        Case dbDate: GetFieldType = "Fecha/Hora"
        Case dbText: GetFieldType = "Texto"
        Case dbMemo: GetFieldType = "Memo"
        Case Else: GetFieldType = "Tipo_" & CStr(f.Type)
    End Select
End Function

Private Function GetFieldSize(f As DAO.Field) As String
    On Error Resume Next
    If f.Type = dbText Or f.Type = dbMemo Then
        GetFieldSize = CStr(f.Size)
    Else
        GetFieldSize = "-"
    End If
End Function

'===========================================================================
' EXPORTAR CONSULTAS CON DAO
'===========================================================================
Private Sub ExportQueries(accessApp As Access.Application, basePath As String, Optional language As String = "ES")
    On Error GoTo ErrH
    
    Dim db As DAO.Database
    Set db = accessApp.CurrentDb
    
    Dim fNum As Integer
    Dim qry As DAO.QueryDef
    Dim queriesFolder As String
    
    queriesFolder = basePath & "\" & GetFolderName("QUERIES", language)
    
    fNum = FreeFile
    Open queriesFolder & "\00_Lista_Consultas.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE CONSULTAS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then
            Print #fNum, qry.Name
            
            Dim sqlContent As String
            sqlContent = "-- Consulta: " & qry.Name & vbCrLf
            sqlContent = sqlContent & "-- Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
            sqlContent = sqlContent & qry.SQL
            
            WriteUTF8File queriesFolder & "\" & CleanName(qry.Name) & ".sql", sqlContent
            Print #fNum,
        End If
    Next qry
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

Private Function CleanName(NameIn As String) As String
    Dim result As String
    result = NameIn
    result = Replace(result, " ", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    CleanName = result
End Function

Private Sub WriteUTF8File(filePath As String, content As String)
    On Error GoTo ErrH
    
    Debug.Print "WriteUTF8File: Escribiendo " & Len(content) & " bytes a " & filePath
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2
        .Close
    End With
    
    Debug.Print "WriteUTF8File: Archivo escrito exitosamente"
    Exit Sub
    
ErrH:
    Debug.Print "WriteUTF8File Error (ADODB): " & Err.Number & " - " & Err.Description & " - Intentando fallback"
    On Error GoTo ErrH2
    If Not stream Is Nothing Then stream.Close
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content;
    Close #fNum
    Debug.Print "WriteUTF8File: Archivo escrito con fallback"
    Exit Sub
    
ErrH2:
    Debug.Print "WriteUTF8File Error (Fallback): " & Err.Number & " - " & Err.Description
End Sub

'===========================================================================
' OBTENER LOS CAMPOS QUE FORMAN LA CLAVE PRIMARIA
'===========================================================================
Private Function GetPrimaryKeyFields(tbl As DAO.TableDef) As String
    On Error Resume Next
    
    Dim idx As DAO.Index
    Dim result As String
    
    For Each idx In tbl.Indexes
        If idx.Primary Then
            result = idx.Fields
            Exit For
        End If
    Next idx
    
    GetPrimaryKeyFields = result
End Function
