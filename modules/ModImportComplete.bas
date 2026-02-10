Attribute VB_Name = "ModImportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModImportComplete
' VERSION: 2.1.0
' AUTOR: Juanjo Luna (juanjo@luna-soft.es)
' FECHA: 2026-02-05
' PROYECTO: GitHub Copilot Access Analyzer Skill
'
' PROPÓSITO: Importar archivos exportados de vuelta a base de datos Access
'   - Soporte multiidioma (ES, EN, DE, FR, IT)
'   - Importación selectiva de objetos
'   - Manejo automático de errores con logs
'
' USO: RunCompleteImport "C:\path\to\target.accdb", "C:\export\folder", "ES"
'===========================================================================

' Función para llamar desde PowerShell con Eval
Public Function RunCompleteImport(ByVal targetDbPath As String, ByVal importFolder As String, Optional ByVal language As String = "ES") As Boolean
    On Error GoTo ErrHandler
    
    Dim logPath As String
    logPath = importFolder & "\00_LOG_IMPORTACION.txt"
    
    ' Validar idioma
    Select Case UCase(language)
        Case "ES", "EN", "DE", "FR", "IT"
            ' OK
        Case Else
            language = "EN"
    End Select
    
    ' Inicializar log
    InitLog logPath
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "INICIO DE IMPORTACION COMPLETA A ACCESS"
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    AppendLog logPath, "Base de datos destino: " & targetDbPath
    AppendLog logPath, "Carpeta de importación: " & importFolder
    AppendLog logPath, "Idioma: " & language
    AppendLog logPath, "Seleccion: completa"
    AppendLog logPath, ""
    
    Dim accessApp As Access.Application
    
    ' Validar archivo destino
    If Dir(targetDbPath) = "" Then
        Debug.Print "Archivo destino no encontrado: " & targetDbPath
        AppendLog logPath, "[ERROR] Archivo destino no encontrado: " & targetDbPath
        RunCompleteImport = False
        Exit Function
    End If
    AppendLog logPath, "[01:00] Archivo destino validado OK"
    
    ' Validar carpeta de importación
    If Dir(importFolder, vbDirectory) = "" Then
        Debug.Print "Carpeta de importación no encontrada: " & importFolder
        AppendLog logPath, "[ERROR] Carpeta de importación no encontrada: " & importFolder
        RunCompleteImport = False
        Exit Function
    End If
    AppendLog logPath, "[01:01] Carpeta de importación validada OK"
    
    ' Crear nueva instancia de Access
    AppendLog logPath, "[02:00] Abriendo base de datos destino..."
    Set accessApp = New Access.Application
    accessApp.Visible = False
    accessApp.OpenCurrentDatabase targetDbPath, False
    AppendLog logPath, "[02:01] Base de datos abierta exitosamente"
    
    ' Importar todo
    AppendLog logPath, "[03:00] Iniciando importación de objetos..."
    Call ImportarArchivos(accessApp, importFolder, language, logPath)
    
    ' Cerrar
    AppendLog logPath, "[04:00] Guardando cambios y cerrando..."
    accessApp.Quit acQuitSaveAll
    Set accessApp = Nothing
    AppendLog logPath, "[04:01] Base de datos cerrada"
    
    AppendLog logPath, ""
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "IMPORTACION COMPLETADA EXITOSAMENTE"
    AppendLog logPath, "=" & String(68, "=")
    Debug.Print "Importación completada: " & targetDbPath
    RunCompleteImport = True
    Exit Function
    
ErrHandler:
    Debug.Print "Import Error: " & Err.Number & " - " & Err.Description
    AppendLog logPath, "[ERROR] " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
    RunCompleteImport = False
End Function

' Función para llamar desde PowerShell con selección de objetos
Public Function RunSelectedImport(ByVal targetDbPath As String, ByVal importFolder As String, Optional ByVal language As String = "ES", Optional ByVal tableList As String = "", Optional ByVal queryList As String = "", Optional ByVal formList As String = "", Optional ByVal reportList As String = "", Optional ByVal macroList As String = "", Optional ByVal moduleList As String = "") As Boolean
    On Error GoTo ErrHandler
    
    Dim logPath As String
    logPath = importFolder & "\00_LOG_IMPORTACION.txt"
    
    Dim tableFilter As Object
    Dim queryFilter As Object
    Dim formFilter As Object
    Dim reportFilter As Object
    Dim macroFilter As Object
    Dim vbaFilter As Object
    
    Set tableFilter = BuildNameSet(tableList)
    Set queryFilter = BuildNameSet(queryList)
    Set formFilter = BuildNameSet(formList)
    Set reportFilter = BuildNameSet(reportList)
    Set macroFilter = BuildNameSet(macroList)
    Set vbaFilter = BuildNameSet(moduleList)
    
    ' Validar idioma
    Select Case UCase(language)
        Case "ES", "EN", "DE", "FR", "IT"
            ' OK
        Case Else
            language = "EN"
    End Select
    
    ' Inicializar log
    InitLog logPath
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "INICIO DE IMPORTACION SELECTIVA A ACCESS"
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    AppendLog logPath, "Base de datos destino: " & targetDbPath
    AppendLog logPath, "Carpeta de importación: " & importFolder
    AppendLog logPath, "Idioma: " & language
    AppendLog logPath, "Seleccion: Tablas=" & CountCsv(tableList) & ", Consultas=" & CountCsv(queryList) & ", Formularios=" & CountCsv(formList) & ", Informes=" & CountCsv(reportList) & ", Macros=" & CountCsv(macroList) & ", VBA=" & CountCsv(moduleList)
    If Len(moduleList) > 0 Then
        AppendLog logPath, "Lista VBA: " & moduleList
    End If
    If Len(tableList) > 0 Then
        AppendLog logPath, "Lista Tablas: " & tableList
    End If
    AppendLog logPath, ""
    
    Dim accessApp As Access.Application
    
    If Dir(targetDbPath) = "" Then
        AppendLog logPath, "[ERROR] Archivo destino no encontrado: " & targetDbPath
        RunSelectedImport = False
        Exit Function
    End If
    
    If Dir(importFolder, vbDirectory) = "" Then
        AppendLog logPath, "[ERROR] Carpeta de importación no encontrada: " & importFolder
        RunSelectedImport = False
        Exit Function
    End If
    
    AppendLog logPath, "[02:00] Abriendo base de datos destino..."
    Set accessApp = New Access.Application
    accessApp.Visible = False
    accessApp.OpenCurrentDatabase targetDbPath, False
    AppendLog logPath, "[02:01] Base de datos abierta exitosamente"
    
    AppendLog logPath, "[03:00] Iniciando importación de objetos (selectivo)..."
    Call ImportarArchivos(accessApp, importFolder, language, logPath, tableFilter, queryFilter, formFilter, reportFilter, macroFilter, vbaFilter)
    
    AppendLog logPath, "[04:00] Guardando cambios y cerrando..."
    accessApp.Quit acQuitSaveAll
    Set accessApp = Nothing
    AppendLog logPath, "[04:01] Base de datos cerrada"
    
    AppendLog logPath, ""
    AppendLog logPath, "=" & String(68, "=")
    AppendLog logPath, "IMPORTACION COMPLETADA EXITOSAMENTE"
    AppendLog logPath, "=" & String(68, "=")
    
    RunSelectedImport = True
    Exit Function
    
ErrHandler:
    AppendLog logPath, "[ERROR] " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
    RunSelectedImport = False
End Function

'===========================================================================
' IMPORTAR TODOS LOS ARCHIVOS
'===========================================================================
Private Sub ImportarArchivos(ByRef accessApp As Access.Application, ByVal basePath As String, Optional ByVal language As String = "ES", Optional ByVal logPath As String = "", Optional ByVal tableFilter As Object = Nothing, Optional ByVal queryFilter As Object = Nothing, Optional ByVal formFilter As Object = Nothing, Optional ByVal reportFilter As Object = Nothing, Optional ByVal macroFilter As Object = Nothing, Optional ByVal vbaFilter As Object = Nothing)
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim myFile As Object
    Dim objectName As String
    Dim objectType As String
    Dim imported As Integer
    Dim importedQueries As Integer
    Dim importedForms As Integer
    Dim importedReports As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Importar tablas (XML) - SELECTIVO
    AppendLog logPath, "[03:00-XML] Importando tablas (XML)..."
    Dim tablesImported As Integer
    tablesImported = 0
    Dim xmlTablesFolder As String
    xmlTablesFolder = basePath & "\" & GetFolderName("TABLES", language) & "\" & GetFolderName("XML", language)
    If fso.FolderExists(xmlTablesFolder) Then
        Set folder = fso.GetFolder(xmlTablesFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            ' Importar archivos .table y .tabledata usando ImportXML
            If objectType = "table" Or objectType = "tabledata" Then
                objectName = fso.GetBaseName(myFile.Name)
                
                ' Verificar si esta tabla debe importarse (filtro selectivo)
                If Not ShouldImport(objectName, tableFilter) Then
                    GoTo NextTable
                End If
                
                On Error Resume Next
                ' ImportXML maneja automáticamente estructura y datos
                accessApp.ImportXML myFile.Path
                If Err.Number = 0 Then
                    tablesImported = tablesImported + 1
                    AppendLog logPath, "  OK: " & myFile.Name
                Else
                    AppendLog logPath, "  [ERROR] Tabla " & myFile.Name & ": " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
NextTable:
            End If
        Next
    End If
    AppendLog logPath, "[03:00-XML-FIN] Tablas (XML) importadas: " & tablesImported

    ' Importar consultas
    AppendLog logPath, "[03:01] Importando consultas..."
    importedQueries = 0
    Dim queriesFolder As String
    queriesFolder = basePath & "\" & GetFolderName("QUERIES", language)
    If fso.FolderExists(queriesFolder) Then
        Set folder = fso.GetFolder(queriesFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Then
                objectName = fso.GetBaseName(myFile.Name)
                If objectName <> "00_Lista_Consultas" And Left$(objectName, 6) <> "ERROR_" Then
                    If Not ShouldImport(objectName, queryFilter) Then
                        GoTo NextQuery
                    End If
                    
                    On Error Resume Next
                    ' Intentar eliminar consulta existente
                    accessApp.DoCmd.DeleteObject acQuery, objectName
                    Err.Clear
                    
                    ' Importar consulta usando LoadFromText
                    accessApp.LoadFromText acQuery, objectName, myFile.Path
                    If Err.Number = 0 Then
                        importedQueries = importedQueries + 1
                    Else
                        AppendLog logPath, "  [ERROR] Consulta " & objectName & ": " & Err.Number & " - " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
NextQuery:
            End If
        Next
    End If
    AppendLog logPath, "[03:02] Consultas importadas: " & importedQueries
    
    ' Importar formularios
    AppendLog logPath, "[03:03] Importando formularios..."
    importedForms = 0
    Dim formsFolder As String
    formsFolder = basePath & "\" & GetFolderName("FORMS", language)
    If fso.FolderExists(formsFolder) Then
        Set folder = fso.GetFolder(formsFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "form" Then
                objectName = fso.GetBaseName(myFile.Name)
                If Not ShouldImport(objectName, formFilter) Then
                    GoTo NextForm
                End If
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acForm, objectName
                accessApp.LoadFromText acForm, objectName, myFile.Path
                If Err.Number = 0 Then
                    importedForms = importedForms + 1
                Else
                    AppendLog logPath, "  [ERROR] Formulario " & objectName & ": " & Err.Number
                    Err.Clear
                End If
                On Error GoTo 0
            End If
NextForm:
        Next
    End If
    AppendLog logPath, "[03:04] Formularios importados: " & importedForms
    
    ' Importar informes
    AppendLog logPath, "[03:05] Importando informes..."
    importedReports = 0
    Dim reportsFolder As String
    reportsFolder = basePath & "\" & GetFolderName("REPORTS", language)
    If fso.FolderExists(reportsFolder) Then
        Set folder = fso.GetFolder(reportsFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "report" Then
                objectName = fso.GetBaseName(myFile.Name)
                If Not ShouldImport(objectName, reportFilter) Then
                    GoTo NextReport
                End If
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acReport, objectName
                accessApp.LoadFromText acReport, objectName, myFile.Path
                If Err.Number = 0 Then
                    importedReports = importedReports + 1
                Else
                    AppendLog logPath, "  [ERROR] Informe " & objectName & ": " & Err.Number
                    Err.Clear
                End If
                On Error GoTo 0
            End If
NextReport:
        Next
    End If
    AppendLog logPath, "[03:06] Informes importados: " & importedReports
    
    ' Importar macros
    AppendLog logPath, "[03:07] Importando macros..."
    Dim macrosImported As Integer
    macrosImported = 0
    Dim macrosFolder As String
    macrosFolder = basePath & "\" & GetFolderName("MACROS", language)
    If fso.FolderExists(macrosFolder) Then
        Set folder = fso.GetFolder(macrosFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "mac" Then
                objectName = fso.GetBaseName(myFile.Name)
                If Not ShouldImport(objectName, macroFilter) Then
                    GoTo NextMacro
                End If
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acMacro, objectName
                accessApp.LoadFromText acMacro, objectName, myFile.Path
                If Err.Number = 0 Then
                    macrosImported = macrosImported + 1
                Else
                    AppendLog logPath, "  [ERROR] Macro " & objectName & ": " & Err.Number
                    Err.Clear
                End If
                On Error GoTo 0
            End If
NextMacro:
        Next
    End If
    AppendLog logPath, "[03:08] Macros importadas: " & macrosImported
    
    ' Importar módulos VBA
    AppendLog logPath, "[03:09] Importando módulos VBA..."
    Dim vbaImported As Integer
    vbaImported = 0
    Dim vbaFolder As String
    vbaFolder = basePath & "\" & GetFolderName("VBA", language)
    If fso.FolderExists(vbaFolder) Then
        Set folder = fso.GetFolder(vbaFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            ' Soportar .bas (módulos estándar y clases exportados con SaveAsText)
            If objectType = "bas" Then
                objectName = fso.GetBaseName(myFile.Name)
                If objectName <> "00_ERROR" And Left$(objectName, 5) <> "Form_" And Left$(objectName, 7) <> "Report_" Then
                    If Not ShouldImport(objectName, vbaFilter) Then
                        GoTo NextVba
                    End If
                    
                    On Error Resume Next
                    ' Intentar eliminar módulo existente
                    accessApp.DoCmd.DeleteObject acModule, objectName
                    Err.Clear
                    
                    ' Importar módulo usando LoadFromText
                    accessApp.LoadFromText acModule, objectName, myFile.Path
                    If Err.Number = 0 Then
                        vbaImported = vbaImported + 1
                        AppendLog logPath, "  OK: Módulo " & objectName & " (.bas)"
                    Else
                        AppendLog logPath, "  [ERROR] Módulo " & objectName & ": " & Err.Number & " - " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
NextVba:
            End If
        Next
    End If
    AppendLog logPath, "[03:10] Módulos VBA importados: " & vbaImported
    
    ' Resumen
    AppendLog logPath, ""
    AppendLog logPath, "RESUMEN DE IMPORTACION:"
    AppendLog logPath, "  Tablas (XML): " & tablesImported
    AppendLog logPath, "  Consultas: " & importedQueries
    AppendLog logPath, "  Formularios: " & importedForms
    AppendLog logPath, "  Informes: " & importedReports
    AppendLog logPath, "  Macros: " & macrosImported
    AppendLog logPath, "  Módulos VBA: " & vbaImported
    AppendLog logPath, "  Total: " & (tablesImported + importedQueries + importedForms + importedReports + macrosImported + vbaImported)
    
    Set fso = Nothing
End Sub

Private Function BuildNameSet(ByVal csv As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim trimmed As String
    trimmed = Trim$(csv)
    If Len(trimmed) = 0 Then
        Set BuildNameSet = dict
        Exit Function
    End If
    
    Dim part As Variant
    For Each part In Split(trimmed, ",")
        Dim name As String
        name = NormalizeName(Trim$(CStr(part)))
        If Len(name) > 0 Then
            dict(LCase$(name)) = True
        End If
    Next
    
    Set BuildNameSet = dict
End Function

Private Function ShouldImport(ByVal name As String, ByVal filter As Object) As Boolean
    If filter Is Nothing Then
        ShouldImport = True
        Exit Function
    End If
    If filter.Count = 0 Then
        ShouldImport = False
        Exit Function
    End If
    ShouldImport = filter.Exists(LCase$(name))
End Function

Private Function NormalizeName(ByVal name As String) As String
    Dim value As String
    value = name
    If Len(value) = 0 Then
        NormalizeName = ""
        Exit Function
    End If
    If Left$(value, 1) = "\" Or Left$(value, 1) = "/" Then
        value = Mid$(value, 2)
    End If
    If InStrRev(value, "\") > 0 Then
        value = Mid$(value, InStrRev(value, "\") + 1)
    End If
    If InStrRev(value, "/") > 0 Then
        value = Mid$(value, InStrRev(value, "/") + 1)
    End If
    NormalizeName = value
End Function

Private Function CountCsv(ByVal csv As String) As Long
    Dim trimmed As String
    trimmed = Trim$(csv)
    If Len(trimmed) = 0 Then
        CountCsv = 0
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(trimmed, ",")
    Dim i As Long
    Dim count As Long
    count = 0
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            count = count + 1
        End If
    Next
    CountCsv = count
End Function

'===========================================================================
' FUNCIONES AUXILIARES
'===========================================================================
Private Sub WriteErrorFile(ByVal filePath As String, ByVal content As String)
    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
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
        
        Case "XML"
            Select Case UCase(language)
                Case "ES": result = "XML"
                Case "EN": result = "XML"
                Case "DE": result = "XML"
                Case "FR": result = "XML"
                Case "IT": result = "XML"
                Case Else: result = "XML"
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
' FUNCIONES DE LOGGING
'===========================================================================
Private Sub InitLog(logPath As String)
    On Error Resume Next
    If Len(logPath) > 0 Then
        Kill logPath  ' Eliminar log anterior si existe
    End If
    On Error GoTo 0
End Sub

Private Sub AppendLog(logPath As String, logMessage As String)
    On Error GoTo ErrH
    
    If Len(logPath) = 0 Then Exit Sub
    
    Dim fNum As Integer
    Dim timestamp As String
    
    fNum = FreeFile
    timestamp = Format(Now, "hh:nn:ss")
    
    ' Abrir en append mode
    If Dir(logPath) <> "" Then
        Open logPath For Append As #fNum
    Else
        Open logPath For Output As #fNum
    End If
    
    Print #fNum, "[" & timestamp & "] " & logMessage
    Close #fNum
    
    ' También mostrar en Debug
    Debug.Print "[LOG] " & logMessage
    Exit Sub
    
ErrH:
    On Error GoTo 0
End Sub

'===========================================================================
' IMPORTAR CONSULTA SQL (QueryDef)
'===========================================================================
Private Sub ImportSqlQuery(ByRef accessApp As Access.Application, ByVal queryName As String, ByVal filePath As String)
    Dim db As DAO.Database
    Dim sqlText As String

    sqlText = ReadUtf8Text(filePath)
    sqlText = StripSqlComments(sqlText)

    Set db = accessApp.CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete queryName
    On Error GoTo 0
    db.CreateQueryDef queryName, sqlText
End Sub

'===========================================================================
' LECTURA UTF-8
'===========================================================================
Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadUtf8Text = stm.ReadText(-1)
    stm.Close
End Function

'===========================================================================
' LIMPIAR COMENTARIOS SQL
'===========================================================================
Private Function StripSqlComments(ByVal sqlText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim output As String

    lines = Split(sqlText, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(Trim$(lines(i)), 2) <> "--" Then
            output = output & lines(i) & vbCrLf
        End If
    Next i
    StripSqlComments = Trim$(output)
End Function
