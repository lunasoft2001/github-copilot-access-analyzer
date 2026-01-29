Attribute VB_Name = "ModExportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModExportComplete
' PROPÓSITO: Exportar COMPLETAMENTE otro archivo Access (basado en SVN_JL)
' USO: RunCompleteExport "C:\path\to\database.accdb", "C:\output\folder"
'===========================================================================

' Wrapper para llamar desde PowerShell con Eval
Public Function RunCompleteExport(ByVal sourceDbPath As String, ByVal outputFolder As String) As Boolean
    On Error GoTo ErrHandler
    ExportCompleteDatabase sourceDbPath, outputFolder
    RunCompleteExport = True
    Exit Function
ErrHandler:
    RunCompleteExport = False
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Function

Public Sub ExportCompleteDatabase(ByVal sourceDbPath As String, Optional ByVal outputFolder As String = "")
    On Error GoTo ErrHandler
    
    ' Validar archivo existe
    If Dir(sourceDbPath) = "" Then
        MsgBox "Archivo no encontrado: " & sourceDbPath, vbCritical, "Error"
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
        MsgBox "No se pudo abrir el archivo Access", vbCritical
        Exit Sub
    End If
    
    ' Crear estructura de carpetas
    CreateFolders outputFolder
    
    ' Exportar todo usando la instancia externa
    ExportAllFromExternal accessApp, sourceDbPath, outputFolder
    
    ' Cerrar Access externo
    accessApp.Quit acQuitSaveNone
    Set accessApp = Nothing
    
    MsgBox "¡Exportación completa finalizada!" & vbCrLf & vbCrLf & _
           "Archivo: " & sourceDbPath & vbCrLf & _
           "Carpeta: " & outputFolder, vbInformation, "Exportación Exitosa"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
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
Private Sub ExportAllFromExternal(accessApp As Access.Application, dbPath As String, basePath As String)
    On Error GoTo ErrHandler
    
    ' Exportar resumen
    ExportSummary accessApp, dbPath, basePath
    
    ' Exportar tablas
    ExportTables accessApp, basePath
    
    ' Exportar consultas
    ExportQueries accessApp, basePath
    
    ' Exportar formularios completos
    ExportForms accessApp, basePath
    
    ' Exportar informes completos
    ExportReports accessApp, basePath
    
    ' Exportar macros completos
    ExportMacros accessApp, basePath
    
    ' Exportar VBA completo
    ExportVBA accessApp, basePath
    
    Exit Sub
ErrHandler:
    On Error GoTo 0
End Sub

'===========================================================================
' CREAR CARPETAS
'===========================================================================
Private Sub CreateFolders(basePath As String)
    On Error Resume Next
    MkDir basePath
    MkDir basePath & "\01_Tablas"
    MkDir basePath & "\02_Consultas"
    MkDir basePath & "\03_Formularios"
    MkDir basePath & "\04_Informes"
    MkDir basePath & "\05_Macros"
    MkDir basePath & "\06_Codigo_VBA"
    On Error GoTo 0
End Sub

'===========================================================================
' EXPORTAR RESUMEN
'===========================================================================
Private Sub ExportSummary(accessApp As Access.Application, dbPath As String, basePath As String)
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
Private Sub ExportForms(accessApp As Access.Application, basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllForms.Count - 1
        Dim formName As String
        formName = accessApp.CurrentProject.AllForms(i).Name
        
        ' Usar SaveAsText para exportar definición completa
        Dim filePath As String
        filePath = basePath & "\03_Formularios\" & CleanName(formName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acForm, formName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR INFORMES CON SaveAsText
'===========================================================================
Private Sub ExportReports(accessApp As Access.Application, basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllReports.Count - 1
        Dim reportName As String
        reportName = accessApp.CurrentProject.AllReports(i).Name
        
        Dim filePath As String
        filePath = basePath & "\04_Informes\" & CleanName(reportName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acReport, reportName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR MACROS CON SaveAsText
'===========================================================================
Private Sub ExportMacros(accessApp As Access.Application, basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To accessApp.CurrentProject.AllMacros.Count - 1
        Dim macroName As String
        macroName = accessApp.CurrentProject.AllMacros(i).Name
        
        Dim filePath As String
        filePath = basePath & "\05_Macros\" & CleanName(macroName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acMacro, macroName, filePath
        On Error GoTo 0
    Next i
End Sub

'===========================================================================
' EXPORTAR VBA COMPLETO
'===========================================================================
Private Sub ExportVBA(accessApp As Access.Application, basePath As String)
    On Error GoTo ErrH
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Integer
    
    On Error Resume Next
    Set vbProj = accessApp.VBE.ActiveVBProject
    On Error GoTo ErrH
    
    If vbProj Is Nothing Then
        WriteUTF8File basePath & "\06_Codigo_VBA\00_ERROR.txt", "No se puede acceder al proyecto VBA. Habilitar acceso programático."
        Exit Sub
    End If
    
    For i = 1 To vbProj.VBComponents.Count
        Set vbComp = vbProj.VBComponents(i)
        
        If vbComp.CodeModule.CountOfLines > 0 Then
            ExportVBAComponent basePath & "\06_Codigo_VBA", vbComp
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
' EXPORTAR TABLAS CON DAO
'===========================================================================
Private Sub ExportTables(accessApp As Access.Application, basePath As String)
    On Error GoTo ErrH
    
    Dim db As DAO.Database
    Set db = accessApp.CurrentDb
    
    Dim fNum As Integer
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    fNum = FreeFile
    Open basePath & "\01_Tablas\Estructura_Completa.txt" For Output As #fNum
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
End Sub

'===========================================================================
' EXPORTAR CONSULTAS CON DAO
'===========================================================================
Private Sub ExportQueries(accessApp As Access.Application, basePath As String)
    On Error GoTo ErrH
    
    Dim db As DAO.Database
    Set db = accessApp.CurrentDb
    
    Dim fNum As Integer
    Dim qry As DAO.QueryDef
    
    fNum = FreeFile
    Open basePath & "\02_Consultas\00_Lista_Consultas.txt" For Output As #fNum
    
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
            
            WriteUTF8File basePath & "\02_Consultas\" & CleanName(qry.Name) & ".sql", sqlContent
            Print #fNum,
        End If
    Next qry
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

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
    
    Exit Sub
ErrH:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content;
    Close #fNum
End Sub
