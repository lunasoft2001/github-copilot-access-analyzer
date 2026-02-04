' ===============================================
' MÓDULO VBA: ExportTodoSimple
' Tipo: Módulo
' Exportado: 2026-01-28 14:34:16
' ===============================================

' ===============================================
' MÓDULO VBA: Exportar
' Tipo: Módulo
' Exportado: 2026-01-15 21:48:37
' ===============================================

Option Compare Database
Option Explicit

'===========================================================================
' EXPORTADOR COMPLETO AUTOCONTENIDO - VERSIÓN SIMPLIFICADA
' Exporta la aplicación completa de Access sin dependencias externas
' Uso: ExportAll
' Autor: Juanjo Luna 2025
' Traducido al inglés por: Juan Soto 2025 AccessExperts.com
' IA: Claude 4.0
'===========================================================================

Public Sub ExportAll(Optional ByVal folderPath As String = "")
    On Error GoTo ErrH
    
    Dim basePath As String
    
    ' Determinar carpeta de destino
    If Len(folderPath) = 0 Then
        basePath = CurrentProject.Path & "\Exportacion_" & Format(Now, "yyyymmdd_hhnnss")
    Else
        basePath = folderPath
    End If
    
    ' Crear estructura de carpetas
    CreateFolders basePath
    
    ' Exportar todo
    ExportSummary basePath
    ExportAllTables basePath
    ExportAllQueries basePath
    ExportAllForms basePath
    ExportAllReports basePath
    ExportAllMacros basePath
    ExportAllVBA basePath
    
    Debug.Print "Exportación completada: " & basePath
    
    Exit Sub
ErrH:
    Debug.Print "Error durante la exportación: " & err.Number & " - " & err.Description
End Sub

'===========================================================================
' CREAR ESTRUCTURA DE CARPETAS
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
' EXPORTAR RESUMEN GENERAL
'===========================================================================
Private Sub ExportSummary(basePath As String)
    On Error GoTo ErrH
    
    Dim db As dao.Database: Set db = CurrentDb
    Dim content As String
    
    ' Construir contenido en memoria
    content = "=============================================================" & vbCrLf
    content = content & "EXPORTACIÓN COMPLETA DE APLICACIÓN ACCESS" & vbCrLf
    content = content & "=============================================================" & vbCrLf
    content = content & "Aplicación: " & Nz(CurrentProject.Name, "(sin nombre)") & vbCrLf
    content = content & "Archivo: " & CurrentProject.FullName & vbCrLf
    content = content & "Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "Codificación: UTF-8 (compatible con VS Code)" & vbCrLf
    content = content & "=============================================================" & vbCrLf & vbCrLf
    
    content = content & "?? INVENTARIO DE OBJETOS:" & vbCrLf
    content = content & "- Tablas: " & CountTables(db) & vbCrLf
    content = content & "- Consultas: " & CountQueries(db) & vbCrLf
    content = content & "- Formularios: " & CurrentProject.AllForms.count & vbCrLf
    content = content & "- Informes: " & CurrentProject.AllReports.count & vbCrLf
    content = content & "- Macros: " & CurrentProject.AllMacros.count & vbCrLf
    content = content & "- Módulos VBA: " & CurrentProject.AllModules.count & vbCrLf & vbCrLf
    
    content = content & "?? ESTRUCTURA EXPORTADA:" & vbCrLf
    content = content & "    01_Tablas/          ?? Estructura completa de BD" & vbCrLf
    content = content & "    02_Consultas/       ?? Archivos .sql individuales" & vbCrLf
    content = content & "    03_Formularios/     ?? Controles + código VBA" & vbCrLf
    content = content & "    04_Informes/        ?? Código VBA de informes" & vbCrLf
    content = content & "    05_Macros/          ?? Definiciones de macros" & vbCrLf
    content = content & "    06_Codigo_VBA/      ?? Módulos .bas y .cls" & vbCrLf & vbCrLf
    
    content = content & "?? PARA USAR EN VS CODE:" & vbCrLf
    content = content & "1. Abrir carpeta: code """ & basePath & """" & vbCrLf
    content = content & "2. Revisar este archivo primero" & vbCrLf
    content = content & "3. Explorar carpetas por tipo de objeto" & vbCrLf
    content = content & "4. El código VBA está en 06_Codigo_VBA/" & vbCrLf & vbCrLf
    
    content = content & "? CARACTERES ESPECIALES MANEJADOS:" & vbCrLf
    content = content & "- Acentos (á, é, í, ó, ú) se muestran correctamente" & vbCrLf
    content = content & "- Caracteres especiales (como ñ) se muestran perfectamente" & vbCrLf
    content = content & "- No es necesario cambiar la codificación en VS Code"
    
    ' Escribir archivo con codificación UTF-8
    WriteUTF8File basePath & "\00_RESUMEN_APLICACION.txt", content
    
    Exit Sub
ErrH:
    On Error GoTo 0
End Sub

Private Function TEST()
End Function
Private Function CountQueries(db As dao.Database) As Integer
    Dim qry As dao.QueryDef
    Dim COUNTER As Integer
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then COUNTER = COUNTER + 1
    Next qry
    CountQueries = COUNTER

End Function

'===========================================================================
' EXPORTAR TODAS LAS TABLAS
'===========================================================================
Private Sub ExportAllTables(basePath As String)
    On Error GoTo ErrH

    Dim fNum As Integer
    Dim db As dao.Database: Set db = CurrentDb
    Dim tbl As dao.TableDef
    Dim fld As dao.Field
    Dim tableType As String
    Dim linkSource As String

    fNum = FreeFile
    Open basePath & "\01_Tablas\Estructura_Completa.txt" For Output As #fNum

    Print #fNum, "ESTRUCTURA COMPLETA DE BASE DE DATOS"
    Print #fNum, String(80, "=")
    Print #fNum,

    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            Print #fNum, "[TABLA] " & tbl.Name
            Print #fNum, String(50, "-")
            ' Tipo de tabla: local o vinculada
            If (tbl.Attributes And dbAttachedTable) <> 0 Or (tbl.Attributes And dbAttachedODBC) <> 0 Then
                tableType = "Vinculada"
                On Error Resume Next
                linkSource = Nz(tbl.Connect, "")
                On Error GoTo 0
            Else
                tableType = "Local"
                linkSource = ""
            End If
            Print #fNum, "Tipo: " & tableType
            If tableType = "Vinculada" Then
                Print #fNum, "   Origen: " & linkSource
            End If
            For Each fld In tbl.Fields
                Print #fNum, fld.Name & " | " & GetFieldType(fld) & _
                          " | Tamaño:" & GetFieldSize(fld) & _
                          " | Requerido:" & IIf(fld.Required, "Sí", "No")
            Next fld
            Print #fNum,
        End If
    Next tbl

    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' EXPORTAR TODAS LAS CONSULTAS
'===========================================================================
Private Sub ExportAllQueries(basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer, fSQL As Integer
    Dim db As dao.Database: Set db = CurrentDb
    Dim qry As dao.QueryDef
    
    ' Resumen de consultas
    fNum = FreeFile
    Open basePath & "\02_Consultas\00_Lista_Consultas.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE TODAS LAS CONSULTAS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then
            Print #fNum, qry.Name
            
            ' Exportar SQL individual con UTF-8
            Dim sqlContent As String
            sqlContent = "-- Consulta: " & qry.Name & vbCrLf
            sqlContent = sqlContent & "-- Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
            sqlContent = sqlContent & "-- Codificación: UTF-8" & vbCrLf & vbCrLf
            sqlContent = sqlContent & qry.SQL
            
            WriteUTF8File basePath & "\02_Consultas\" & cleanName(qry.Name) & ".sql", sqlContent
            
            Print #fNum, "   Exportado como: " & cleanName(qry.Name) & ".sql"
            Print #fNum,
        End If
    Next qry
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If fSQL <> 0 Then Close #fSQL
End Sub

'===========================================================================
' EXPORTAR TODOS LOS FORMULARIOS
'===========================================================================
Private Sub ExportAllForms(basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    Dim i As Integer
    
    fNum = FreeFile
    Open basePath & "\03_Formularios\00_Lista_Formularios.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE TODOS LOS FORMULARIOS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    For i = 0 To CurrentProject.AllForms.count - 1
        Dim formName As String
        formName = CurrentProject.AllForms(i).Name
        Print #fNum, formName
        ' Exportar código VBA si existe
        ExportFormVBACode basePath, formName
        ' Exportar definición del formulario usando SaveAsText (si es posible)
        ExportFormDefinition basePath, formName
        Print #fNum,
    Next i
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' EXPORTAR TODOS LOS INFORMES
'===========================================================================
Private Sub ExportAllReports(basePath As String)
    On Error GoTo ErrH
    Dim fNum As Integer
    Dim i As Integer
    fNum = FreeFile
    Open basePath & "\04_Informes\00_Lista_Informes.txt" For Output As #fNum
    Print #fNum, "LISTADO DE TODOS LOS INFORMES"
    Print #fNum, String(50, "=")
    Print #fNum,
    For i = 0 To CurrentProject.AllReports.count - 1
        Dim reportName As String
        reportName = CurrentProject.AllReports(i).Name
        Print #fNum, reportName
        ' Exportar código VBA si existe
        ExportReportVBACode basePath, reportName
        ' Exportar definición del informe usando SaveAsText (si es posible)
        ExportReportDefinition basePath, reportName
        Print #fNum,
    Next i
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' EXPORTAR TODOS LOS MACROS
'===========================================================================
Private Sub ExportAllMacros(basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    Dim i As Integer
    
    fNum = FreeFile
    Open basePath & "\05_Macros\00_Lista_Macros.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE TODOS LOS MACROS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    For i = 0 To CurrentProject.AllMacros.count - 1
        Dim macroName As String
        macroName = CurrentProject.AllMacros(i).Name
        Print #fNum, macroName
        Print #fNum,

        ' Exportar macro a archivo separado usando SaveAsText
        ExportMacroDefinition basePath, macroName
    Next i
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' EXPORTAR DEFINICIÓN DE MACRO (SaveAsText)
'===========================================================================
Private Sub ExportMacroDefinition(basePath As String, macroName As String)
    On Error GoTo ErrH
    Dim filePath As String
    Dim NameClean As String
    
    NameClean = cleanName(macroName)
    filePath = basePath & "\05_Macros\" & NameClean & ".macro.txt"
    If TrySaveAsText(acMacro, macroName, filePath) Then
        ' éxito
    Else
        WriteUTF8File filePath, "-- No se pudo exportar la definición (SaveAsText) para: " & macroName
    End If
    Exit Sub
ErrH:
    On Error Resume Next
End Sub

'===========================================================================
' EXPORTAR TODO EL CÓDIGO VBA
'===========================================================================
Private Sub ExportAllVBA(basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    Dim i As Integer
    Dim vbProj As Object
    Dim vbComp As Object
    
    fNum = FreeFile
    Open basePath & "\06_Codigo_VBA\00_Lista_Modulos.txt" For Output As #fNum
    
    Print #fNum, "CÓDIGO VBA COMPLETO"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    On Error GoTo ErrH
    
    If Not vbProj Is Nothing Then
        For i = 1 To vbProj.VBComponents.count
            Set vbComp = vbProj.VBComponents(i)
            
            Print #fNum, vbComp.Name & " (" & GetComponentTypeVBA(vbComp.type) & ")"
            
            ' Exportar código si tiene alguno
            If vbComp.CodeModule.CountOfLines > 0 Then
                ExportModuleCode basePath, vbComp
                Print #fNum, "   Exportado como: " & cleanName(vbComp.Name) & ".bas"
            End If
            
            Print #fNum,
        Next i
    Else
        Print #fNum, "No se pudo acceder al proyecto VBA"
        Print #fNum, "Verificar que el acceso programático a VBA esté habilitado"
    End If
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' FUNCIONES AUXILIARES PARA EXPORTAR CÓDIGO VBA
'===========================================================================
Private Sub ExportFormVBACode(basePath As String, formName As String)
    On Error Resume Next
    
    Dim vbComp As Object
    Set vbComp = Application.VBE.ActiveVBProject.VBComponents("Form_" & formName)
    
    If Not vbComp Is Nothing Then
        If vbComp.CodeModule.CountOfLines > 0 Then
            ExportModuleCode basePath & "\03_Formularios", vbComp, formName & "_codigo"
        End If
    End If
End Sub

Private Sub ExportReportVBACode(basePath As String, reportName As String)
    On Error Resume Next
    
    Dim vbComp As Object
    Set vbComp = Application.VBE.ActiveVBProject.VBComponents("Report_" & reportName)
    
    If Not vbComp Is Nothing Then
        If vbComp.CodeModule.CountOfLines > 0 Then
            ExportModuleCode basePath & "\04_Informes", vbComp, reportName & "_codigo"
        End If
    End If
End Sub

Private Sub ExportModuleCode(basePath As String, vbComp As Object, Optional customName As String = "")
    On Error GoTo ErrH
    
    Dim fileName As String
    Dim filePath As String
    Dim content As String
    Dim i As Long
    
    If Len(customName) > 0 Then
        fileName = cleanName(customName)
    Else
        fileName = cleanName(vbComp.Name)
    End If
    
    filePath = basePath & "\" & fileName & ".bas"
    
    ' Construir el contenido completo en memoria
    content = "' ===============================================" & vbCrLf
    content = content & "' MÓDULO VBA: " & vbComp.Name & vbCrLf
    content = content & "' Tipo: " & GetComponentTypeVBA(vbComp.type) & vbCrLf
    content = content & "' Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "' ===============================================" & vbCrLf & vbCrLf
    
    ' Exportar todo el código línea por línea
    For i = 1 To vbComp.CodeModule.CountOfLines
        content = content & vbComp.CodeModule.Lines(i, 1) & vbCrLf
    Next i
    
    ' Escribir archivo con codificación UTF-8
    WriteUTF8File filePath, content
    
    Exit Sub
    
ErrH:
    On Error GoTo 0
End Sub

'===========================================================================
' FUNCIÓN AUXILIAR PARA ESCRIBIR ARCHIVOS EN UTF-8
'===========================================================================
Private Sub WriteUTF8File(filePath As String, content As String)
    On Error GoTo ErrH
    
    Dim stream As Object
    
    ' Crear objeto ADODB.Stream para escribir UTF-8
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
    
    Exit Sub
    
ErrH:
    ' Si UTF-8 falla, usar método tradicional
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    
    ' Plan B: escribir con método tradicional
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content;
    Close #fNum
End Sub

'===========================================================================
' FUNCIONES AUXILIARES
'===========================================================================
Private Function IsUserTable(tbl As dao.TableDef) As Boolean
    On Error Resume Next
    If (tbl.Attributes And (dbSystemObject Or dbHiddenObject)) <> 0 Then Exit Function
    IsUserTable = Not (Left$(UCase$(tbl.Name), 4) = "MSYS" Or Left$(UCase$(tbl.Name), 4) = "USYS")
End Function

Private Function IsUserQuery(qry As dao.QueryDef) As Boolean
    IsUserQuery = Not (Left$(qry.Name, 4) = "~sq_" Or Left$(UCase$(qry.Name), 4) = "MSYS")
End Function

Private Function CountTables(db As dao.Database) As Integer
    Dim tbl As dao.TableDef, COUNTER As Integer
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then COUNTER = COUNTER + 1
    Next tbl
    CountTables = COUNTER
End Function


Private Function GetFieldType(f As dao.Field) As String
    Select Case f.type
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
        Case Else: GetFieldType = "Tipo_" & CStr(f.type)
    End Select
End Function

Private Function GetFieldSize(f As dao.Field) As String
    On Error Resume Next
    If f.type = dbText Or f.type = dbMemo Then
        GetFieldSize = CStr(f.Size)
    Else
        GetFieldSize = "-"
    End If
End Function

Private Function GetComponentTypeVBA(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeVBA = "Módulo"
        Case 2: GetComponentTypeVBA = "Clase"
        Case 3: GetComponentTypeVBA = "Formulario"
        Case 100: GetComponentTypeVBA = "Informe"
        Case Else: GetComponentTypeVBA = "Tipo_" & componentType
    End Select
End Function

Private Function cleanName(NameIn As String) As String
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
    cleanName = result
End Function

'===========================================================================
' EXPORTAR DEFINICIONES DE FORMULARIOS E INFORMES (SaveAsText)
'===========================================================================
Private Sub ExportFormDefinition(basePath As String, formName As String)
    On Error GoTo ErrH

    Dim filePath As String
    Dim NameClean As String
    NameClean = cleanName(formName)
    filePath = basePath & "\03_Formularios\" & NameClean & ".formulario.txt"

    If TrySaveAsText(acForm, formName, filePath) Then
        ' éxito
    Else
        ' si falla, crear un archivo pequeño con advertencia
        WriteUTF8File filePath, "-- No se pudo exportar la definición (SaveAsText) para: " & formName
    End If

    Exit Sub
ErrH:
    On Error Resume Next
End Sub

Private Sub ExportReportDefinition(basePath As String, reportName As String)
    On Error GoTo ErrH

    Dim filePath As String
    Dim NameClean As String
    NameClean = cleanName(reportName)
    filePath = basePath & "\04_Informes\" & NameClean & ".informe.txt"

    If TrySaveAsText(acReport, reportName, filePath) Then
        ' éxito
    Else
        WriteUTF8File filePath, "-- No se pudo exportar la definición (SaveAsText) para: " & reportName
    End If

    Exit Sub
ErrH:
    On Error Resume Next
End Sub

' Intento seguro de Application.SaveAsText (devuelve True en caso de éxito)
Private Function TrySaveAsText(objType As Long, objName As String, outputPath As String) As Boolean
    On Error GoTo ErrH

    ' El método SaveAsText requiere que la carpeta exista
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Left$(outputPath, InStrRev(outputPath, "\") - 1)) Then
        On Error Resume Next
        MkDir Left$(outputPath, InStrRev(outputPath, "\") - 1)
        On Error GoTo ErrH
    End If

    ' Usar Application.SaveAsText
    Application.SaveAsText objType, objName, outputPath
    TrySaveAsText = True
    Exit Function
ErrH:
    TrySaveAsText = False
End Function


