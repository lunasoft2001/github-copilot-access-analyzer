Attribute VB_Name = "ModExportExternal"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModExportExternal
' PROPÓSITO: Exportar objetos de OTRO archivo Access (externo)
' USO: ExportExternalDatabase "C:\path\to\database.accdb", "C:\output\folder"
'===========================================================================

' Wrapper para llamar desde PowerShell con Eval (debe retornar un valor)
Public Function RunExport(ByVal sourceDbPath As String, ByVal outputFolder As String) As Boolean
    On Error GoTo ErrHandler
    ExportExternalDatabase sourceDbPath, outputFolder
    RunExport = True
    Exit Function
ErrHandler:
    RunExport = False
End Function

Public Sub ExportExternalDatabase(ByVal sourceDbPath As String, Optional ByVal outputFolder As String = "")
    On Error GoTo ErrHandler
    
    ' Si los parámetros están vacíos, intentar leerlos de TempVars (llamada desde PowerShell)
    If Len(sourceDbPath) = 0 Then
        On Error Resume Next
        sourceDbPath = Nz(Application.TempVars("TargetDB"), "")
        On Error GoTo ErrHandler
    End If
    
    If Len(outputFolder) = 0 Then
        On Error Resume Next
        outputFolder = Nz(Application.TempVars("OutputDir"), "")
        On Error GoTo ErrHandler
    End If
    
    ' Validar que el archivo existe
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
    
    ' Abrir base de datos externa
    Dim dbExternal As DAO.Database
    Set dbExternal = DBEngine.Workspaces(0).OpenDatabase(sourceDbPath, False, True) ' ReadOnly = True
    
    ' Crear estructura de carpetas
    CreateExportFolders outputFolder
    
    ' Exportar todo
    ExportSummaryExternal dbExternal, sourceDbPath, outputFolder
    ExportTablesExternal dbExternal, outputFolder
    ExportQueriesExternal dbExternal, outputFolder
    ExportFormsExternal dbExternal, sourceDbPath, outputFolder
    ExportReportsExternal dbExternal, sourceDbPath, outputFolder
    ExportMacrosExternal dbExternal, sourceDbPath, outputFolder
    ExportVBAExternal dbExternal, sourceDbPath, outputFolder
    
    ' Cerrar base de datos externa
    dbExternal.Close
    Set dbExternal = Nothing
    
    MsgBox "¡Exportación completa finalizada!" & vbCrLf & vbCrLf & _
           "Archivo: " & sourceDbPath & vbCrLf & _
           "Carpeta: " & outputFolder & vbCrLf & vbCrLf & _
           "Ahora puedes abrir esta carpeta en VS Code para trabajar.", vbInformation, "Exportación Exitosa"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error durante la exportación: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
    On Error Resume Next
    If Not dbExternal Is Nothing Then dbExternal.Close
End Sub

'===========================================================================
' CREAR ESTRUCTURA DE CARPETAS
'===========================================================================
Private Sub CreateExportFolders(basePath As String)
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
Private Sub ExportSummaryExternal(db As DAO.Database, dbPath As String, basePath As String)
    On Error GoTo ErrH
    
    Dim content As String
    content = "=============================================================" & vbCrLf
    content = content & "EXPORTACIÓN DE BASE DE DATOS ACCESS" & vbCrLf
    content = content & "=============================================================" & vbCrLf
    content = content & "Archivo: " & dbPath & vbCrLf
    content = content & "Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "Codificación: UTF-8" & vbCrLf
    content = content & "=============================================================" & vbCrLf & vbCrLf
    
    content = content & "INVENTARIO DE OBJETOS:" & vbCrLf
    content = content & "- Tablas: " & CountTablesExt(db) & vbCrLf
    content = content & "- Consultas: " & CountQueriesExt(db) & vbCrLf
    content = content & "- Formularios: " & CountFormsExt(dbPath) & vbCrLf
    content = content & "- Informes: " & CountReportsExt(dbPath) & vbCrLf
    content = content & "- Macros: " & CountMacrosExt(dbPath) & vbCrLf
    
    WriteUTF8File basePath & "\00_RESUMEN.txt", content
    
    Exit Sub
ErrH:
    On Error GoTo 0
End Sub

'===========================================================================
' EXPORTAR TABLAS
'===========================================================================
Private Sub ExportTablesExternal(db As DAO.Database, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    fNum = FreeFile
    Open basePath & "\01_Tablas\Estructura_Completa.txt" For Output As #fNum
    
    Print #fNum, "ESTRUCTURA COMPLETA DE BASE DE DATOS"
    Print #fNum, String(80, "=")
    Print #fNum,
    
    For Each tbl In db.TableDefs
        If IsUserTableExt(tbl) Then
            Print #fNum, "[TABLA] " & tbl.Name
            Print #fNum, String(50, "-")
            
            For Each fld In tbl.Fields
                Print #fNum, fld.Name & " | " & GetFieldTypeExt(fld) & _
                          " | Tamaño:" & GetFieldSizeExt(fld) & _
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
' EXPORTAR CONSULTAS
'===========================================================================
Private Sub ExportQueriesExternal(db As DAO.Database, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    Dim qry As DAO.QueryDef
    
    fNum = FreeFile
    Open basePath & "\02_Consultas\00_Lista_Consultas.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE CONSULTAS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    For Each qry In db.QueryDefs
        If IsUserQueryExt(qry) Then
            Print #fNum, qry.Name
            
            Dim sqlContent As String
            sqlContent = "-- Consulta: " & qry.Name & vbCrLf
            sqlContent = sqlContent & "-- Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
            sqlContent = sqlContent & qry.SQL
            
            WriteUTF8File basePath & "\02_Consultas\" & CleanNameExt(qry.Name) & ".sql", sqlContent
            Print #fNum,
        End If
    Next qry
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' EXPORTAR FORMULARIOS (usando SaveAsText sobre archivo externo)
'===========================================================================
Private Sub ExportFormsExternal(db As DAO.Database, dbPath As String, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    fNum = FreeFile
    Open basePath & "\03_Formularios\00_Lista_Formularios.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE FORMULARIOS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    ' Necesitamos abrir el archivo como CurrentDatabase temporalmente
    ' Esto es complejo desde un database externo
    ' Mejor opción: usar MSysObjects para listar
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32768 AND Left(Name,1) <> '~' ORDER BY Name")
    
    Do While Not rs.EOF
        Print #fNum, rs!Name
        Print #fNum,
        rs.MoveNext
    Loop
    rs.Close
    
    Print #fNum, vbCrLf & "NOTA: La exportación de definiciones de formularios requiere"
    Print #fNum, "abrir el archivo como CurrentDatabase. Use SaveAsText manualmente"
    Print #fNum, "o el módulo ModImportExportLocal dentro del archivo objetivo."
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If Not rs Is Nothing Then rs.Close
End Sub

'===========================================================================
' EXPORTAR INFORMES
'===========================================================================
Private Sub ExportReportsExternal(db As DAO.Database, dbPath As String, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    fNum = FreeFile
    Open basePath & "\04_Informes\00_Lista_Informes.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE INFORMES"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32764 AND Left(Name,1) <> '~' ORDER BY Name")
    
    Do While Not rs.EOF
        Print #fNum, rs!Name
        Print #fNum,
        rs.MoveNext
    Loop
    rs.Close
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If Not rs Is Nothing Then rs.Close
End Sub

'===========================================================================
' EXPORTAR MACROS
'===========================================================================
Private Sub ExportMacrosExternal(db As DAO.Database, dbPath As String, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    fNum = FreeFile
    Open basePath & "\05_Macros\00_Lista_Macros.txt" For Output As #fNum
    
    Print #fNum, "LISTADO DE MACROS"
    Print #fNum, String(50, "=")
    Print #fNum,
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32766 ORDER BY Name")
    
    Do While Not rs.EOF
        Print #fNum, rs!Name
        Print #fNum,
        rs.MoveNext
    Loop
    rs.Close
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If Not rs Is Nothing Then rs.Close
End Sub

'===========================================================================
' EXPORTAR VBA (requiere acceso al VBProject del archivo externo)
'===========================================================================
Private Sub ExportVBAExternal(db As DAO.Database, dbPath As String, basePath As String)
    On Error GoTo ErrH
    
    Dim fNum As Integer
    fNum = FreeFile
    Open basePath & "\06_Codigo_VBA\00_NOTA.txt" For Output As #fNum
    
    Print #fNum, "EXPORTACIÓN DE CÓDIGO VBA"
    Print #fNum, String(50, "=")
    Print #fNum,
    Print #fNum, "La exportación de código VBA desde una base de datos externa"
    Print #fNum, "requiere que el archivo sea abierto como CurrentDatabase."
    Print #fNum,
    Print #fNum, "Para exportar el código VBA:"
    Print #fNum, "1. Abre el archivo: " & dbPath
    Print #fNum, "2. Importa el módulo ModImportExportLocal.bas"
    Print #fNum, "3. Ejecuta: ExportAllVBALocal(""" & basePath & """)"
    Print #fNum,
    Print #fNum, "O usa el script PowerShell: access-export-vba-local.ps1"
    
    Close #fNum
    Exit Sub
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Sub

'===========================================================================
' FUNCIONES AUXILIARES
'===========================================================================
Private Function IsUserTableExt(tbl As DAO.TableDef) As Boolean
    On Error Resume Next
    If (tbl.Attributes And (dbSystemObject Or dbHiddenObject)) <> 0 Then Exit Function
    IsUserTableExt = Not (Left$(UCase$(tbl.Name), 4) = "MSYS" Or Left$(UCase$(tbl.Name), 4) = "USYS")
End Function

Private Function IsUserQueryExt(qry As DAO.QueryDef) As Boolean
    IsUserQueryExt = Not (Left$(qry.Name, 4) = "~sq_" Or Left$(UCase$(qry.Name), 4) = "MSYS")
End Function

Private Function CountTablesExt(db As DAO.Database) As Integer
    Dim tbl As DAO.TableDef
    Dim cnt As Integer
    For Each tbl In db.TableDefs
        If IsUserTableExt(tbl) Then cnt = cnt + 1
    Next tbl
    CountTablesExt = cnt
End Function

Private Function CountQueriesExt(db As DAO.Database) As Integer
    Dim qry As DAO.QueryDef
    Dim cnt As Integer
    For Each qry In db.QueryDefs
        If IsUserQueryExt(qry) Then cnt = cnt + 1
    Next qry
    CountQueriesExt = cnt
End Function

Private Function CountFormsExt(dbPath As String) As Integer
    On Error Resume Next
    Dim db As DAO.Database
    Set db = DBEngine.Workspaces(0).OpenDatabase(dbPath, False, True)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS cnt FROM MSysObjects WHERE Type = -32768 AND Left(Name,1) <> '~'")
    If Not rs.EOF Then CountFormsExt = rs!cnt
    rs.Close
    db.Close
End Function

Private Function CountReportsExt(dbPath As String) As Integer
    On Error Resume Next
    Dim db As DAO.Database
    Set db = DBEngine.Workspaces(0).OpenDatabase(dbPath, False, True)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS cnt FROM MSysObjects WHERE Type = -32764 AND Left(Name,1) <> '~'")
    If Not rs.EOF Then CountReportsExt = rs!cnt
    rs.Close
    db.Close
End Function

Private Function CountMacrosExt(dbPath As String) As Integer
    On Error Resume Next
    Dim db As DAO.Database
    Set db = DBEngine.Workspaces(0).OpenDatabase(dbPath, False, True)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS cnt FROM MSysObjects WHERE Type = -32766")
    If Not rs.EOF Then CountMacrosExt = rs!cnt
    rs.Close
    db.Close
End Function

Private Function GetFieldTypeExt(f As DAO.Field) As String
    Select Case f.Type
        Case dbBoolean: GetFieldTypeExt = "Sí/No"
        Case dbByte: GetFieldTypeExt = "Byte"
        Case dbInteger: GetFieldTypeExt = "Entero"
        Case dbLong: GetFieldTypeExt = "Entero largo"
        Case dbCurrency: GetFieldTypeExt = "Moneda"
        Case dbSingle: GetFieldTypeExt = "Simple"
        Case dbDouble: GetFieldTypeExt = "Doble"
        Case dbDate: GetFieldTypeExt = "Fecha/Hora"
        Case dbText: GetFieldTypeExt = "Texto"
        Case dbMemo: GetFieldTypeExt = "Memo"
        Case Else: GetFieldTypeExt = "Tipo_" & CStr(f.Type)
    End Select
End Function

Private Function GetFieldSizeExt(f As DAO.Field) As String
    On Error Resume Next
    If f.Type = dbText Or f.Type = dbMemo Then
        GetFieldSizeExt = CStr(f.Size)
    Else
        GetFieldSizeExt = "-"
    End If
End Function

Private Function CleanNameExt(NameIn As String) As String
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
    CleanNameExt = result
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
