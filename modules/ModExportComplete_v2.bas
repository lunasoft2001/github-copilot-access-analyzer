Attribute VB_Name = "ModExportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModExportComplete
' PROPÓSITO: Exportar COMPLETAMENTE base de datos Access externa
' USO: RunCompleteExport "C:\path\to\database.accdb", "C:\output\folder"
'===========================================================================

' Función para llamar desde PowerShell con Eval
Public Function RunCompleteExport(ByVal sourceDbPath As String, ByVal outputFolder As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim accessApp As Access.Application
    Dim db As DAO.Database
    
    ' Validar archivo
    If Dir(sourceDbPath) = "" Then
        MsgBox "Archivo no encontrado: " & sourceDbPath, vbCritical
        RunCompleteExport = False
        Exit Function
    End If
    
    ' Crear carpetas
    Call CreateExportFolders(outputFolder)
    
    ' Crear nueva instancia de Access
    Set accessApp = New Access.Application
    accessApp.Visible = False
    accessApp.OpenCurrentDatabase sourceDbPath, False
    
    Set db = accessApp.CurrentDb
    
    ' Exportar todo
    Call ExportResumen(accessApp, db, sourceDbPath, outputFolder)
    Call ExportTablas(db, outputFolder)
    Call ExportConsultas(accessApp, outputFolder)
    Call ExportFormularios(accessApp, outputFolder)
    Call ExportInformes(accessApp, outputFolder)
    Call ExportMacros(accessApp, outputFolder)
    Call ExportModulosVBA(accessApp, outputFolder)
    
    ' Cerrar
    accessApp.Quit acQuitSaveNone
    Set accessApp = Nothing
    
    MsgBox "Exportación completa finalizada en:" & vbCrLf & outputFolder, vbInformation
    RunCompleteExport = True
    Exit Function
    
ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
    RunCompleteExport = False
End Function

'===========================================================================
' CREAR ESTRUCTURA DE CARPETAS
'===========================================================================
Private Sub CreateExportFolders(ByVal basePath As String)
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
Private Sub ExportResumen(ByRef accessApp As Access.Application, ByRef db As DAO.Database, _
                          ByVal dbPath As String, ByVal basePath As String)
    On Error GoTo ErrH
    
    Dim content As String
    Dim tbl As DAO.TableDef
    Dim qry As DAO.QueryDef
    Dim cntTbl As Integer, cntQry As Integer
    
    ' Contar objetos
    For Each tbl In db.TableDefs
        If Left$(tbl.Name, 4) <> "MSys" And Left$(tbl.Name, 1) <> "~" Then
            If (tbl.Attributes And dbSystemObject) = 0 Then cntTbl = cntTbl + 1
        End If
    Next
    
    For Each qry In db.QueryDefs
        If Left$(qry.Name, 1) <> "~" And Left$(qry.Name, 4) <> "MSys" Then cntQry = cntQry + 1
    Next
    
    content = "=============================================================" & vbCrLf
    content = content & "EXPORTACIÓN COMPLETA DE ACCESS" & vbCrLf
    content = content & "=============================================================" & vbCrLf
    content = content & "Archivo: " & dbPath & vbCrLf
    content = content & "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "=============================================================" & vbCrLf & vbCrLf
    content = content & "INVENTARIO:" & vbCrLf
    content = content & "- Tablas: " & cntTbl & vbCrLf
    content = content & "- Consultas: " & cntQry & vbCrLf
    content = content & "- Formularios: " & accessApp.CurrentProject.AllForms.Count & vbCrLf
    content = content & "- Informes: " & accessApp.CurrentProject.AllReports.Count & vbCrLf
    content = content & "- Macros: " & accessApp.CurrentProject.AllMacros.Count & vbCrLf
    content = content & "- Módulos VBA: " & accessApp.CurrentProject.AllModules.Count & vbCrLf
    
    Call WriteFileUTF8(basePath & "\00_RESUMEN.txt", content)
    Exit Sub
ErrH:
End Sub

'===========================================================================
' EXPORTAR TABLAS
'===========================================================================
Private Sub ExportTablas(ByRef db As DAO.Database, ByVal basePath As String)
    On Error Resume Next
    
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim content As String
    
    content = "ESTRUCTURA DE TABLAS" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf
    
    For Each tbl In db.TableDefs
        If Left$(tbl.Name, 4) <> "MSys" And Left$(tbl.Name, 1) <> "~" Then
            If (tbl.Attributes And dbSystemObject) = 0 Then
                content = content & "[TABLA] " & tbl.Name & vbCrLf
                content = content & String(50, "-") & vbCrLf
                
                For Each fld In tbl.Fields
                    content = content & "  " & fld.Name & " | Tipo: " & Tipocampo(fld.Type) & vbCrLf
                Next
                content = content & vbCrLf
            End If
        End If
    Next
    
    Call WriteFileUTF8(basePath & "\01_Tablas\Estructura.txt", content)
End Sub

'===========================================================================
' EXPORTAR CONSULTAS
'===========================================================================
Private Sub ExportConsultas(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    Dim queryName As String
    Dim filePath As String
    
    For i = 0 To accessApp.CurrentData.AllQueries.Count - 1
        queryName = accessApp.CurrentData.AllQueries(i).Name
        
        ' Ignorar consultas temporales
        If Left$(queryName, 4) <> "~sq_" And Left$(queryName, 4) <> "MSys" Then
            filePath = basePath & "\02_Consultas\" & LimpiarNombre(queryName) & ".txt"
            
            On Error Resume Next
            accessApp.SaveAsText acQuery, queryName, filePath
            On Error GoTo 0
        End If
    Next
End Sub

'===========================================================================
' EXPORTAR FORMULARIOS
'===========================================================================
Private Sub ExportFormularios(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    Dim formName As String
    Dim filePath As String
    
    For i = 0 To accessApp.CurrentProject.AllForms.Count - 1
        formName = accessApp.CurrentProject.AllForms(i).Name
        filePath = basePath & "\03_Formularios\" & LimpiarNombre(formName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acForm, formName, filePath
        On Error GoTo 0
    Next
End Sub

'===========================================================================
' EXPORTAR INFORMES
'===========================================================================
Private Sub ExportInformes(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    Dim reportName As String
    Dim filePath As String
    
    For i = 0 To accessApp.CurrentProject.AllReports.Count - 1
        reportName = accessApp.CurrentProject.AllReports(i).Name
        filePath = basePath & "\04_Informes\" & LimpiarNombre(reportName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acReport, reportName, filePath
        On Error GoTo 0
    Next
End Sub

'===========================================================================
' EXPORTAR MACROS
'===========================================================================
Private Sub ExportMacros(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    Dim macroName As String
    Dim filePath As String
    
    For i = 0 To accessApp.CurrentProject.AllMacros.Count - 1
        macroName = accessApp.CurrentProject.AllMacros(i).Name
        filePath = basePath & "\05_Macros\" & LimpiarNombre(macroName) & ".txt"
        
        On Error Resume Next
        accessApp.SaveAsText acMacro, macroName, filePath
        On Error GoTo 0
    Next
End Sub

'===========================================================================
' EXPORTAR MÓDULOS VBA
'===========================================================================
Private Sub ExportModulosVBA(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim i As Integer
    Dim moduleName As String
    Dim filePath As String
    
    ' Exportar módulos estándar
    For i = 0 To accessApp.CurrentProject.AllModules.Count - 1
        moduleName = accessApp.CurrentProject.AllModules(i).Name
        filePath = basePath & "\06_Codigo_VBA\" & LimpiarNombre(moduleName) & ".bas"
        
        On Error Resume Next
        accessApp.SaveAsText acModule, moduleName, filePath
        On Error GoTo 0
    Next
End Sub

'===========================================================================
' FUNCIONES AUXILIARES
'===========================================================================
Private Function LimpiarNombre(ByVal nombre As String) As String
    Dim resultado As String
    resultado = nombre
    resultado = Replace(resultado, " ", "_")
    resultado = Replace(resultado, "/", "_")
    resultado = Replace(resultado, "\", "_")
    resultado = Replace(resultado, ":", "_")
    resultado = Replace(resultado, "*", "_")
    resultado = Replace(resultado, "?", "_")
    resultado = Replace(resultado, """", "_")
    resultado = Replace(resultado, "<", "_")
    resultado = Replace(resultado, ">", "_")
    resultado = Replace(resultado, "|", "_")
    LimpiarNombre = resultado
End Function

Private Function TipoCAMPO(ByVal tipoDato As Integer) As String
    Select Case tipoDato
        Case dbBoolean: TipoCAMPO = "Sí/No"
        Case dbByte: TipoCAMPO = "Byte"
        Case dbInteger: TipoCAMPO = "Entero"
        Case dbLong: TipoCAMPO = "Entero largo"
        Case dbCurrency: TipoCAMPO = "Moneda"
        Case dbSingle: TipoCAMPO = "Simple"
        Case dbDouble: TipoCAMPO = "Doble"
        Case dbDate: TipoCAMPO = "Fecha/Hora"
        Case dbText: TipoCAMPO = "Texto"
        Case dbMemo: TipoCAMPO = "Memo"
        Case Else: TipoCAMPO = "Tipo_" & CStr(tipoDato)
    End Select
End Function

Private Sub WriteFileUTF8(ByVal filePath As String, ByVal content As String)
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
    ' Fallback a archivo normal
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content;
    Close #fNum
End Sub
