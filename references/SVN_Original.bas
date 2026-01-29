' ===============================================
' MÓDULO VBA: SVN
' Tipo: Módulo
' Exportado: 2026-01-28 16:29:01
' ===============================================

Option Compare Database
Option Explicit

Enum MSysObjectType
    Tables_Local = 1
    Access_Object_Database = 2
    Access_Object_Container = 3
    Tables_Linked_ODBC = 4
    Queries = 5
    Tables_Linked = 6
    SubDataSheet = 8
    Constraint = 9
    Data_Access_Page = -32756
    Database_Document = -32757
    User = -32758
    Form = -32768
    Reports = -32764
    Macros = -32766
    Modules = -32761
End Enum

Public Function ExportaFuentes()
Dim sExportpath As String
Dim oAccess As Object
Dim n As String
Dim condi As String

On Error GoTo Control_Error

If Forms!SVNExportar!NuevaVersion Then
'    If MsgBox("¿Quieres eliminar todos los ficheros de la carpeta " & Forms!SVNExportar!Destino & "?", vbQuestion + vbYesNo) = vbYes Then
        On Error Resume Next
        If Right(Forms!SVNExportar!Destino, 1) = "\" Then
            Kill Forms!SVNExportar!Destino & "*.*"
        Else
            Kill Forms!SVNExportar!Destino & "\*.*"
        End If
        Err.Clear
        On Error GoTo 0
'    End If
    condi = "DELETE * FROM ChangeControl WHERE Nombre = '" & Forms!SVNExportar!Nombre & "'" & _
            " AND Fuente='" & Forms!SVNExportar!Origen & "'"
    Ejecuta_Consulta condi
End If

n = ""

If Forms!SVNExportar!Tabla_Estructura = True Then n = "1" Else n = "0"
If Forms!SVNExportar!Tabla_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Tabla_Vinculada_MDB_Estructura = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Tabla_Vinculada_MDB_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Tabla_Vinculada_ODBC_Estructura = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Tabla_Vinculada_ODBC_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Relaciones = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Querys = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Forms = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Reports = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Macros = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Modules = True Then n = n & "1" Else n = n & "0"
If Forms!SVNExportar!Referencias = True Then n = n & "1" Else n = n & "0"

'If Not TableExists(Forms!SVNExportar!Nombre & "_MSysObjects_Remoto") Then
'    DoCmd.TransferDatabase acLink, "Microsoft Access", Forms!SVNExportar!Origen, acTable, "MSysObjects", Forms!SVNExportar!Nombre & "_MSysObjects_Remoto"
'    condi = "SELECT * INTO [" & Forms!SVNExportar!Nombre & "_MSysObjects_Local" & "] FROM [" & Forms!SVNExportar!Nombre & "_MSysObjects_Remoto" & "];"
'    Ejecuta_Consulta (condi)
'End If

Set oAccess = fGetRefNoAutoexec(Forms!SVNExportar!Origen)

sExportpath = Forms!SVNExportar!Destino

exportModulesTxt oAccess, sExportpath, n

Set oAccess = Nothing

Forms!SVNExportar!Fecha = Now()

If Forms!SVNExportar.Form.Dirty Then Forms!SVNExportar.Form.Dirty = False

'Borrar_Objeto acTable, Forms!SVNExportar!Nombre & "_MSysObjects_Local"
'condi = "SELECT * INTO [" & Forms!SVNExportar!Nombre & "_MSysObjects_Local" & "] FROM [" & Forms!SVNExportar!Nombre & "_MSysObjects_Remoto" & "];"
'Ejecuta_Consulta (condi)

Fin:
Err.Clear
On Error Resume Next
Application.SysCmd acSysCmdClearStatus
DoEvents
MsgBox "Exportación finalizada", vbInformation, "INFORMACION"
Exit Function

Control_Error:
MsgBox getErr, vbCritical, "ERROR"
Err.Clear
Set oAccess = Nothing
Resume Fin
Resume

End Function

Private Function exportModulesTxt(oAcc As Object, sExportpath As String, Config As String)
Dim myObj As Object
Dim sObjectName As String
Dim i As Integer
Dim db As dao.Database
Dim Archivo As String
Dim Titulo As String
Dim condi As String

Set db = OpenDatabase(Forms!SVNExportar!Origen)

'Tablas
For Each myObj In oAcc.CurrentData.AllTables
    If Left(myObj.Name, 4) <> "MSys" Then
        Application.SysCmd acSysCmdSetStatus, "Procesando Tabla: " & myObj.Name
        If Replace(CDbl(Nz(oAcc.CurrentDb.TableDefs(myObj.FullName).Properties("LastUpdated"), 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo IN (0,262144,2097152,537919488,1048576)"), 0), ",", ".") Then
            If myObj.Attributes = 0 Or myObj.Attributes = 262144 Then 'Other Attributes: 8 Local Hidden, 262152 Local Hidden
                Application.SysCmd acSysCmdSetStatus, "Exportando Tabla: " & myObj.Name
                If Mid(Config, 1, 2) = "11" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledata", sExportpath & "\" & myObj.Name & ".table", , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledata", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledata", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                ElseIf Mid(Config, 1, 2) = "10" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, , sExportpath & "\" & myObj.Name & ".table", , , acUTF8, 32, , 1
                ElseIf Mid(Config, 1, 2) = "01" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledata", , , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledata", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledata", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                End If
            'mdb Vinculada
            ElseIf myObj.Attributes = 2097152 Then 'Other Attributes: 2097160 Link Hidden, 2359296 Link Complex, 2359304 Link Complex Hidden
                Application.SysCmd acSysCmdSetStatus, "Exportando Tabla: " & myObj.Name
                If Mid(Config, 3, 2) = "11" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledataMDB", sExportpath & "\" & myObj.Name & ".tableMDB", , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledataMDB", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledataMDB", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                ElseIf Mid(Config, 3, 2) = "10" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, , sExportpath & "\" & myObj.Name & ".tableMDB", , , acUTF8, 32, , 1
                ElseIf Mid(Config, 3, 2) = "01" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledataMDB", , , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledataMDB", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledataMDB", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                End If
            'ODBC vinculada con y sin contraseña guardada
            ElseIf myObj.Attributes = 537919488 Or myObj.Attributes = 1048576 Then 'Other Attributes: 1048584 Link ODBC Hidden, 537919496 Link ODBC Hidden
                Application.SysCmd acSysCmdSetStatus, "Exportando Tabla: " & myObj.Name
                If Mid(Config, 5, 2) = "11" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledataODBC", sExportpath & "\" & myObj.Name & ".tableODBC", , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledataODBC", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledataODBC", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                ElseIf Mid(Config, 5, 2) = "10" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, , sExportpath & "\" & myObj.Name & ".tableODBC", , , acUTF8, 32, , 1
                ElseIf Mid(Config, 5, 2) = "01" Then
                    oAcc.Application.ExportXML acExportTable, myObj.Name, sExportpath & "\" & myObj.Name & ".tabledataODBC", , , , acUTF8, 32, , 1
'                    Sanitizar_Exportacion sExportpath, myObj.Name & ".tabledataODBC", "generated=", False, -2
                    Sanitizar_Fichero sExportpath & "\" & myObj.Name & ".tabledataODBC", "generated=" & Chr(34) & "·" & Chr(34) & "|", -2, False
                End If
            End If
            DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, "0,262144,2097152,537919488,1048576", myObj.Attributes, CDbl(oAcc.CurrentDb.TableDefs(myObj.FullName).Properties("LastUpdated")), CDbl(oAcc.CurrentDb.TableDefs(myObj.FullName).Properties("DateCreated"))
        End If
    End If
DoEvents
Next

'Relaciones
If Mid(Config, 7, 1) = "1" Then
    For i = 0 To db.Relations.Count - 1
        Application.SysCmd acSysCmdSetStatus, "Exportando Relaciones" & sObjectName
        sObjectName = db.Relations(i).Name
        If Left$(sObjectName, 1) <> "[" Then
            Application.SysCmd acSysCmdSetStatus, "Exportando Relaciones" & sObjectName
            writeRelationshipToFile db.Relations(i), sExportpath & "\" & sObjectName & ".rel"
        End If
    DoEvents
    Next i
End If
'Querys (5)
If Mid(Config, 8, 1) = "1" Then
    For Each myObj In oAcc.CurrentData.AllQueries
        Application.SysCmd acSysCmdSetStatus, "Procesando Query: " & myObj.Name
        If Not (myObj.Name Like "~sq_*") Then  'Ignora las querys a medias
            If Replace(CDbl(Nz(oAcc.CurrentDb.QueryDefs(myObj.FullName).Properties("LastUpdated"), 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo=" & MSysObjectType.Queries), 0), ",", ".") Then
                Application.SysCmd acSysCmdSetStatus, "Exportando Query: " & myObj.Name
                oAcc.Application.SaveAsText acQuery, myObj.Name, sExportpath & "\" & myObj.Name & ".query"
                DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, MSysObjectType.Queries, MSysObjectType.Queries, CDbl(oAcc.CurrentDb.QueryDefs(myObj.FullName).Properties("LastUpdated")), CDbl(oAcc.CurrentDb.QueryDefs(myObj.FullName).Properties("DateCreated"))
            End If
        End If
    DoEvents
    Next
End If
'Forms (-32768)
If Mid(Config, 9, 1) = "1" Then
    For Each myObj In oAcc.CurrentProject.AllForms
        Application.SysCmd acSysCmdSetStatus, "Procesando Form: " & myObj.FullName
        If Replace(CDbl(Nz(oAcc.CurrentProject.AllForms(myObj.FullName).DateModified, 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo=" & MSysObjectType.Form), 0), ",", ".") Then
            Application.SysCmd acSysCmdSetStatus, "Exportando Form: " & myObj.FullName
            oAcc.Application.SaveAsText acForm, myObj.FullName, sExportpath & "\" & myObj.FullName & ".form"
            DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, MSysObjectType.Form, MSysObjectType.Form, CDbl(oAcc.CurrentProject.AllForms(myObj.FullName).DateModified), CDbl(oAcc.CurrentProject.AllForms(myObj.FullName).DateCreated)
'            Sanitizar_Exportacion sExportpath, myObj.FullName & ".form", "Checksum =|NoSaveCTIWhenDisabled =", False
            Sanitizar_Fichero sExportpath & "\" & myObj.FullName & ".form", "Checksum =·" & Chr(13) & Chr(10) & "|NoSaveCTIWhenDisabled =·" & Chr(13) & Chr(10) & "|", -1, True
        End If
    Next
'    If Forms!SVNExportar!NuevaVersion Then
'        Sanitizar_Exportacion sExportpath, ".form", "Checksum =|NoSaveCTIWhenDisabled =", False
'    End If
End If
'Reports (-32764)
If Mid(Config, 10, 1) = "1" Then
    For Each myObj In oAcc.CurrentProject.AllReports
        Application.SysCmd acSysCmdSetStatus, "Procesando Report: " & myObj.FullName
        If Replace(CDbl(Nz(oAcc.CurrentProject.AllReports(myObj.FullName).DateModified, 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo=" & MSysObjectType.Reports), 0), ",", ".") Then
            Application.SysCmd acSysCmdSetStatus, "Exportando Report: " & myObj.FullName
            oAcc.Application.SaveAsText acReport, myObj.FullName, sExportpath & "\" & myObj.FullName & ".report"
            DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, MSysObjectType.Reports, MSysObjectType.Reports, CDbl(oAcc.CurrentProject.AllReports(myObj.FullName).DateModified), CDbl(oAcc.CurrentProject.AllReports(myObj.FullName).DateCreated)
'            Sanitizar_Exportacion sExportpath, myObj.FullName & ".report", "Checksum =|NoSaveCTIWhenDisabled =", False
            Sanitizar_Fichero sExportpath & "\" & myObj.FullName & ".report", "Checksum =·" & Chr(13) & Chr(10) & "|NoSaveCTIWhenDisabled =·" & Chr(13) & Chr(10) & "|", -1, True
        End If
    DoEvents
    Next
'    If Forms!SVNExportar!NuevaVersion Then
'        Sanitizar_Exportacion sExportpath, "*.report", "Checksum =|NoSaveCTIWhenDisabled =", False
'    End If
End If
'Macros (-32766)
If Mid(Config, 11, 1) = "1" Then
    For Each myObj In oAcc.CurrentProject.AllMacros
        Application.SysCmd acSysCmdSetStatus, "Procesando Macro: " & myObj.FullName
        If Replace(CDbl(Nz(oAcc.CurrentProject.AllMacros(myObj.FullName).DateModified, 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo=" & MSysObjectType.Macros), 0), ",", ".") Then
            Application.SysCmd acSysCmdSetStatus, "Exportando Macro: " & myObj.FullName
            oAcc.Application.SaveAsText acMacro, myObj.FullName, sExportpath & "\" & myObj.FullName & ".mac"
            DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, MSysObjectType.Macros, MSysObjectType.Macros, CDbl(oAcc.CurrentProject.AllMacros(myObj.FullName).DateModified), CDbl(oAcc.CurrentProject.AllMacros(myObj.FullName).DateCreated)
        End If
    DoEvents
    Next
End If
'Modules (-32761)
If Mid(Config, 12, 1) = "1" Then
    For Each myObj In oAcc.CurrentProject.AllModules
        Application.SysCmd acSysCmdSetStatus, "Procesando Módulo: " & myObj.FullName
        If Replace(CDbl(Nz(oAcc.CurrentProject.AllModules(myObj.FullName).DateModified, 1)), ",", ".") <> Replace(Nz(DLookup("DblDateUpdate", "ChangeControl", "Nombre='" & Forms!SVNExportar!Nombre & "' AND Fuente='" & Forms!SVNExportar!Origen & "' AND Objeto='" & myObj.FullName & "' AND Tipo=" & MSysObjectType.Modules), 0), ",", ".") Then
            Application.SysCmd acSysCmdSetStatus, "Exportando Módulo: " & myObj.FullName
            oAcc.Application.SaveAsText acModule, myObj.FullName, sExportpath & "\" & myObj.FullName & ".bas"
            DAO_ChangeControl Forms!SVNExportar!Nombre, Forms!SVNExportar!Origen, myObj.FullName, MSysObjectType.Modules, MSysObjectType.Modules, CDbl(oAcc.CurrentProject.AllModules(myObj.FullName).DateModified), CDbl(oAcc.CurrentProject.AllModules(myObj.FullName).DateCreated)
        End If
    DoEvents
    Next
End If

Set myObj = Nothing

'Referencias
If Mid(Config, 13, 1) = "1" Then
    Application.SysCmd acSysCmdSetStatus, "Procesando Referencias"
    Archivo = sExportpath & "\Referencias.txt"
    Open Archivo For Output As #1
        Titulo = "Archivo: " & oAcc.VBE.ActiveVBProject.fileName
        Print #1, "Archivo: " & oAcc.VBE.ActiveVBProject.fileName
        Print #1, "Versión de Access: " & oAcc.Version
        Print #1, "Formato de fichero: " & oAcc.CurrentProject.FileFormat
        Print #1, "Código del proyecto: " & UCase(oAcc.VBE.ActiveVBProject.Name)
        'Print #1, "Fecha: " & Date & " Hora: " & Time
        Print #1, String(Len(Titulo), "_")
        Print #1, vbCrLf
        Print #1, "REFERENCIAS"
        Print #1, "-----------"
        'Print #1, String(Len("REFERENCIAS"), "_")
        'Print #1, vbCrLf
        For Each myObj In oAcc.VBE.ActiveVBProject.References
            Application.SysCmd acSysCmdSetStatus, "Exportando Referencia: " & Trim(myObj.Name)
            Print #1, "Nombre: " & Trim(myObj.Name) & ". Descripción: " & Trim(myObj.description) & ". Path: " & Trim(myObj.FullPath) & ". GUID: " & Trim(myObj.GUID) & ". Major: " & Trim(myObj.Major) & ". Minor: " & Trim(myObj.Minor) & "."
        DoEvents
        Next
        Print #1, vbCrLf
    Close #1
End If
Application.SysCmd acSysCmdClearStatus
DoEvents
End Function
Public Sub ImportaFuentes()

Dim sPath As String
Dim oAccess As Object
Dim n As String

On Error GoTo Control_Error

n = ""

If Forms!SVNImportar!Tabla_Estructura = True Then n = "1" Else n = "0"
If Forms!SVNImportar!Tabla_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Tabla_Vinculada_MDB_Estructura = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Tabla_Vinculada_MDB_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Tabla_Vinculada_ODBC_Estructura = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Tabla_Vinculada_ODBC_Datos = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Querys = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Forms = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Reports = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Macros = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Modules = True Then n = n & "1" Else n = n & "0"
If Forms!SVNImportar!Referencias = True Then n = n & "1" Else n = n & "0"

Set oAccess = fGetRefNoAutoexec(Forms!SVNImportar!Destino)

sPath = Forms!SVNImportar!Origen

importModulesTxt oAccess, sPath, n

Set oAccess = Nothing

If (Err <> 0) And (Err.description <> Null) Then
    MsgBox getErr, vbCritical, "ERROR"
    Err.Clear
End If

Exit Sub

Control_Error:
    'Error de automatizacion. Hay que volver a empezar
    If Err.Number = -2147467259 Or Err.Number = 2501 Then
        MsgBox "Se ha producido un error de automatización", vbInformation, "INFORMACION"
        Err.Clear
        Set oAccess = Nothing
        ImportaFuentes
    Else
        MsgBox getErr, vbCritical, "ERROR"
    End If
End Sub

Function importModulesTxt(oAcc As Object, sImportpath As String, Config As String)

    Dim myObj As Object

    If InStr(1, Left(Config, 6), "1") <> 0 Then
        For Each myObj In oAcc.Application.CurrentData.AllTables
            'If myObj.Attributes = 0 Then 'Ignora las tablas de sistema (Msys...)
            If Left(myObj.Name, 4) <> "MSys" Then
                Application.SysCmd acSysCmdSetStatus, "Borrando Tabla: " & myObj.Name
                oAcc.Application.DoCmd.DeleteObject acTable, myObj.Name
            End If
        DoEvents
        Next
    End If
    
    If Mid(Config, 7, 1) = "1" Then
        For Each myObj In oAcc.Application.CurrentData.AllQueries
            'If Not (myObj.Name Like "~sq_*") Then  'Ignora las querys a medias
                Application.SysCmd acSysCmdSetStatus, "Borrando Query: " & myObj.Name
                oAcc.Application.DoCmd.DeleteObject acQuery, myObj.Name
            'End If
        DoEvents
        Next
    End If
    
    If Mid(Config, 8, 1) = "1" Then
        For Each myObj In oAcc.Application.CurrentProject.AllForms
            Application.SysCmd acSysCmdSetStatus, "Borrando Form: " & myObj.Name
            oAcc.Application.DoCmd.DeleteObject acForm, myObj.Name
        DoEvents
        Next
    End If
    
    If Mid(Config, 9, 1) = "1" Then
        For Each myObj In oAcc.Application.CurrentProject.AllReports
            Application.SysCmd acSysCmdSetStatus, "Borrando Report: " & myObj.Name
            oAcc.Application.DoCmd.DeleteObject acReport, myObj.Name
        DoEvents
        Next
    End If
    
    If Mid(Config, 10, 1) = "1" Then
        For Each myObj In oAcc.Application.CurrentProject.AllMacros
            Application.SysCmd acSysCmdSetStatus, "Borrando Macro: " & myObj.Name
            oAcc.Application.DoCmd.DeleteObject acMacro, myObj.Name
        DoEvents
        Next
    End If
    
    If Mid(Config, 11, 1) = "1" Then
        For Each myObj In oAcc.Application.CurrentProject.AllModules
            Application.SysCmd acSysCmdSetStatus, "Borrando Módulo: " & myObj.Name
            oAcc.Application.DoCmd.DeleteObject acModule, myObj.Name
        DoEvents
        Next
    End If
    
    
    Set myObj = Nothing
    
    Dim folder
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(sImportpath)

    Dim myFile, objectname, ObjectType
    For Each myFile In folder.Files
        ObjectType = fso.GetExtensionName(myFile.Name)
        objectname = fso.GetBaseName(myFile.Name)
        
        If (ObjectType = "tabledata") Then
            If Mid(Config, 1, 2) = "10" Or Mid(Config, 1, 2) = "11" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Datos de la Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        ElseIf (ObjectType = "table") Then
            If Mid(Config, 1, 2) = "01" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        ElseIf (ObjectType = "tabledataMDB") Then
            If Mid(Config, 1, 2) = "10" Or Mid(Config, 1, 2) = "11" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Datos de la Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        ElseIf (ObjectType = "tableMDB") Then
            If Mid(Config, 1, 2) = "01" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        ElseIf (ObjectType = "tabledataODBC") Then
            If Mid(Config, 1, 2) = "10" Or Mid(Config, 1, 2) = "11" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Datos de la Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        ElseIf (ObjectType = "tableODBC") Then
            If Mid(Config, 1, 2) = "01" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Tabla: " & myFile.Path
                oAcc.Application.ImportXML myFile.Path
            End If
        
        'ElseIf (objecttype = "table") Then
        '    Application.ImportXML myFile.Path
        ElseIf (ObjectType = "query") Then
            If Mid(Config, 7, 1) = "1" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Query: " & myFile.Path
                oAcc.Application.LoadFromText acQuery, objectname, myFile.Path
            End If
        ElseIf (ObjectType = "form") Then
            If Mid(Config, 8, 1) = "1" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Form: " & myFile.Path
                oAcc.Application.LoadFromText acForm, objectname, myFile.Path
            End If
        ElseIf (ObjectType = "report") Then
            If Mid(Config, 9, 1) = "1" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Report: " & myFile.Path
                oAcc.Application.LoadFromText acReport, objectname, myFile.Path
            End If
        ElseIf (ObjectType = "mac") Then
            If Mid(Config, 10, 1) = "1" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Macro: " & myFile.Path
                oAcc.Application.LoadFromText acMacro, objectname, myFile.Path
            End If
        ElseIf (ObjectType = "bas") Then
            If Mid(Config, 11, 1) = "1" Then
                Application.SysCmd acSysCmdSetStatus, "Importando Módulo: " & myFile.Path
                oAcc.Application.LoadFromText acModule, objectname, myFile.Path
            End If
        End If
    DoEvents
    Next
    Set fso = Nothing
    Set folder = Nothing
    Set myFile = Nothing
    Set objectname = Nothing
    Set ObjectType = Nothing
Application.SysCmd acSysCmdClearStatus
DoEvents
End Function

Private Sub writeRelationshipToFile(objRelationship As Object, sOutput As String)
    On Error GoTo errHandler
    
    Dim nFile As Long, bFileOpen As Boolean
    Dim sFields As String, fld As dao.Field
    
    nFile = FreeFile
    
    Open sOutput For Output As nFile
    bFileOpen = True
    
    With objRelationship
        Print #nFile, "Table:=", .Table
        Print #nFile, "ForeignTable:=", .ForeignTable
        Print #nFile, "Attributes:=", .Attributes
  
        For Each fld In objRelationship.Fields
            sFields = sFields & fld.Name & ", "
        Next fld
  
        If Len(sFields) > 0 Then
            sFields = Left$(sFields, Len(sFields) - 2)
        End If
  
        Print #nFile, "Fields:=", sFields
    End With
    
exitHandler:
    If bFileOpen = True Then
        Close #nFile
    End If
    
    Exit Sub
    
errHandler:

    MsgBox "Error in clsObjectDump::writeRelationshipToFile() at line " & Erl() & ": " & Err.description & " (" & Err.Number & "); sOutput = " & sOutput
    
    Resume Next
    
End Sub

Public Function getErr()
    Dim strError
    strError = "From " & Err.Source & ":" & vbCrLf & _
               "    Description: " & Err.description & vbCrLf & _
               "    Error Number: " & Err.Number & vbCrLf
    getErr = strError
End Function


