Attribute VB_Name = "ModImportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModImportComplete
' PROPÓSITO: Importar archivos exportados de vuelta a base de datos Access
' USO: RunCompleteImport "C:\path\to\target.accdb", "C:\export\folder"
'===========================================================================

' Función para llamar desde PowerShell con Eval
Public Function RunCompleteImport(ByVal targetDbPath As String, ByVal importFolder As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim accessApp As Access.Application
    
    ' Validar archivo destino
    If Dir(targetDbPath) = "" Then
        MsgBox "Archivo destino no encontrado: " & targetDbPath, vbCritical
        RunCompleteImport = False
        Exit Function
    End If
    
    ' Validar carpeta de importación
    If Dir(importFolder, vbDirectory) = "" Then
        MsgBox "Carpeta de importación no encontrada: " & importFolder, vbCritical
        RunCompleteImport = False
        Exit Function
    End If
    
    ' Crear nueva instancia de Access
    Set accessApp = New Access.Application
    accessApp.Visible = False
    accessApp.OpenCurrentDatabase targetDbPath, False
    
    ' Importar todo
    Call ImportarArchivos(accessApp, importFolder)
    
    ' Cerrar
    accessApp.Quit acQuitSaveAll
    Set accessApp = Nothing
    
    MsgBox "Importación completa finalizada", vbInformation
    RunCompleteImport = True
    Exit Function
    
ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
    RunCompleteImport = False
End Function

'===========================================================================
' IMPORTAR TODOS LOS ARCHIVOS
'===========================================================================
Private Sub ImportarArchivos(ByRef accessApp As Access.Application, ByVal basePath As String)
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim myFile As Object
    Dim objectName As String
    Dim objectType As String
    Dim imported As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Importar consultas
    If fso.FolderExists(basePath & "\02_Consultas") Then
        Set folder = fso.GetFolder(basePath & "\02_Consultas")
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Then
                objectName = fso.GetBaseName(myFile.Name)
                
                ' Intentar borrar consulta existente (ignorar errores si no existe)
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acQuery, objectName
                Err.Clear
                On Error GoTo 0
                
                ' Intentar importar consulta
                On Error Resume Next
                accessApp.LoadFromText acQuery, objectName, myFile.Path
                If Err.Number <> 0 Then
                    ' Guardar error en archivo
                    Dim errFile As String
                    errFile = basePath & "\02_Consultas\ERROR_" & objectName & ".txt"
                    Call WriteErrorFile(errFile, "Error importando consulta: " & objectName & vbCrLf & _
                                       "Archivo: " & myFile.Path & vbCrLf & _
                                       "Error: " & Err.Number & " - " & Err.Description)
                    Err.Clear
                Else
                    imported = imported + 1
                End If
                On Error GoTo 0
            End If
        Next
    End If
    
    ' Importar formularios
    If fso.FolderExists(basePath & "\03_Formularios") Then
        Set folder = fso.GetFolder(basePath & "\03_Formularios")
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "form" Then
                objectName = fso.GetBaseName(myFile.Name)
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acForm, objectName
                accessApp.LoadFromText acForm, objectName, myFile.Path
                imported = imported + 1
                On Error GoTo 0
            End If
        Next
    End If
    
    ' Importar informes
    If fso.FolderExists(basePath & "\04_Informes") Then
        Set folder = fso.GetFolder(basePath & "\04_Informes")
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "report" Then
                objectName = fso.GetBaseName(myFile.Name)
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acReport, objectName
                accessApp.LoadFromText acReport, objectName, myFile.Path
                imported = imported + 1
                On Error GoTo 0
            End If
        Next
    End If
    
    ' Importar macros
    If fso.FolderExists(basePath & "\05_Macros") Then
        Set folder = fso.GetFolder(basePath & "\05_Macros")
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "mac" Then
                objectName = fso.GetBaseName(myFile.Name)
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acMacro, objectName
                accessApp.LoadFromText acMacro, objectName, myFile.Path
                imported = imported + 1
                On Error GoTo 0
            End If
        Next
    End If
    
    ' Importar módulos VBA
    If fso.FolderExists(basePath & "\06_Codigo_VBA") Then
        Set folder = fso.GetFolder(basePath & "\06_Codigo_VBA")
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "bas" Then
                objectName = fso.GetBaseName(myFile.Name)
                On Error Resume Next
                accessApp.DoCmd.DeleteObject acModule, objectName
                accessApp.LoadFromText acModule, objectName, myFile.Path
                imported = imported + 1
                On Error GoTo 0
            End If
        Next
    End If
    
    Set fso = Nothing
End Sub

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
