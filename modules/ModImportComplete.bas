Attribute VB_Name = "ModImportComplete"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModImportComplete
' PROPÓSITO: Importar archivos exportados de vuelta a base de datos Access
' USO: RunCompleteImport "C:\path\to\target.accdb", "C:\export\folder"
'===========================================================================

' Función para llamar desde PowerShell con Eval
Public Function RunCompleteImport(ByVal targetDbPath As String, ByVal importFolder As String, Optional ByVal language As String = "ES") As Boolean
    On Error GoTo ErrHandler
    
    ' Validar idioma
    Select Case UCase(language)
        Case "ES", "EN", "DE", "FR", "IT"
            ' OK
        Case Else
            language = "EN"
    End Select
    
    Dim accessApp As Access.Application
    
    ' Validar archivo destino
    If Dir(targetDbPath) = "" Then
        Debug.Print "Archivo destino no encontrado: " & targetDbPath
        RunCompleteImport = False
        Exit Function
    End If
    
    ' Validar carpeta de importación
    If Dir(importFolder, vbDirectory) = "" Then
        Debug.Print "Carpeta de importación no encontrada: " & importFolder
        RunCompleteImport = False
        Exit Function
    End If
    
    ' Crear nueva instancia de Access
    Set accessApp = New Access.Application
    accessApp.Visible = False
    accessApp.OpenCurrentDatabase targetDbPath, False
    
    ' Importar todo
    Call ImportarArchivos(accessApp, importFolder, language)
    
    ' Cerrar
    accessApp.Quit acQuitSaveAll
    Set accessApp = Nothing
    
    Debug.Print "Importación completada: " & targetDbPath
    RunCompleteImport = True
    Exit Function
    
ErrHandler:
    Debug.Print "Import Error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not accessApp Is Nothing Then accessApp.Quit acQuitSaveNone
    RunCompleteImport = False
End Function

'===========================================================================
' IMPORTAR TODOS LOS ARCHIVOS
'===========================================================================
Private Sub ImportarArchivos(ByRef accessApp As Access.Application, ByVal basePath As String, Optional ByVal language As String = "ES")
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim myFile As Object
    Dim objectName As String
    Dim objectType As String
    Dim imported As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Importar consultas
    Dim queriesFolder As String
    queriesFolder = basePath & "\" & GetFolderName("QUERIES", language)
    If fso.FolderExists(queriesFolder) Then
        Set folder = fso.GetFolder(queriesFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "txt" Or objectType = "sql" Then
                objectName = fso.GetBaseName(myFile.Name)
                If objectName <> "00_Lista_Consultas" Then
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
                        errFile = queriesFolder & "\ERROR_" & objectName & ".txt"
                        Call WriteErrorFile(errFile, "Error importando consulta: " & objectName & vbCrLf & _
                                           "Archivo: " & myFile.Path & vbCrLf & _
                                           "Error: " & Err.Number & " - " & Err.Description)
                        Err.Clear
                    Else
                        imported = imported + 1
                    End If
                    On Error GoTo 0
                End If
            End If
        Next
    End If
    
    ' Importar formularios
    Dim formsFolder As String
    formsFolder = basePath & "\" & GetFolderName("FORMS", language)
    If fso.FolderExists(formsFolder) Then
        Set folder = fso.GetFolder(formsFolder)
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
    Dim reportsFolder As String
    reportsFolder = basePath & "\" & GetFolderName("REPORTS", language)
    If fso.FolderExists(reportsFolder) Then
        Set folder = fso.GetFolder(reportsFolder)
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
    Dim macrosFolder As String
    macrosFolder = basePath & "\" & GetFolderName("MACROS", language)
    If fso.FolderExists(macrosFolder) Then
        Set folder = fso.GetFolder(macrosFolder)
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
    Dim vbaFolder As String
    vbaFolder = basePath & "\" & GetFolderName("VBA", language)
    If fso.FolderExists(vbaFolder) Then
        Set folder = fso.GetFolder(vbaFolder)
        For Each myFile In folder.Files
            objectType = fso.GetExtensionName(myFile.Name)
            If objectType = "bas" Then
                objectName = fso.GetBaseName(myFile.Name)
                If objectName <> "00_ERROR" Then
                    On Error Resume Next
                    accessApp.DoCmd.DeleteObject acModule, objectName
                    accessApp.LoadFromText acModule, objectName, myFile.Path
                    imported = imported + 1
                    On Error GoTo 0
                End If
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

'===========================================================================
' OBTENER NOMBRE DE CARPETA LOCALIZADO
'===========================================================================
Private Function GetFolderName(folderType As String, Optional language As String = "ES") As String
    Dim result As String
    
    Select Case UCase(folderType)
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
