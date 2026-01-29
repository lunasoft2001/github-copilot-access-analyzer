Attribute VB_Name = "ModExportVBALocal"
Option Compare Database
Option Explicit

'===========================================================================
' MÓDULO: ModExportVBALocal
' PROPÓSITO: Exportar código VBA cuando este archivo ES el CurrentDatabase
' USO: Importar este módulo al archivo objetivo y ejecutar ExportAllVBALocal
'===========================================================================

Public Sub ExportAllVBALocal(Optional ByVal outputFolder As String = "")
    On Error GoTo ErrHandler
    
    ' Determinar carpeta de salida
    If Len(outputFolder) = 0 Then
        outputFolder = CurrentProject.Path & "\VBA_Export_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    
    ' Crear carpeta si no existe
    On Error Resume Next
    MkDir outputFolder
    On Error GoTo ErrHandler
    
    Dim fNum As Integer
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Integer
    
    fNum = FreeFile
    Open outputFolder & "\00_Lista_Modulos.txt" For Output As #fNum
    
    Print #fNum, "CÓDIGO VBA COMPLETO"
    Print #fNum, String(50, "=")
    Print #fNum, "Exportado desde: " & CurrentProject.FullName
    Print #fNum, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fNum,
    
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    On Error GoTo ErrHandler
    
    If Not vbProj Is Nothing Then
        For i = 1 To vbProj.VBComponents.Count
            Set vbComp = vbProj.VBComponents(i)
            
            Print #fNum, vbComp.Name & " (" & GetComponentTypeNameLocal(vbComp.Type) & ")"
            
            If vbComp.CodeModule.CountOfLines > 0 Then
                ExportVBAComponent outputFolder, vbComp
                Print #fNum, "   Exportado como: " & CleanNameLocal(vbComp.Name) & ".bas"
            End If
            
            Print #fNum,
        Next i
    Else
        Print #fNum, "No se pudo acceder al proyecto VBA"
    End If
    
    Close #fNum
    
    MsgBox "Código VBA exportado a:" & vbCrLf & outputFolder, vbInformation, "Exportación VBA Completa"
    
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub ExportVBAComponent(basePath As String, vbComp As Object)
    On Error GoTo ErrH
    
    Dim fileName As String
    Dim filePath As String
    Dim content As String
    Dim i As Long
    
    fileName = CleanNameLocal(vbComp.Name)
    filePath = basePath & "\" & fileName & ".bas"
    
    content = "' ===============================================" & vbCrLf
    content = content & "' MÓDULO VBA: " & vbComp.Name & vbCrLf
    content = content & "' Tipo: " & GetComponentTypeNameLocal(vbComp.Type) & vbCrLf
    content = content & "' Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    content = content & "' ===============================================" & vbCrLf & vbCrLf
    
    For i = 1 To vbComp.CodeModule.CountOfLines
        content = content & vbComp.CodeModule.Lines(i, 1) & vbCrLf
    Next i
    
    WriteUTF8FileLocal filePath, content
    
    Exit Sub
ErrH:
    On Error GoTo 0
End Sub

Private Function GetComponentTypeNameLocal(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeNameLocal = "Módulo"
        Case 2: GetComponentTypeNameLocal = "Clase"
        Case 3: GetComponentTypeNameLocal = "Formulario"
        Case 100: GetComponentTypeNameLocal = "Informe"
        Case Else: GetComponentTypeNameLocal = "Tipo_" & componentType
    End Select
End Function

Private Function CleanNameLocal(NameIn As String) As String
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
    CleanNameLocal = result
End Function

Private Sub WriteUTF8FileLocal(filePath As String, content As String)
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
