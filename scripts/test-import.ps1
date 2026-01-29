# Test Import
param(
    [string]$DatabasePath,
    [string]$ModulePath
)

$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.OpenCurrentDatabase($DatabasePath, $true)  # exclusive

try {
    $vbProject = $access.VBE.ActiveVBProject
    Write-Host "VBA Project name: $($vbProject.Name)"
    Write-Host "Module to import: $ModulePath"
    Write-Host "Module exists: $(Test-Path $ModulePath)"
    
    Write-Host "Attempting import..."
    $result = $vbProject.VBComponents.Import($ModulePath)
    
    if ($result) {
        Write-Host "Import successful! Module: $($result.Name)"
    } else {
        Write-Host "Import returned null"
    }
    
    Write-Host "`nCurrent modules:"
    foreach ($comp in $vbProject.VBComponents) {
        Write-Host "  - $($comp.Name)"
    }
}
catch {
    Write-Error "Error: $_"
    Write-Error $_.Exception
}

$access.CloseCurrentDatabase()
$access.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
