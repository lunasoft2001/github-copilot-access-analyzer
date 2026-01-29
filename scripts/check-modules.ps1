# Check if module exists in database
param([string]$DatabasePath)

$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.OpenCurrentDatabase($DatabasePath, $false)

try {
    $vbProject = $access.VBE.ActiveVBProject
    Write-Host "Modules in database:" -ForegroundColor Cyan
    foreach ($comp in $vbProject.VBComponents) {
        Write-Host "  - $($comp.Name) ($($comp.Type))"
    }
}
catch {
    Write-Error "Cannot access VBA: $_"
}

$access.CloseCurrentDatabase()
$access.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
