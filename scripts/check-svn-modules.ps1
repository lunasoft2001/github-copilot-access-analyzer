# Check modules in SVN_JL.accdb
param([string]$DatabasePath = "C:\export\svn\SVN_JL.accdb")

$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.OpenCurrentDatabase($DatabasePath, $false)

try {
    $vbProject = $access.VBE.ActiveVBProject
    Write-Host "`nModules in $DatabasePath :" -ForegroundColor Cyan
    Write-Host "=" * 60
    
    foreach ($comp in $vbProject.VBComponents) {
        $type = switch ($comp.Type) {
            1 { "Module" }
            2 { "Class" }
            3 { "Form" }
            100 { "Report" }
            default { "Type_$($comp.Type)" }
        }
        
        $lines = $comp.CodeModule.CountOfLines
        Write-Host "  $($comp.Name) - $type ($lines lines)"
    }
    
    Write-Host ""
}
catch {
    Write-Error "Cannot access VBA: $_"
}

$access.CloseCurrentDatabase()
$access.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
