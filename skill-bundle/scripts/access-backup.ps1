# ============================================================================
# Access Database Backup Script
# Creates timestamped backup of Access database file
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$false)]
    [string]$BackupFolder = ""
)

# Validate database file exists
if (-not (Test-Path $DatabasePath)) {
    Write-Error "Database file not found: $DatabasePath"
    exit 1
}

# Get file info
$dbFile = Get-Item $DatabasePath
$dbFolder = $dbFile.DirectoryName
$dbBaseName = $dbFile.BaseName
$dbExtension = $dbFile.Extension

# Determine backup folder
if ([string]::IsNullOrEmpty($BackupFolder)) {
    $BackupFolder = $dbFolder
}

# Create backup folder if needed
if (-not (Test-Path $BackupFolder)) {
    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
}

# Generate timestamped backup filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFileName = "${dbBaseName}_BACKUP_${timestamp}${dbExtension}"
$backupPath = Join-Path $BackupFolder $backupFileName

try {
    # Copy database file
    Copy-Item -Path $DatabasePath -Destination $backupPath -Force
    
    Write-Host "✓ Backup created successfully" -ForegroundColor Green
    Write-Host "  Original: $DatabasePath"
    Write-Host "  Backup:   $backupPath"
    Write-Host "  Size:     $([math]::Round($dbFile.Length / 1MB, 2)) MB"
    
    # Return backup path for programmatic use
    return $backupPath
}
catch {
    Write-Error "Failed to create backup: $_"
    exit 1
}
