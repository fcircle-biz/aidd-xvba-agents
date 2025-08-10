# VBA File Conversion Tool (PowerShell)
# Convert from UTF-8 to Shift-JIS encoding

Write-Host "========================================"
Write-Host "VBA File Conversion Tool"
Write-Host "========================================"
Write-Host ""

$sourceDir = Join-Path $PSScriptRoot "customize\vba-files"
$targetDir = Join-Path $PSScriptRoot "vba-files"

Write-Host "Source: $sourceDir"
Write-Host "Target: $targetDir"
Write-Host ""

# Check if source directory exists
if (!(Test-Path $sourceDir)) {
    Write-Host "Error: Source directory not found" -ForegroundColor Red
    Write-Host "Path: $sourceDir"
    Read-Host "Press Enter to exit"
    exit 1
}

# Create target directories
if (!(Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
if (!(Test-Path "$targetDir\Module")) { New-Item -ItemType Directory -Path "$targetDir\Module" -Force | Out-Null }
if (!(Test-Path "$targetDir\Class")) { New-Item -ItemType Directory -Path "$targetDir\Class" -Force | Out-Null }

Write-Host "Converting files to SHIFT-JIS encoding..."

try {
    $convertedFiles = @()
    
    Get-ChildItem -Path $sourceDir -Recurse -Include '*.bas','*.cls','*.frm' | ForEach-Object {
        $relativePath = $_.FullName.Replace($sourceDir, '').TrimStart('\')
        $targetPath = Join-Path $targetDir $relativePath
        $targetDirPath = Split-Path $targetPath -Parent
        
        # Create target directory if it doesn't exist
        if (!(Test-Path $targetDirPath)) {
            New-Item -ItemType Directory -Path $targetDirPath -Force | Out-Null
        }
        
        # Read file in UTF-8 and write in Shift-JIS
        $content = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($targetPath, $content, [System.Text.Encoding]::GetEncoding('Shift_JIS'))
        
        Write-Host "Converted: $relativePath"
        $convertedFiles += $relativePath
    }
    
    Write-Host ""
    Write-Host "========================================"
    Write-Host "Conversion Complete" -ForegroundColor Green
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Files converted:"
    
    $moduleFiles = Get-ChildItem "$targetDir\Module\*.bas" -ErrorAction SilentlyContinue
    $classFiles = Get-ChildItem "$targetDir\Class\*.cls" -ErrorAction SilentlyContinue
    
    if ($moduleFiles) {
        Write-Host "Module files:"
        $moduleFiles | ForEach-Object { Write-Host "  - $($_.Name)" }
    }
    
    if ($classFiles) {
        Write-Host "Class files:"
        $classFiles | ForEach-Object { Write-Host "  - $($_.Name)" }
    }
    
} catch {
    Write-Host ""
    Write-Host "Error: Conversion failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"