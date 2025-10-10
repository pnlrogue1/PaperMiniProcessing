# Get the current working directory
$currentDir = Get-Location

# Define the source directory containing the zip files
$sourceDir = Join-Path $currentDir "Source"

# Ensure the source directory exists
if (-Not (Test-Path $sourceDir)) {
    Write-Error "Source directory does not exist: $sourceDir"
    exit
}

# Get all zip files in the source directory
$zipFiles = Get-ChildItem -Path $sourceDir -Filter "*.zip"

Write-Output ""

# Rename VTT files that start with "VTT - " to end with " - VTT"
Write-Output "Renaming VTT files..."
$vttFiles = $zipFiles | Where-Object { $_.Name -match "^VTT - " }
foreach ($vttFile in $vttFiles) {
    $newName = $vttFile.Name -replace "^VTT - ", ""
    $newName = $newName -replace "\.zip$", " - VTT.zip"
    Rename-Item -LiteralPath $vttFile.FullName -NewName $newName -Force
    Write-Output "Renamed: $($vttFile.Name) -> $newName"
}

Write-Output ""

# Rename files to remove " - Tier( ?)\d" from their filenames
Write-Output "Renaming ZIP files to remove tier info..."
$zipFiles = Get-ChildItem -Path $sourceDir -Filter "*.zip" # Refresh the list
foreach ($zipFile in $zipFiles) {
    $newName = $zipFile.Name -replace " - Tier ?\d", ""
    if ($newName -ne $zipFile.Name) {
        Rename-Item -LiteralPath $zipFile.FullName -NewName $newName -Force
        Write-Output "Renamed: $($zipFile.Name) -> $newName"
    }
}

Write-Output ""

Write-Output "Renaming PDF files to remove tier info..."
$pdfFiles = Get-ChildItem -Path $sourceDir -Filter "*.pdf"
foreach ($pdfFile in $pdfFiles) {
    $newName = $pdfFile.Name -replace " - Tier ?\d", ""
    if ($newName -ne $pdfFile.Name) {
        Rename-Item -LiteralPath $pdfFile.FullName -NewName $newName -Force
        Write-Output "Renamed: $($pdfFile.Name) -> $newName"
    }
}

# Refresh the file lists after renaming
$pdfFiles = Get-ChildItem -Path $sourceDir -Filter "*.pdf"
$zipFiles = Get-ChildItem -Path $sourceDir -Filter "*.zip"

Write-Output ""

# Create directories in $currentDir based on PDF file names (without extension)
Write-Output "Creating directories based on PDF file names..."
foreach ($pdfFile in $pdfFiles) {
    $dirName = [System.IO.Path]::GetFileNameWithoutExtension($pdfFile.Name)
    $targetDir = Join-Path $currentDir $dirName
    if (-Not (Test-Path $targetDir)) {
        New-Item -Path $targetDir -ItemType Directory | Out-Null
        Write-Output "Created directory: $targetDir"
    } else {
        Write-Output "Directory already exists: $targetDir"
    }
}

Write-Output ""

Write-Output "Moving PDF files into their respective directories..."
foreach ($pdfFile in $pdfFiles) {
    $dirName = [System.IO.Path]::GetFileNameWithoutExtension($pdfFile.Name)
    $targetDir = Join-Path $currentDir $dirName
    Move-Item -Path $pdfFile.FullName -Destination "$targetDir/" -Force
    Write-Output "Moved: $($pdfFile.Name) -> $targetDir/"
}

Write-Output ""

Write-Output "Checking and extracting ZIP files with VTT or PNG suffix..."

# Regex pattern to match ZIP files ending with " - VTT.zip" or " - PNG.zip"
$pattern = ' - (VTT|PNGs)\.zip$'

$specialZipFiles = $zipFiles | Where-Object { $_.Name -match $pattern }

foreach ($zipFile in $specialZipFiles) {
    $zipPath = $zipFile.FullName
    $zipName = $zipFile.Name
    $baseName = $zipName -replace $pattern, ""

    # Load the ZIP archive
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)

    # Check if all entries are inside a directory (not in root)
    $entries = $zip.Entries
    $allInDir = $true
    $topLevelDirs = @()

    foreach ($entry in $entries) {
        $parts = $entry.FullName.Split('/')
        if ($parts.Count -gt 1 -and $parts[0]) {
            $topLevelDirs += $parts[0]
        } elseif ($entry.FullName -notmatch '/$' -and $entry.FullName -notmatch '\\$') {
            # File is in root of ZIP
            $allInDir = $false
            break
        }
    }

    $zip.Dispose()

    $targetDir = Join-Path $currentDir $baseName

    if ($allInDir -and $topLevelDirs.Count -gt 0) {
        Write-Output "Extracting $zipName to $targetDir (contains directory)..."
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $targetDir)
    } else {
        # Create a subdirectory named after the zip file (without extension)
        $subDir = Join-Path $targetDir ([System.IO.Path]::GetFileNameWithoutExtension($zipName))
        if (-Not (Test-Path $subDir)) {
            New-Item -Path $subDir -ItemType Directory | Out-Null
        }
        Write-Output "Extracting $zipName to $subDir (root files detected)..."
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $subDir)
    }

    # Delete the zip file after extraction
    Remove-Item -Path $zipPath -Force
    Write-Output "Deleted zip file: $zipName"
}

Write-Output ""
