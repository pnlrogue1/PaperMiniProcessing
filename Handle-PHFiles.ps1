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

