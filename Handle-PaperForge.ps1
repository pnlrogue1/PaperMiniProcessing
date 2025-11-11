# Get the current working directory
$currentDir = Get-Location

# Define the source directory containing the zip files
$sourceDir = Join-Path $currentDir "Source"

# Ensure the source directory exists
if (-Not (Test-Path $sourceDir)) {
    Write-Error "Source directory does not exist: $sourceDir"
    exit
}

# Get all zip files in the source directory matching the pattern
$zipFiles = Get-ChildItem -Path $sourceDir -Filter "*.zip" #| Where-Object { $_.Name -match $zipPattern }

# Process each matching zip file
Write-Output "Extracting Zip Files..."
foreach ($zipFile in $zipFiles) {
    # Extract the zip file to the current working directory
    $sourcePath = $zipFile.FullName
    Expand-Archive -LiteralPath $sourcePath -DestinationPath $currentDir -Force
}

# Rename the resulting directories
$rawMiniDirectories = Get-ChildItem -Path $currentDir -Directory | Where-Object { $_.Name -match "Paperforge.+Tier\d" }

Write-Output "Renaming Directories..."

foreach ($miniDir in $rawMiniDirectories) {
    $newMiniDir = $miniDir.Name -replace "\[.+\]", ""
    $newMiniDir = $newMiniDir -replace "_Tier\d", ""
    $newMiniDir = $newMiniDir -replace "Paperforge", ""
    $newMiniDir = $newMiniDir -replace "_", " - "
    $pfId = $newMiniDir.Split(" ")[0]
    Rename-Item -LiteralPath $miniDir.Name -NewName $newMiniDir -Force
    
    # PNGs
    Move-Item -LiteralPath $(Join-Path $newMiniDir "PNGs") -Destination $(Join-Path $newMiniDir "$pfId - PNGs") -Force

    # CutFile
    $cutDir = Get-ChildItem | Where-Object { $_.Name -match "$pfId" } | Where-Object { $_.Name -match "Cutfile" }
    $newCutDir = $cutDir.Name -replace "\[.+\]", ""
    $newCutDir = $newCutDir -replace "Paperforge", ""
    $newCutDir = $newCutDir -replace "_.+", ""
    $pfId = $newCutDir
    $newCutDir = $newCutDir + " - Cutfile"
    Move-Item -LiteralPath $cutDir -Destination $(Join-Path $newMiniDir $newCutDir) -Force

    # VTT File
    $vttDir = Get-ChildItem | Where-Object { $_.Name -match "^VTT$pfId" }
    $newVttDir = $vttDir.Name.Split("_")[0].Replace("VTT", "").Trim() + " - VTT"
    Move-Item -LiteralPath $vttDir -Destination $(Join-Path $newMiniDir $newVttDir) -Force
}

Write-Host "Processing complete."