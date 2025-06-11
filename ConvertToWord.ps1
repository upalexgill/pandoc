# Script to convert Markdown to Word document
# Requires pandoc to be installed

param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to the Markdown file to convert")]
    [string]$MarkdownFile,

    [Parameter(Mandatory = $false, Position = 1, HelpMessage = "Optional output Word file path (auto-generated if not specified)")]
    [string]$WordFile = ""
)


# Check if pandoc is installed
try {
    $null = Get-Command pandoc -ErrorAction Stop
    Write-Host "Pandoc found. Ready to convert document." -ForegroundColor Green
}
catch {
    Write-Host "Pandoc not found. Installing Pandoc via winget..." -ForegroundColor Yellow
    
    # Try to use winget to install pandoc
    try {
        Write-Host "Attempting to install Pandoc via winget..."
        winget install --id JohnMacFarlane.Pandoc --accept-package-agreements --accept-source-agreements
        Write-Host "Pandoc installed successfully." -ForegroundColor Green
        
        # Refresh the PATH environment variable to include the new Pandoc installation
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
          # Double check if we can find pandoc now
        try {
            $null = Get-Command pandoc -ErrorAction Stop
        } catch {
            Write-Host "Pandoc was installed but cannot be found in PATH." -ForegroundColor Yellow
            Write-Host "You may need to restart your terminal or computer." -ForegroundColor Yellow
            
            # Try with full path to pandoc
            $pandocPath = "C:\Program Files\Pandoc\pandoc.exe"
            if (Test-Path $pandocPath) {
                Write-Host "Found Pandoc at $pandocPath" -ForegroundColor Green
                $global:usePandocFullPath = $true
            } else {
                Write-Host "Cannot locate Pandoc executable. Manual restart may be required." -ForegroundColor Red
                exit 1
            }
        }
    }
    catch {
        Write-Host "Failed to install Pandoc via winget. Manual installation required." -ForegroundColor Red
        Write-Host "Please install Pandoc from https://pandoc.org/installing.html" -ForegroundColor Red
        Write-Host "After installing, run this script again." -ForegroundColor Red
        exit 1
    }
}

# If output file not specified, use same name as input but with .docx extension
if ([string]::IsNullOrWhiteSpace($WordFile)) {
    $WordFile = [System.IO.Path]::ChangeExtension($MarkdownFile, ".docx")
}

# Check if markdown file exists
if (-not (Test-Path $MarkdownFile)) {
    Write-Host "Markdown file not found at: $MarkdownFile" -ForegroundColor Red
    exit 1
}

# Convert the Markdown file to Word
try {
    Write-Host "Converting Markdown to Word document..." -ForegroundColor Cyan
      # Get absolute paths to avoid permission issues
    $MarkdownFileAbsolute = if ([System.IO.Path]::IsPathRooted($MarkdownFile)) {
        $MarkdownFile
    } else {
        Join-Path (Get-Location) $MarkdownFile
    }    # Verify the file exists with absolute path
    if (-not (Test-Path $MarkdownFileAbsolute)) {
        Write-Host "Cannot find markdown file at: $MarkdownFileAbsolute" -ForegroundColor Red
        exit 1
    }
    
    # Ensure output directory exists and get absolute path
    $OutputDir = Split-Path $WordFile -Parent
    if (-not [string]::IsNullOrWhiteSpace($OutputDir) -and -not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
      $WordFileAbsolute = if ([System.IO.Path]::IsPathRooted($WordFile)) {
        $WordFile
    } else {
        Join-Path (Get-Location) $WordFile
    }
    
    Write-Host "Input file: $MarkdownFileAbsolute" -ForegroundColor Gray
    Write-Host "Output file: $WordFileAbsolute" -ForegroundColor Gray
    
    # Use full path to pandoc if needed
    if ($global:usePandocFullPath) {
        $pandocExe = "C:\Program Files\Pandoc\pandoc.exe"
    } else {
        $pandocExe = "pandoc"
    }    # First try with reference document
    $referenceDoc = Join-Path (Get-Location) "docs\reference.docx"
    if (Test-Path $referenceDoc) {
        Write-Host "Using reference document for styling..." -ForegroundColor Yellow
        $conversionResult = & $pandocExe $MarkdownFileAbsolute -o $WordFileAbsolute --reference-doc=$referenceDoc 2>&1
    } else {
        Write-Host "Converting with default styling..." -ForegroundColor Yellow
        $conversionResult = & $pandocExe $MarkdownFileAbsolute -o $WordFileAbsolute 2>&1
    }
      # Check if conversion was successful
    if ($LASTEXITCODE -eq 0 -and (Test-Path $WordFileAbsolute)) {
        Write-Host "Conversion complete! Word document created at: $WordFileAbsolute" -ForegroundColor Green
        # Open the file
        try {
            Invoke-Item $WordFileAbsolute
        }
        catch {
            Write-Host "Could not open file automatically. Please open manually: $WordFileAbsolute" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Conversion failed. Error details:" -ForegroundColor Red
        Write-Host $conversionResult -ForegroundColor Red
    }
} catch {
    Write-Host "Error during conversion: $_" -ForegroundColor Red
}