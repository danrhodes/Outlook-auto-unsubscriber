# Script to compile PowerShell scripts to EXE using PS2EXE
# This will compile Find_Unsubscribe_Links-Enhanced.ps1 to an executable

Write-Host "=== PowerShell to EXE Compiler ===" -ForegroundColor Cyan
Write-Host ""

# Check if ps2exe is installed
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "PS2EXE module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
        Write-Host "PS2EXE installed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Failed to install PS2EXE: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Please install manually by running:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name ps2exe -Scope CurrentUser" -ForegroundColor White
        exit 1
    }
}

# Import the module
Import-Module ps2exe

# Get the script directory
$scriptPath = Split-Path -Parent $PSCommandPath
if (-not $scriptPath) {
    $scriptPath = Get-Location
}

# Define source and output paths
$sourceScript = Join-Path $scriptPath "Find_Unsubscribe_Links-Enhanced.ps1"
$outputExe = Join-Path $scriptPath "UnsubscribeLinksScanner.exe"

# Check if source script exists
if (-not (Test-Path $sourceScript)) {
    Write-Host "Error: Source script not found at: $sourceScript" -ForegroundColor Red
    exit 1
}

Write-Host "Source script: $sourceScript" -ForegroundColor Cyan
Write-Host "Output EXE: $outputExe" -ForegroundColor Cyan
Write-Host ""

# Compile options
Write-Host "Select compilation options:" -ForegroundColor Yellow
Write-Host "1. Basic EXE (console application)"
Write-Host "2. EXE with icon (you'll need to provide an .ico file)"
Write-Host "3. EXE with admin privileges required"
Write-Host "4. Hidden console (runs silently)"
$compileOption = Read-Host "Select option (1-4, default: 1)"

# Base compilation parameters
$params = @{
    InputFile = $sourceScript
    OutputFile = $outputExe
    NoConsole = $false
    NoOutput = $false
    NoError = $false
    Verbose = $true
}

# Apply selected options
switch ($compileOption) {
    "2" {
        $iconPath = Read-Host "Enter path to .ico file (or press Enter to skip)"
        if ($iconPath -and (Test-Path $iconPath)) {
            $params.IconFile = $iconPath
        }
    }
    "3" {
        $params.RequireAdmin = $true
        Write-Host "Note: The EXE will require administrator privileges to run." -ForegroundColor Yellow
    }
    "4" {
        $params.NoConsole = $true
        Write-Host "Note: The EXE will run without showing a console window." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Compiling PowerShell script to EXE..." -ForegroundColor Cyan

try {
    # Compile the script
    Invoke-PS2EXE @params

    if (Test-Path $outputExe) {
        Write-Host ""
        Write-Host "SUCCESS! EXE created at:" -ForegroundColor Green
        Write-Host "  $outputExe" -ForegroundColor White
        Write-Host ""
        Write-Host "File size: $([math]::Round((Get-Item $outputExe).Length / 1MB, 2)) MB" -ForegroundColor Cyan

        $runNow = Read-Host "`nWould you like to run the EXE now? (y/n)"
        if ($runNow -match '^y') {
            Start-Process $outputExe
        }
    } else {
        Write-Host "Error: EXE was not created." -ForegroundColor Red
    }
} catch {
    Write-Host "Error during compilation: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "- Antivirus software may block EXE creation" -ForegroundColor White
    Write-Host "- PowerShell execution policy restrictions" -ForegroundColor White
    Write-Host "- Insufficient permissions" -ForegroundColor White
}

Write-Host ""
Write-Host "Note: Some antivirus software may flag the EXE as suspicious." -ForegroundColor Yellow
Write-Host "This is normal for compiled PowerShell scripts. You may need to add an exception." -ForegroundColor Yellow
