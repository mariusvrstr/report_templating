

# Define the packages you want to install
$packages = @("numpy", "pandas", "docxtpl", "docx.shared") 

# Check if Python is installed
$python = Get-Command python -ErrorAction SilentlyContinue

if (-not $python) {
    Write-Output "Python not found. Installing with winget..."

    # Install Python using winget
    winget install --id "Python.Python.3" -e --source winget

    # Wait for install to finish and reload path (you may need to restart session)
    $env:PATH += ";$env:LOCALAPPDATA\Programs\Python\Python310\Scripts;$env:LOCALAPPDATA\Programs\Python\Python310\"
    $python = Get-Command python -ErrorAction SilentlyContinue

    if (-not $python) {
        Write-Error "Python installation failed or not available in PATH."
        exit 1
    }
} else {
    Write-Output "Python is already installed: $($python.Source)"
}

# Ensure pip is installed
# Check if pip is available
$pip = Get-Command pip -ErrorAction SilentlyContinue

if (-not $pip) {
    Write-Output "pip not found. Attempting to install it using ensurepip..."

    # Try to install pip via Python
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        python -m ensurepip --upgrade
    } else {
        Write-Error "Python is not installed or not in PATH. Cannot install pip."
        exit 1
    }

    # Refresh environment (for current session only)
    $env:PATH += ";$env:LOCALAPPDATA\Programs\Python\Python311\Scripts;$env:LOCALAPPDATA\Programs\Python\Python311\"
    $pip = Get-Command pip -ErrorAction SilentlyContinue
}

# Upgrade pip if available now
if ($pip) {
    Write-Output "pip is installed. Upgrading to latest version..."
    python -m pip install --upgrade pip
} else {
    Write-Error "pip installation failed or pip is still not available in PATH."
}

# Force re-install PIP
# python -m ensurepip --upgrade --default-pip


# Install the required Python packages
foreach ($pkg in $packages) {
    Write-Output "Installing Python package: $pkg"
    python312 -m pip install $pkg
}

Write-Output "Python and packages are ready."







