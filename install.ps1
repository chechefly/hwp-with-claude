# hwp with claude - smart installer (auto-installs Python if missing)
# Messages kept in English to avoid PowerShell 5.1 encoding issues.
$ErrorActionPreference = "Continue"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=============================================="
Write-Host "   hwp with claude - Installer"
Write-Host "=============================================="
Write-Host ""

function Find-Python {
    # Prefer the 'py' launcher (never a Microsoft Store stub).
    foreach ($c in @("py", "python")) {
        $cmd = Get-Command $c -ErrorAction SilentlyContinue
        if ($cmd) {
            try {
                $v = & $c --version 2>&1
                if ($LASTEXITCODE -eq 0 -and "$v" -match "Python 3") { return $c }
            } catch {}
        }
    }
    # Known install locations (user scope)
    $cand = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending | Select-Object -First 1
    if ($cand) { return $cand.FullName }
    return $null
}

$py = Find-Python

if (-not $py) {
    Write-Host "[1/2] Python not found. Installing Python automatically..."
    $installed = $false

    # Method A: winget (built into modern Windows, no admin needed with --scope user)
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "      -> using winget..."
        winget install -e --id Python.Python.3.12 --scope user `
            --accept-package-agreements --accept-source-agreements 2>&1 | Out-Host
        $installed = $true
    }

    # Method B: fallback - download official installer and run silently
    if (-not $installed -or -not (Find-Python)) {
        Write-Host "      -> downloading Python installer from python.org..."
        try {
            $url = "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"
            $out = "$env:TEMP\python-hwp-installer.exe"
            Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
            Write-Host "      -> installing (silent)..."
            Start-Process -Wait -FilePath $out -ArgumentList `
                "/quiet","InstallAllUsers=0","PrependPath=1","Include_launcher=1","Include_pip=1"
        } catch {
            Write-Host "[!] Auto-install failed: $_"
        }
    }

    # Refresh PATH in this session so we can find the new Python
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
    $py = Find-Python
}

if (-not $py) {
    Write-Host ""
    Write-Host "[X] Could not set up Python automatically."
    Write-Host "    Please install Python from https://www.python.org/downloads/"
    Write-Host "    (check 'Add Python to PATH'), then run install.bat again."
    exit 1
}

Write-Host ""
Write-Host "[2/2] Python ready: $py"
Write-Host "      Running setup..."
Write-Host ""

& $py "$Here\install.py"
