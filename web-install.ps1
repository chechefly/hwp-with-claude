# hwp with claude - one-command web installer
# Usage:  irm https://raw.githubusercontent.com/chechefly/hwp-with-claude/main/web-install.ps1 | iex
# Downloads the tool to a STABLE location, installs Python if needed, registers with Claude.
$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

Write-Host ""
Write-Host "=============================================="
Write-Host "   hwp with claude - one-command installer"
Write-Host "=============================================="
Write-Host ""

$dest = Join-Path $env:LOCALAPPDATA "hwp-with-claude"
$zip  = Join-Path $env:TEMP "hwp-with-claude.zip"
$tmp  = Join-Path $env:TEMP "hwp-with-claude-extract"
$url  = "https://github.com/chechefly/hwp-with-claude/archive/refs/heads/main.zip"

Write-Host "[1/3] Downloading..."
Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing

Write-Host "[2/3] Installing to: $dest"
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
Expand-Archive -Path $zip -DestinationPath $tmp -Force
$inner = Get-ChildItem $tmp -Directory | Select-Object -First 1
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Move-Item $inner.FullName $dest
Remove-Item $zip -Force -ErrorAction SilentlyContinue

Write-Host "[3/3] Setting up (Python + register)..."
Write-Host ""
& "$dest\install.ps1"
