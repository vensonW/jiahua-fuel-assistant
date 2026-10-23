$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location $projectRoot
try {
    python scripts/build_brand_assets.py
    if ($LASTEXITCODE -ne 0) { throw "Brand asset preparation failed." }
    python -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { throw "Regression checks failed." }
    python -m PyInstaller --noconfirm --clean fuel_invoice.spec
    if ($LASTEXITCODE -ne 0) { throw "EXE build failed." }
    $builtExe = Get-ChildItem -LiteralPath (Join-Path $projectRoot "dist") -Filter "*.exe" -File
    if ($builtExe.Count -ne 1) { throw "Expected exactly one application EXE." }
    Copy-Item -LiteralPath $builtExe.FullName -Destination $projectRoot -Force
    Write-Host "Build complete. The EXE is in the project root and dist."
}
finally {
    Pop-Location
}
