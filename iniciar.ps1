$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$pythonProyecto = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $pythonProyecto)) {
    Write-Error 'Primero cree el entorno e instale requirements.txt. Consulte README.md.'
}
& $pythonProyecto -m streamlit run (Join-Path $PSScriptRoot 'streamlit_app.py')
