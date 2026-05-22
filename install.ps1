$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "PDF 工具箱 - 安装依赖 (PyMuPDF 开源版)"
Write-Host ""

function Find-Python {
    foreach ($name in @("python", "python3", "py")) {
        if (Get-Command $name -ErrorAction SilentlyContinue) {
            if ($name -eq "py") { return @("py", "-3") }
            return @($name)
        }
    }
    return $null
}

$py = Find-Python
if (-not $py) {
    Write-Host "ERROR: 未找到 Python"
    exit 1
}

& @py (Join-Path $PSScriptRoot "common\install_deps.py")
exit $LASTEXITCODE
