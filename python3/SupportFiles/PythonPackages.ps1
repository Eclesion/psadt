$pythonExe = $ExecutionContext.InvokeCommand.ExpandString($Global:Pkg.PythonExePath)

if (-not (Test-Path -LiteralPath $pythonExe -PathType Leaf)) {
    throw "Python executable not found: $pythonExe"
}

Write-ADTLogEntry -Message "Upgrading pip..." -Severity 1
$result = Start-ADTProcess -FilePath $pythonExe -ArgumentList '-m pip install --upgrade pip' -NoWait:$false -PassThru
if ($result.ExitCode -ne 0) {
    throw "pip upgrade failed with exit code [$($result.ExitCode)]."
}

if ($Global:Pkg.ContainsKey('PipPackages') -and $Global:Pkg.PipPackages -and $Global:Pkg.PipPackages.Count -gt 0) {
    $pkgList = $Global:Pkg.PipPackages -join ' '
    $args = "-m pip install $pkgList"

    if ($Global:Pkg.ContainsKey('PipExtraArgs') -and $Global:Pkg.PipExtraArgs) {
        $args = "$args $($Global:Pkg.PipExtraArgs)"
    }

    Write-ADTLogEntry -Message "Installing Python packages: $pkgList" -Severity 1
    $result = Start-ADTProcess -FilePath $pythonExe -ArgumentList $args -NoWait:$false -PassThru
    if ($result.ExitCode -ne 0) {
        throw "pip install failed with exit code [$($result.ExitCode)]."
    }
}

if (
    $Global:Pkg.ContainsKey('ValidatePythonImports') -and $Global:Pkg.ValidatePythonImports -and
    $Global:Pkg.ContainsKey('PythonImportModules') -and $Global:Pkg.PythonImportModules -and
    $Global:Pkg.PythonImportModules.Count -gt 0
) {
    $tempPy = Join-Path $env:TEMP ("test_python_imports_{0}.py" -f ([guid]::NewGuid().ToString('N')))
    $moduleList = ($Global:Pkg.PythonImportModules | ForEach-Object { "'$_'" }) -join ', '

    $code = @"
import sys
import importlib.util

modules = [$moduleList]
missing = [m for m in modules if importlib.util.find_spec(m) is None]
sys.exit(1 if missing else 0)
"@

    Set-Content -LiteralPath $tempPy -Value $code -Encoding UTF8

    try {
        $result = Start-ADTProcess -FilePath $pythonExe -ArgumentList "`"$tempPy`"" -NoWait:$false -PassThru
        if ($result.ExitCode -ne 0) {
            throw "Python import validation failed."
        }
        Write-ADTLogEntry -Message "Python import validation succeeded." -Severity 1
    }
    finally {
        Remove-Item -LiteralPath $tempPy -Force -ErrorAction SilentlyContinue
    }
}