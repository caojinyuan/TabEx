param(
    [ValidateSet('Ensure', 'Install', 'Run', 'Build')]
    [string]$Action = 'Ensure'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = '3.13.13'
$packageHash = '2BEACFF5A52B33E0919EA6C06CCCDDDD9E245EF55F69E9ACBDB6FE2CA8B7831B49D4515438DD2B3DAB8D9ED772F10DEA839BEFDC6D0DBF809101D3EB6A111D29'
$runtimeRoot = Join-Path $root '.venv'
$installation = Join-Path $runtimeRoot "python-$version"
$python = Join-Path $installation 'python.exe'
$lock = $null
$stage = $null

try {
    if (-not [Environment]::Is64BitOperatingSystem) {
        throw 'The project Python requires 64-bit Windows.'
    }
    foreach ($name in @('PYTHONHOME', 'PYTHONPATH', 'PYTHONSTARTUP', 'PYTHONUSERBASE', 'PYTHONSAFEPATH',
                        'VIRTUAL_ENV', 'PIP_TARGET', 'PIP_PREFIX', 'PIP_USER', 'PIP_REQUIRE_VIRTUALENV')) {
        Remove-Item "Env:$name" -ErrorAction SilentlyContinue
    }
    $env:PYTHONNOUSERSITE = '1'
    $env:PYTHONUTF8 = '1'
    $env:PIP_CONFIG_FILE = 'NUL'
    $legacyBackup = $null
    if (Test-Path (Join-Path $runtimeRoot 'pyvenv.cfg')) {
        $legacyBackup = Join-Path $root ('.venv-legacy-' + [Guid]::NewGuid().ToString('N'))
        Move-Item -Path $runtimeRoot -Destination $legacyBackup
        Write-Host "Previous virtual environment preserved at $legacyBackup"
    }
    New-Item -ItemType Directory -Force -Path $runtimeRoot | Out-Null
    if ($legacyBackup -and (Test-Path (Join-Path $legacyBackup 'cache'))) {
        Copy-Item -Path (Join-Path $legacyBackup 'cache') -Destination $runtimeRoot -Recurse
    }
    if ($Action -ne 'Ensure' -or -not (Test-Path $python)) {
        $lock = [IO.File]::Open((Join-Path $runtimeRoot 'setup.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    }

    if (-not (Test-Path $python)) {
        if (Test-Path $installation) {
            throw "Incomplete runtime at $installation. Rename that directory and retry."
        }
        $cache = Join-Path $runtimeRoot 'cache'
        New-Item -ItemType Directory -Force -Path $cache | Out-Null
        $package = Join-Path $cache "python-$version.zip"
        if (-not (Test-Path $package)) {
            $url = "https://api.nuget.org/v3-flatcontainer/python/$version/python.$version.nupkg"
            Write-Host "Downloading official Python $version x64 ..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $download = "$package.download"
            try {
                Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $download
                if ((Get-FileHash -Algorithm SHA512 -Path $download).Hash -ne $packageHash) {
                    throw 'Python package SHA-512 verification failed.'
                }
                Move-Item -Path $download -Destination $package -Force
            } finally {
                Remove-Item $download -Force -ErrorAction SilentlyContinue
            }
        }
        if ((Get-FileHash -Algorithm SHA512 -Path $package).Hash -ne $packageHash) {
            throw "Cached Python package is invalid. Remove $package and retry."
        }
        $stage = Join-Path $runtimeRoot ('.setup-' + [Guid]::NewGuid().ToString('N'))
        Expand-Archive -Path $package -DestinationPath $stage
        if (-not (Test-Path (Join-Path $stage 'tools\python.exe'))) {
            throw 'Python package does not contain the expected runtime.'
        }
        Move-Item -Path (Join-Path $stage 'tools') -Destination $installation
        Remove-Item -Path $stage -Recurse -Force
        $stage = $null
    }

    & $python -I -c "import struct, sys; from pathlib import Path; assert sys.version_info[:3] == (3, 13, 13), sys.version; assert struct.calcsize('P') == 8; assert Path(sys.base_prefix).resolve() == Path(sys.argv[1]).resolve(); assert Path(sys.prefix).resolve() == Path(sys.argv[1]).resolve(); import pip; print('Project Python:', sys.executable); print('Version:', sys.version.split()[0])" $installation
    if ($LASTEXITCODE -ne 0) {
        throw 'Project Python validation failed. No system Python fallback is allowed.'
    }

    Push-Location $root
    try {
        switch ($Action) {
            'Install' {
                & $python -X utf8 -m pip --isolated install --no-user -r (Join-Path $root 'requirements.txt')
                if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
            }
            'Run' {
                & $python -X utf8 -c 'import PyQt5, comtypes, win32api, requests, send2trash'
                if ($LASTEXITCODE -ne 0) { throw 'Run 0_install_requirements.bat before starting TabEx.' }
                Start-Process -FilePath (Join-Path $installation 'pythonw.exe') -WorkingDirectory $root -ArgumentList ('-OO "' + (Join-Path $root 'TabEx.py') + '"')
            }
            'Build' {
                Write-Host 'Close all TabEx windows before building. Existing EXE is kept until validation passes.'
                & $python -X utf8 (Join-Path $root 'tools\build_release.py')
                if ($LASTEXITCODE -ne 0) { throw 'Build or runtime validation failed. The previous EXE was not replaced.' }
            }
        }
    } finally {
        Pop-Location
    }
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)"
    exit 1
} finally {
    if ($lock) { $lock.Dispose() }
    if ($stage -and (Test-Path $stage)) { Remove-Item $stage -Recurse -Force }
}