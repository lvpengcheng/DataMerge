param([switch]$InitializeDatabase, [string]$PrepareArchivesIn)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$pythonVersion = '3.11.17'
$dotnetVersion = '9.0.21'
$runtimeRoot = Join-Path $PSScriptRoot 'runtime'
$pythonHome = Join-Path $runtimeRoot "python-$pythonVersion"
$dotnetHome = Join-Path $runtimeRoot "dotnet-$dotnetVersion"
$venvHome = Join-Path $PSScriptRoot 'venv-iis'
$pythonArchive = @{
    Name = 'cpython-3.11.17+20261013-x86_64-pc-windows-msvc-install_only_stripped.tar.gz'
    Url = 'https://github.com/astral-sh/python-build-standalone/releases/download/20261013/cpython-3.11.17%2B20261013-x86_64-pc-windows-msvc-install_only_stripped.tar.gz'
    Hash = '54a38c5e7cc1a5fd2b0182b495f29606dc1455d18ebfd3b5bf0a3c09253b5031'
    Algorithm = 'SHA256'
}
$dotnetArchive = @{
    Name = "dotnet-runtime-$dotnetVersion-win-x64.zip"
    Url = "https://builds.dotnet.microsoft.com/dotnet/Runtime/$dotnetVersion/dotnet-runtime-$dotnetVersion-win-x64.zip"
    Hash = 'b345456a7f595204c50080450a17ab92e9840093c3a0bc9869c5673ef8c2a4421717be3f58ea112fd44fff0d09eb0d5b44aadbe86e4ce46452340eb163b87374'
    Algorithm = 'SHA512'
}

function Invoke-Checked($Executable, [string[]]$Arguments) {
    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Executable exited with code $LASTEXITCODE."
    }
}

function Assert-RuntimeArchive($Archive, $Spec) {
    $hasher = [Security.Cryptography.HashAlgorithm]::Create($Spec.Algorithm)
    $stream = [IO.File]::OpenRead($archive)
    try {
        $actualHash = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '')
    } finally {
        $stream.Dispose()
        $hasher.Dispose()
    }
    if ($actualHash -ne $Spec.Hash) {
        throw "Checksum mismatch: $Archive. Replace it with the archive from $($Spec.Url)."
    }
}

function Get-RuntimeArchive($Spec, $Directory) {
    $archive = Join-Path $Directory $Spec.Name
    if (Test-Path -LiteralPath $archive -PathType Leaf) {
        Assert-RuntimeArchive $archive $Spec
        return $archive
    }
    $partial = "$archive.part-$([Guid]::NewGuid().ToString('N'))"
    Write-Host "Downloading $($Spec.Name)..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        try {
            Invoke-WebRequest -UseBasicParsing -Uri $Spec.Url -OutFile $partial -TimeoutSec 180
        } catch {
            throw "Cannot download $($Spec.Name). Download $($Spec.Url) and put it in $Directory. $($_.Exception.Message)"
        }
        Assert-RuntimeArchive $partial $Spec
        Move-Item -LiteralPath $partial -Destination $archive
    } finally {
        if (Test-Path -LiteralPath $partial -PathType Leaf) {
            Remove-Item -LiteralPath $partial -Force
        }
    }
    return $archive
}

if ($PrepareArchivesIn) {
    New-Item -ItemType Directory -Force -Path $PrepareArchivesIn | Out-Null
    Get-RuntimeArchive $pythonArchive $PrepareArchivesIn | Out-Null
    Get-RuntimeArchive $dotnetArchive $PrepareArchivesIn | Out-Null
    Write-Host 'Both Windows runtime archives verified.'
    return
}

function Test-Python($Executable, [switch]$VirtualEnvironment) {
    if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {
        throw "Python executable missing: $Executable."
    }
    $check = "import ctypes, platform, ssl, struct; assert platform.python_version() == '$pythonVersion' and struct.calcsize('P') == 8, (platform.python_version(), platform.machine())"
    if ($VirtualEnvironment) {
        $check += '; import pathlib, sys; assert pathlib.Path(sys.base_prefix).resolve() == pathlib.Path(sys.argv[1]).resolve(), sys.base_prefix'
    }
    Invoke-Checked $Executable @('-c', $check, $pythonHome)
}

$installLock = $null
Push-Location -LiteralPath $PSScriptRoot
try {
    New-Item -ItemType Directory -Force -Path $runtimeRoot | Out-Null
    $installLock = [IO.File]::Open((Join-Path $runtimeRoot '.install.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    $stage = Join-Path $runtimeRoot ('.install-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        if (-not (Test-Path -LiteralPath $pythonHome)) {
            Write-Host "Installing Python $pythonVersion..."
            $archive = Get-RuntimeArchive $pythonArchive $PSScriptRoot
            Invoke-Checked 'tar.exe' @('-xzf', $archive, '-C', $stage)
            Test-Python (Join-Path $stage 'python\python.exe')
            Move-Item -LiteralPath (Join-Path $stage 'python') -Destination $pythonHome
        }
        $basePython = Join-Path $pythonHome 'python.exe'
        Test-Python $basePython

        if (-not (Test-Path -LiteralPath $dotnetHome)) {
            Write-Host "Installing .NET $dotnetVersion..."
            $archive = Get-RuntimeArchive $dotnetArchive $PSScriptRoot
            $dotnetStage = Join-Path $stage 'dotnet'
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [IO.Compression.ZipFile]::ExtractToDirectory($archive, $dotnetStage)
            if (-not (Test-Path -LiteralPath (Join-Path $dotnetStage "shared\Microsoft.NETCore.App\$dotnetVersion\coreclr.dll"))) {
                throw ".NET $dotnetVersion was not found in the runtime archive."
            }
            Move-Item -LiteralPath $dotnetStage -Destination $dotnetHome
        }
    } finally {
        $resolvedStage = (Resolve-Path -LiteralPath $stage).Path
        if ((Split-Path -Parent $resolvedStage) -eq (Resolve-Path -LiteralPath $runtimeRoot).Path -and
            (Split-Path -Leaf $resolvedStage) -match '^\.install-[0-9a-f]{32}$') {
            Remove-Item -LiteralPath $resolvedStage -Recurse -Force
        }
    }

    $python = Join-Path $venvHome 'Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        Invoke-Checked $basePython @('-m', 'venv', '--without-pip', $venvHome)
    }
    Test-Python $python -VirtualEnvironment
    $pipStatus = Invoke-Checked $python @('-c', "from importlib.util import find_spec; print('present' if find_spec('pip') else 'missing')")
    if ($pipStatus -eq 'missing') {
        Write-Host 'Initializing pip in the IIS virtual environment...'
        Invoke-Checked $python @('-X', 'utf8', '-m', 'ensurepip', '--upgrade', '--default-pip')
    }
    Write-Host 'Installing Python dependencies...'
    Invoke-Checked $python @('-X', 'utf8', '-m', 'pip', 'install', '--progress-bar', 'off', '-r', 'requirements.txt', '-c', 'calculation-constraints.txt')
    Invoke-Checked $python @('-X', 'utf8', '-m', 'pip', 'check')

    $env:DOTNET_ROOT = $dotnetHome
    $env:DOTNET_ROOT_X64 = $dotnetHome
    $env:PATH = "$dotnetHome;$env:PATH"
    $runtimeCheck = "import aspose_init; assert aspose_init.init_aspose(); from System import Environment; actual = Environment.Version.ToString(); assert actual == '$dotnetVersion', actual; print('.NET ' + actual)"
    Invoke-Checked $python @('-c', $runtimeCheck)
    Invoke-Checked $python @('--version')

    if (-not (Test-Path -LiteralPath '.env')) {
        Copy-Item -LiteralPath '.env.example' -Destination '.env'
    }
    if ($InitializeDatabase) {
        Invoke-Checked $python @('-m', 'backend.database.init_db')
    }
    Write-Host "IIS environment ready. Configure .env before starting the site."
} finally {
    if ($installLock) { $installLock.Dispose() }
    Pop-Location
}
