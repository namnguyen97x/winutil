[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$IsoSource,

    [string]$OutputPath = (Join-Path (Get-Location) "microwin.iso"),

    [string]$WorkingDirectory = (Join-Path (Get-Location) "microwin-build"),

    [int]$ImageIndex,

    [string]$PreferredEdition = "Professional",

    [bool]$InjectDrivers = $false,

    [string]$DriverPath,

    [bool]$ImportDrivers = $false,

    [bool]$DisableWPBT = $false,

    [bool]$AllowUnsupportedHardware = $false,

    [bool]$SkipFirstLogonAnimation = $true,

    [bool]$IncludeVirtIO = $false,

    [string]$AutoConfigPath,

    [string]$UserName = "User",

    [string]$UserPassword,

    [bool]$UseEsd = $false,

    [bool]$CopyToUSB = $false,

    [bool]$KeepWorkingDirectory = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    throw "Invoke-MicrowinHeadless.ps1 must be run from an elevated PowerShell session."
}

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
Write-Host "Repository root detected at $repoRoot"

# Load all functions so Invoke-Microwin and dependencies are available.
Get-ChildItem -Path (Join-Path $repoRoot "functions") -Recurse -Filter "*.ps1" | Sort-Object FullName | ForEach-Object {
    . $_.FullName
}

# Ensure the global $sync variable exists for functions that expect it
if (-not (Get-Variable -Name sync -Scope Global -ErrorAction SilentlyContinue)) {
    $global:sync = $null
}

function Ensure-Oscdimg {
    $oscdimgPath = Join-Path $env:TEMP "oscdimg.exe"
    $hasOscdimg = [bool](Get-Command oscdimg.exe -ErrorAction SilentlyContinue) -or (Test-Path -Path $oscdimgPath -PathType Leaf)
    if (-not $hasOscdimg) {
        Write-Host "oscdimg.exe not found. Downloading from repository..."
        Microwin-GetOscdimg -oscdimgPath $oscdimgPath
    }
}

Ensure-Oscdimg

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Path $outputFullPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$workRoot = [System.IO.Path]::GetFullPath($WorkingDirectory)
Write-Host "Using working directory: $workRoot"
New-Item -ItemType Directory -Path $workRoot -Force | Out-Null

$mountDir = Join-Path $workRoot "mount"
$scratchDir = Join-Path $workRoot "scratch"
$downloadsDir = Join-Path $workRoot "downloads"

foreach ($path in @($mountDir, $scratchDir, $downloadsDir)) {
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Recurse -Force
    }
    New-Item -ItemType Directory -Path $path -Force | Out-Null
}

if ($InjectDrivers -and [string]::IsNullOrWhiteSpace($DriverPath)) {
    throw "DriverPath must be provided when InjectDrivers is enabled."
}

if ($AutoConfigPath) {
    $AutoConfigPath = [System.IO.Path]::GetFullPath($AutoConfigPath)
    if (-not (Test-Path -Path $AutoConfigPath -PathType Leaf)) {
        throw "AutoConfigPath '$AutoConfigPath' does not exist."
    }
}

$resolvedDriverPath = $null
if ($DriverPath) {
    $resolvedDriverPath = [System.IO.Path]::GetFullPath($DriverPath)
    if (-not (Test-Path -Path $resolvedDriverPath)) {
        throw "DriverPath '$resolvedDriverPath' does not exist."
    }
}

function Get-LocalIsoPath {
    param(
        [string]$Source,
        [string]$DestinationFolder
    )

    if ($Source -match '^(http|https)://') {
        $targetIso = Join-Path $DestinationFolder "source.iso"
        Write-Host "Downloading ISO from $Source ..."

        # If aria2c is available, prefer it for high-speed, multi-connection
        # downloads. Use a large split count and enable continue so partial
        # downloads resume on retries. If aria2c is not found, fall back to
        # Invoke-WebRequest.
        $aria2 = Get-Command aria2c -ErrorAction SilentlyContinue
        if ($aria2) {
            Write-Host "aria2c found; downloading with parallel connections..."
            $targetDir = Split-Path -Path $targetIso -Parent
            $targetName = Split-Path -Path $targetIso -Leaf

            $ariaArgs = @(
                '--max-connection-per-server=16',
                '--split=16',
                '--min-split-size=1M',
                '--max-tries=5',
                '--retry-wait=5',
                '--continue=true',
                '--file-allocation=none',
                "--dir=$targetDir",
                "--out=$targetName",
                $Source
            )

            # Run aria2c and check result
            & aria2c @ariaArgs
            if (-not (Test-Path -Path $targetIso -PathType Leaf)) {
                throw "aria2c failed to download the ISO to '$targetIso'."
            }
        }
        else {
            Write-Host "aria2c not found; falling back to Invoke-WebRequest..."
            Invoke-WebRequest -Uri $Source -OutFile $targetIso
        }

        return $targetIso
    }

    $resolved = [System.IO.Path]::GetFullPath($Source)
    if (-not (Test-Path -Path $resolved -PathType Leaf)) {
        throw "ISO source '$resolved' does not exist."
    }
    return $resolved
}

$isoPath = Get-LocalIsoPath -Source $IsoSource -DestinationFolder $downloadsDir

$mountedImage = $null
try {
    Write-Host "Mounting ISO image..."
    $mountedImage = Mount-DiskImage -ImagePath $isoPath -PassThru
    $volume = $mountedImage | Get-Volume
    $driveLetter = $volume.DriveLetter
    if (-not $driveLetter) {
        throw "Unable to determine drive letter for mounted ISO image."
    }

    Write-Host "Copying files from $driveLetter`: to $mountDir ..."
    Copy-Files "$driveLetter`:" "$mountDir" -Recurse -Force
}
finally {
    if ($mountedImage) {
        Write-Host "Dismounting ISO image..."
        Dismount-DiskImage -ImagePath $mountedImage.ImagePath | Out-Null
    }
}

$installWim = Join-Path $mountDir "sources\install.wim"
$installEsd = Join-Path $mountDir "sources\install.esd"

$wimSource = if (Test-Path -Path $installWim -PathType Leaf) {
    $installWim
} elseif (Test-Path -Path $installEsd -PathType Leaf) {
    $installEsd
} else {
    throw "Neither install.wim nor install.esd was found in the ISO."
}

Write-Host "Inspecting Windows images inside $(Split-Path -Leaf $wimSource)..."
$images = Get-WindowsImage -ImagePath $wimSource

if (-not $images) {
    throw "No images were found inside $wimSource."
}

if ($PSBoundParameters.ContainsKey('ImageIndex')) {
    $selectedImage = $images | Where-Object { $_.ImageIndex -eq $ImageIndex } | Select-Object -First 1
    if (-not $selectedImage) {
        throw "Image index $ImageIndex was not found. Available indexes: $($images.ImageIndex -join ', ')."
    }
} else {
    $selectedImage = $images | Where-Object { $_.EditionId -eq $PreferredEdition } | Select-Object -First 1
    if (-not $selectedImage) {
        Write-Warning "Edition '$PreferredEdition' was not found. Falling back to the first image index."
        $selectedImage = $images | Select-Object -First 1
    }
}

$imageIndexToUse = [int]$selectedImage.ImageIndex
$imageName = $selectedImage.ImageName
Write-Host "Using image index $imageIndexToUse ($imageName)"

$microwinOptions = [pscustomobject]@{
    ImageIndex               = $imageIndexToUse
    ImageName                = $imageName
    CopyToUSB                = $CopyToUSB
    InjectDrivers            = $InjectDrivers
    DriverPath               = $resolvedDriverPath
    ImportDrivers            = $ImportDrivers
    DisableWPBT              = $DisableWPBT
    AllowUnsupportedHardware = $AllowUnsupportedHardware
    SkipFirstLogonAnimation  = $SkipFirstLogonAnimation
    CopyVirtIO               = $IncludeVirtIO
    MountDir                 = $mountDir
    ScratchDir               = $scratchDir
    AutoConfigPath           = $AutoConfigPath
    UserName                 = $UserName
    UserPassword             = $UserPassword
    UseEsd                   = $UseEsd
    OutputIsoPath            = $outputFullPath
}

try {
    Invoke-Microwin -MicrowinOptions $microwinOptions
    Write-Host "`nMicroWin ISO created at $outputFullPath"
}
finally {
    if (-not $KeepWorkingDirectory -and (Test-Path -Path $workRoot)) {
        Write-Host "Cleaning up working directory..."
        Remove-Item -Path $workRoot -Recurse -Force
    }
}
