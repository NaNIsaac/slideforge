<#
.SYNOPSIS
Install, remove or list the math fonts of fonts.json for the current user, without admin rights.

.DESCRIPTION
install downloads each font from its pinned URL, checks the SHA-256 of the download (and of the
font file inside a zip), copies it to the per-user font folder and registers it under HKCU, as
Windows does when a user without admin rights installs a font. A record of what was written goes
to the state file, and remove undoes exactly that: fonts this script did not install are never
touched. PowerPoint and other Office programs see a new font after they are restarted.

-FontDir, -RegistryKey and -State point everything at other places, which the tests use: on Windows
they run against a temporary key under HKCU that already holds values. -RegistryFile sends the
registry values to a JSON file instead, for the tests on other systems.

.EXAMPLE
./fonts.ps1 list
./fonts.ps1 install -Font 'Fira Math' -WhatIf
./fonts.ps1 install -Font 'Fira Math', 'Noto Sans Math'
./fonts.ps1 remove -Font 'Fira Math'
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateSet('install', 'remove', 'list')]
    [string]$Command,
    [string[]]$Font,
    [string]$Manifest = (Join-Path $PSScriptRoot 'fonts.json'),
    [string]$FontDir = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'),
    [string]$RegistryKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts',
    [string]$RegistryFile,
    [string]$State = (Join-Path $env:LOCALAPPDATA 'math-fonts\installed.json'),
    # A folder holding the downloads already (offline use and tests); missing files are downloaded.
    [string]$Source
)

$ErrorActionPreference = 'Stop'

function Read-Json([string]$Path, $Default) {
    if (Test-Path $Path) { Get-Content $Path -Raw | ConvertFrom-Json -AsHashtable } else { $Default }
}

function Write-Json([string]$Path, $Value) {
    New-Item -ItemType Directory -Force (Split-Path $Path) | Out-Null
    $Value | ConvertTo-Json -Depth 5 | Set-Content $Path -Encoding utf8
}

function Get-Registered([string]$Name) {
    if ($RegistryFile) { return (Read-Json $RegistryFile @{})[$Name] }
    (Get-ItemProperty $RegistryKey -Name $Name -ErrorAction SilentlyContinue).$Name
}

function Set-Registered([string]$Name, $Value) {
    if ($RegistryFile) {
        $values = Read-Json $RegistryFile @{}
        if ($null -eq $Value) { $values.Remove($Name) } else { $values[$Name] = $Value }
        Write-Json $RegistryFile $values
    } elseif ($null -eq $Value) {
        Remove-ItemProperty $RegistryKey -Name $Name
    } else {
        # New-Item -Force on an existing key replaces it and drops every value in it (#5).
        if (-not (Test-Path $RegistryKey)) { New-Item $RegistryKey | Out-Null }
        New-ItemProperty $RegistryKey -Name $Name -Value $Value -PropertyType String | Out-Null
    }
}

function Assert-Hash([string]$Path, [string]$Expected, [string]$What) {
    $actual = (Get-FileHash $Path -Algorithm SHA256).Hash
    if ($actual -ne $Expected.ToUpperInvariant()) {
        throw "$What`: SHA-256 $actual, expected $Expected; not installed. The download changed: check the new SHA-256 at the source before updating fonts.json"
    }
}

function Get-FontFile($Entry, [string]$Work) {
    $name = Split-Path $Entry.url -Leaf
    $download = Join-Path $Work $name
    $cached = if ($Source) { Join-Path $Source $name }
    if ($cached -and (Test-Path $cached)) {
        Copy-Item $cached $download
    } else {
        Invoke-WebRequest $Entry.url -OutFile $download
    }
    Assert-Hash $download $Entry.sha256 $name
    if (-not $Entry.member) { return $download }
    $unpacked = Join-Path $Work 'unpacked'
    Expand-Archive $download $unpacked -Force
    $file = Join-Path $unpacked $Entry.member
    Assert-Hash $file $Entry.member_sha256 $Entry.member
    $file
}

$fonts = (Read-Json $Manifest $null).fonts
if ($Font) {
    $unknown = $Font | Where-Object { $_ -notin $fonts.name }
    if ($unknown) { throw "not in $(Split-Path $Manifest -Leaf): $($unknown -join ', ')" }
    $fonts = $fonts | Where-Object { $_.name -in $Font }
}
$installed = Read-Json $State @{}

switch ($Command) {
    'list' {
        foreach ($entry in $fonts) {
            $value = "$($entry.face) (OpenType)"
            [pscustomobject]@{
                Font       = $entry.name
                Version    = $entry.version
                License    = $entry.license
                Installed  = [bool](Get-Registered $value) -or (Test-Path (Join-Path $FontDir $entry.file))
                ByThisTool = $installed.ContainsKey($entry.name)
            }
        }
    }
    'install' {
        foreach ($entry in $fonts) {
            $target = Join-Path $FontDir $entry.file
            $value = "$($entry.face) (OpenType)"
            if ((Test-Path $target) -or (Get-Registered $value)) {
                Write-Host "$($entry.name): already installed, left as it is"
                continue
            }
            $where = if ($RegistryFile) { $RegistryFile } else { $RegistryKey }
            $what = "$($entry.name) $($entry.version) -> $target; registry value '$value' in $where"
            if (-not $RegistryFile -and -not (Test-Path $RegistryKey)) { $what += ' (key created)' }
            if (-not $PSCmdlet.ShouldProcess($what, 'install font')) { continue }
            $work = New-Item -ItemType Directory (Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetRandomFileName()))
            try {
                $file = Get-FontFile $entry $work.FullName
                New-Item -ItemType Directory -Force $FontDir | Out-Null
                Copy-Item $file $target
                Set-Registered $value $target
                $installed[$entry.name] = @{ file = $target; value = $value; version = $entry.version }
                Write-Json $State $installed
                Write-Host "$($entry.name) $($entry.version): installed"
            } finally {
                Remove-Item $work -Recurse -Force
            }
        }
    }
    'remove' {
        foreach ($entry in $fonts) {
            $record = $installed[$entry.name]
            if (-not $record) {
                Write-Host "$($entry.name): not installed by this script, left as it is"
                continue
            }
            if (-not $PSCmdlet.ShouldProcess("$($entry.name) $($record.file)", 'remove font')) { continue }
            if ((Get-Registered $record.value) -eq $record.file) { Set-Registered $record.value $null }
            if (Test-Path $record.file) { Remove-Item $record.file }
            $installed.Remove($entry.name)
            Write-Json $State $installed
            Write-Host "$($entry.name): removed"
        }
    }
}
