### Verify-BackupZipChecksum.ps1
## -------------------------------

<#
.TITLE
    Verify Backup ZIP Checksums

.SYNOPSIS
    Validates ZIP file integrity by comparing computed hashes to .sha256 files.

.DESCRIPTION
    Reads .sha256 files (lines like "<HEX> *filename.zip"), computes the ZIP's
    hash (default SHA256), and reports match/mismatch. Works for a single .sha256
    or all .sha256 files under a directory (with optional recursion).
    Exits 0 if all verifications pass; 1 otherwise.

.AUTHOR
    Created by: TheITManual

.VERSION
    1.4
#>

[CmdletBinding(DefaultParameterSetName='ByDir')]
param(
    # Verify a specific .sha256 file
    [Parameter(ParameterSetName='ByFile', Mandatory=$true, Position=0)]
    [ValidateScript({ Test-Path $_ })]
    [string]$Sha256Path,

    # Verify all .sha256 files in a directory (default)
    [Parameter(ParameterSetName='ByDir', Mandatory=$true, Position=0)]
    [ValidateScript({ Test-Path $_ })]
    [string]$Directory,

    # Hash algorithm (keep SHA256 to match your backup script)
    [ValidateSet('SHA256','SHA1','MD5','SHA384','SHA512')]
    [string]$Algorithm = 'SHA256',

    # Recurse when using -Directory
    [switch]$Recurse,

    # Quiet output (exit code only; useful in CI)
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Parse-Sha256Line {
    param([string]$Line)
    $t = $Line.Trim()
    if (-not $t -or $t.StartsWith('#') -or $t.StartsWith(';')) { return $null }

    # Match: <hex> [space(s)] [*]filename
    $m = [regex]::Match($t, '^([0-9A-Fa-f]+)[\s]+\*?(.*)$')
    if (-not $m.Success -or [string]::IsNullOrWhiteSpace($m.Groups[2].Value)) {
        throw "Unrecognized checksum line format: '$Line'"
    }

    [pscustomobject]@{
        ExpectedHash = $m.Groups[1].Value.ToUpperInvariant()
        FileName     = $m.Groups[2].Value.Trim()
        RawLine      = $Line
    }
}

function Verify-OneShaFile {
    param([string]$ShaPath, [string]$Algorithm)

    $shaDir = Split-Path -Parent $ShaPath
    $lines  = Get-Content -LiteralPath $ShaPath -ErrorAction Stop
    $items  = @()

    foreach ($line in $lines) {
        $entry = Parse-Sha256Line -Line $line
        if (-not $entry) { continue }

        $zipPath = (Join-Path $shaDir $entry.FileName)
        $exists  = Test-Path -LiteralPath $zipPath

        $actualHash = $null
        $match      = $false
        $status     = 'FAILED'
        $message    = $null

        try {
            if (-not $exists) { throw "ZIP not found: $zipPath" }
            $actualHash = (Get-FileHash -Algorithm $Algorithm -LiteralPath $zipPath).Hash.ToUpperInvariant()
            $match  = ($actualHash -eq $entry.ExpectedHash)
            $status = if ($match) { 'SUCCESS' } else { 'MISMATCH' }
            $message = if ($match) { 'Hash OK' } else { 'Hash mismatch' }
        } catch {
            $status  = 'ERROR'
            $message = $_.Exception.Message
        }

        $items += [pscustomobject]@{
            Sha256File   = $ShaPath
            ZipFile      = $zipPath
            Algorithm    = $Algorithm
            ExpectedHash = $entry.ExpectedHash
            ActualHash   = $actualHash
            Match        = $match
            Status       = $status
            Message      = $message
        }
    }

    return $items
}

# Gather .sha256 files to verify
$shaList = @()
switch ($PSCmdlet.ParameterSetName) {
    'ByFile' {
        if ([IO.Path]::GetExtension($Sha256Path) -notin @('.sha256','.txt')) {
            Write-Warning "File does not have .sha256 extension: $Sha256Path"
        }
        $shaList = @((Resolve-Path -LiteralPath $Sha256Path).Path)
    }
    'ByDir' {
        $opt = @{ Path = $Directory; Filter = '*.sha256'; File = $true }
        if ($Recurse) { $opt['Recurse'] = $true }
        $shaList = (Get-ChildItem @opt | Select-Object -ExpandProperty FullName)
        if (-not $shaList) {
            throw "No .sha256 files found in '$Directory'$(if ($Recurse) { ' (recursive)' })"
        }
    }
}

# Verify each .sha256
$allResults = @()
foreach ($sha in $shaList) {
    $allResults += Verify-OneShaFile -ShaPath $sha -Algorithm $Algorithm
}

# Output & exit code
$anyFail = $allResults | Where-Object { $_.Status -ne 'SUCCESS' }

if (-not $Quiet) {
    $allResults |
        Sort-Object Sha256File, ZipFile |
        Format-Table Sha256File, ZipFile, Algorithm, Status, Message -AutoSize

    if ($anyFail) {
        Write-Host "`nOne or more files FAILED verification." -ForegroundColor Red
    } else {
        Write-Host "`nAll files passed verification." -ForegroundColor Green
    }
}

exit ($(if ($anyFail) { 1 } else { 0 }))
