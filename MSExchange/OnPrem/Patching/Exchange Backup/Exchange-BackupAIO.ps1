### Consolidated Exchange Backup Script
## -------------------------------------

<#
.TITLE
    Exchange Server Backup Script (Full Logging + IIS Backup + ZIP Checksum)

.SYNOPSIS
    Back up Exchange transport rules, IIS config (XML + full IIS backup), Windows services,
    and a key Web.config, with per-run folder, transcript logging, and ZIP checksum.

.DESCRIPTION
    Creates a per-run backup folder and captures:
      • Exchange Transport Rules (Export-TransportRuleCollection)
      • IIS App Pools & Sites (XML via appcmd)
      • IIS Full Backup (appcmd add backup "<timestamped_name>")
      • Windows Services (rich details via CIM to CSV)
      • A specified Web.config file
    Generates a ZIP of the run folder, computes a SHA-256 checksum file (.sha256),
    and copies both to a network destination. Works in Windows PowerShell 5.1 and PowerShell 7+.

.AUTHOR
    Created by: TheITManual

.VERSION
    2.5
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$RootBackupDir         = "<RootBackupDir>",         # e.g. D:\Ops\Exchange\Backups\BackedUpFiles (Local backup drive & folder)

    [Parameter(Mandatory=$true)]
    [string]$ExchangeModulePath    = "<ExchangeModulePath>",    # e.g. D:\Exchange\Bin\RemoteExchange.ps1 (Path to Exchange RemoteExchange.ps1)

    [Parameter(Mandatory=$true)]
    [string]$WebConfigSourcePath   = "<WebConfigSourcePath>",   # e.g. D:\Exchange\ClientAccess\ecp\Web.config	(Path to Web.config (ClientAccess\ECP))

    [Parameter(Mandatory=$true)]
    [string]$NetworkDestinationDir = "<NetworkDestinationDir>"  # e.g. \\fileserver\share\Backups\ExchangeOnPrem\ (Network destination folder)
)

# ---------- Safety & context ----------
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$HostName  = $env:COMPUTERNAME
$RunId     = Get-Date -Format 'yyyyMMdd_HHmmss'
$TimeStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$RunFolder = Join-Path $RootBackupDir $RunId

# *** Seed a temporary log path so early logging works ***
$script:LogFilePath = Join-Path $env:TEMP "BackupLog_${HostName}_$TimeStamp.txt"

# ---------- Helpers ----------
function Log-Message {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('SUCCESS','FAILED','INFO')][string]$Status = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts - [$HostName] - [$Status] - $Message" | Out-File -FilePath $script:LogFilePath -Append
}

function Ensure-DirectoryExists {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Log-Message "Created directory: $Path" -Status SUCCESS
    }
}

# ---------- Pre-flight ----------
# Admin check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Please run this script as Administrator."
}

Ensure-DirectoryExists -Path $RootBackupDir
Ensure-DirectoryExists -Path $RunFolder

# Disk space (>= 2 GB free on run drive)
$drive  = (Get-Item $RunFolder).PSDrive
$freeGB = [math]::Round(($drive.Free/1GB),2)
if ($freeGB -lt 2) { throw "Not enough free space on $($drive.Root) ($freeGB GB available)." }

# Network destination reachable (if UNC)
if ($NetworkDestinationDir -like '\\*' -and -not (Test-Path -LiteralPath $NetworkDestinationDir)) {
    throw "Network destination not reachable: $NetworkDestinationDir"
}

# ---------- Logging & transcript ----------
# Move temp log into the run folder and switch to final path
$finalLogPath = Join-Path $RunFolder "BackupLog_${HostName}_$TimeStamp.txt"
if (Test-Path -LiteralPath $script:LogFilePath) {
    Move-Item -LiteralPath $script:LogFilePath -Destination $finalLogPath -Force
}
$script:LogFilePath = $finalLogPath

$TranscriptPath   = Join-Path $RunFolder "Transcript_${HostName}_$TimeStamp.txt"
$TranscriptStarted = $false
Start-Transcript -Path $TranscriptPath -ErrorAction SilentlyContinue | Out-Null
$TranscriptStarted = $true
Log-Message "Backup run started. RunId=$RunId" -Status INFO

try {
    # ---------- Load EMS (Exchange Management Shell) ----------
    if (-not (Get-Command Export-TransportRuleCollection -ErrorAction SilentlyContinue)) {
        if (-not (Test-Path -LiteralPath $ExchangeModulePath)) {
            throw "Exchange module not found at: $ExchangeModulePath"
        }
        . $ExchangeModulePath
        if (Get-Command Connect-ExchangeServer -ErrorAction SilentlyContinue) {
            Connect-ExchangeServer -Auto -ErrorAction Stop | Out-Null
        }
        Log-Message "Exchange Management Shell loaded & connected." -Status SUCCESS
    } else {
        Log-Message "Exchange Management Shell already available." -Status INFO
    }

    # ---------- Resolve IIS appcmd.exe (64-bit safe) ----------
    $appcmdCandidates = @(
        (Join-Path $env:windir 'System32\inetsrv\appcmd.exe'),
        (Join-Path $env:windir 'SysNative\inetsrv\appcmd.exe')
    )
    $AppCmd = $appcmdCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $AppCmd) { throw "IIS appcmd.exe not found in expected locations." }
    Log-Message "Resolved appcmd.exe at: $AppCmd" -Status INFO

    # ---------- Artifact directories ----------
    $TransportRulesPath = Join-Path $RunFolder 'ExchangeTransportRules'
    $IISAppPoolsPath    = Join-Path $RunFolder 'IIS\AppPools'
    $IISSitesPath       = Join-Path $RunFolder 'IIS\Sites'
    $IISFullBackupPath  = Join-Path $RunFolder 'IIS\FullBackup'
    $ServicesPath       = Join-Path $RunFolder 'Services'
    $WebConfigPath      = Join-Path $RunFolder 'Webconfig'
    foreach ($p in @($TransportRulesPath,$IISAppPoolsPath,$IISSitesPath,$IISFullBackupPath,$ServicesPath,$WebConfigPath)) {
        Ensure-DirectoryExists -Path $p
    }

    # ===== Exchange Transport Rules =====
    try {
        if (Get-Command Export-TransportRuleCollection -ErrorAction SilentlyContinue) {
            $file = Export-TransportRuleCollection
            $TransportRulesFile = Join-Path $TransportRulesPath "TransportRules_$TimeStamp.xml"
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                Set-Content -Path $TransportRulesFile -Value $file.FileData -AsByteStream
            } else {
                Set-Content -Path $TransportRulesFile -Value $file.FileData -Encoding Byte
            }
            Log-Message "Transport Rules saved: $TransportRulesFile" -Status SUCCESS
        } else {
            throw "Export-TransportRuleCollection not available."
        }
    } catch {
        Log-Message "Error backing up Transport Rules: $_" -Status FAILED
    }

    # ===== IIS App Pools (XML) =====
    try {
        $IISAppPoolsFile = Join-Path $IISAppPoolsPath "AppPools_$TimeStamp.xml"
        & $AppCmd list apppool /config /xml > $IISAppPoolsFile
        Log-Message "IIS App Pools saved: $IISAppPoolsFile" -Status SUCCESS
    } catch {
        Log-Message "Error backing up IIS App Pools: $_" -Status FAILED
    }

    # ===== IIS Sites (XML) =====
    try {
        $IISSitesFile = Join-Path $IISSitesPath "Sites_$TimeStamp.xml"
        & $AppCmd list site /config /xml > $IISSitesFile
        Log-Message "IIS Sites saved: $IISSitesFile" -Status SUCCESS
    } catch {
        Log-Message "Error backing up IIS Sites: $_" -Status FAILED
    }

    # ===== IIS Full Backup (appcmd add backup "<name>") =====
    try {
        $IISBackupName = "IISBackup_$TimeStamp"
        & $AppCmd add backup "$IISBackupName" | Out-Null
        # Copy the IIS backup folder into our run folder for inclusion in the ZIP
        $IISBackupRoot   = Join-Path $env:windir 'System32\inetsrv\backup'
        $IISBackupSource = Join-Path $IISBackupRoot $IISBackupName
        if (Test-Path -LiteralPath $IISBackupSource) {
            Copy-Item -Path $IISBackupSource -Destination (Join-Path $IISFullBackupPath $IISBackupName) -Recurse -Force
            Log-Message "IIS full backup captured: $IISBackupSource -> $IISFullBackupPath" -Status SUCCESS
        } else {
            Log-Message "IIS backup reported created but folder not found: $IISBackupSource" -Status FAILED
        }
    } catch {
        Log-Message "Error creating IIS full backup: $_" -Status FAILED
    }

    # ===== Windows Services (rich CIM details) =====
    try {
        $ServicesFile = Join-Path $ServicesPath "Services_$TimeStamp.csv"
        Get-CimInstance -Class Win32_Service |
            Select-Object Name, DisplayName, Description, State, Status, StartMode, Started,
                          DelayedAutoStart, StartName, PathName, ProcessId, ServiceType,
                          ExitCode, ServiceSpecificExitCode, ErrorControl, TagId |
            Export-Csv -Path $ServicesFile -NoTypeInformation
        Log-Message "Services state (rich CIM) saved: $ServicesFile" -Status SUCCESS
    } catch {
        Log-Message "Error backing up Services state: $_" -Status FAILED
    }

    # ===== Web.config =====
    try {
        if (-not (Test-Path -LiteralPath $WebConfigSourcePath)) {
            throw "WebConfigSourcePath not found: $WebConfigSourcePath"
        }
        $WebConfigDest = Join-Path $WebConfigPath "Web_$TimeStamp.config"
        Copy-Item -Path $WebConfigSourcePath -Destination $WebConfigDest -Force
        Log-Message "Web.config copied to: $WebConfigDest" -Status SUCCESS
    } catch {
        Log-Message "Error backing up Web.config: $_" -Status FAILED
    }

    Log-Message "Backup capture completed." -Status INFO

    # ---------- Stop transcript to release file lock BEFORE compression ----------
    if ($TranscriptStarted) {
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        $TranscriptStarted = $false
        Log-Message "Transcript stopped prior to compression to release lock." -Status INFO
    }

    # ---------- Zip & checksum ----------
    $ZipFileName  = "BackedUpFiles_${HostName}_$TimeStamp.zip"
    $ZipFilePath  = Join-Path $env:TEMP $ZipFileName
    $HashFileName = "$ZipFileName.sha256"
    $HashFilePath = Join-Path $RunFolder $HashFileName

    try {
        if (Test-Path -LiteralPath $ZipFilePath) { Remove-Item -LiteralPath $ZipFilePath -Force }
        Compress-Archive -Path $RunFolder -DestinationPath $ZipFilePath -Force
        Log-Message "Run folder compressed: $ZipFilePath" -Status SUCCESS

        # Compute SHA-256 checksum and save as .sha256 file (format: "<hash> *filename")
        $zipHash = Get-FileHash -Path $ZipFilePath -Algorithm SHA256
        "{0} *{1}" -f $zipHash.Hash, $ZipFileName | Set-Content -Path $HashFilePath -Encoding ASCII
        Log-Message "ZIP checksum (SHA-256) computed: $($zipHash.Hash)" -Status SUCCESS
    } catch {
        Log-Message "Error compressing or hashing: $_" -Status FAILED
        throw
    }

    # ---------- Copy ZIP + checksum to network ----------
    try {
        Ensure-DirectoryExists -Path $NetworkDestinationDir
        $NetworkZipPath  = Join-Path $NetworkDestinationDir $ZipFileName
        $NetworkHashPath = Join-Path $NetworkDestinationDir $HashFileName
        Copy-Item -Path $ZipFilePath  -Destination $NetworkZipPath  -Force
        Copy-Item -Path $HashFilePath -Destination $NetworkHashPath -Force
        Log-Message "ZIP + checksum copied to network: $NetworkZipPath, $NetworkHashPath" -Status SUCCESS

        # Clean up temp ZIP; keep local .sha256 alongside run folder for audit (optional)
        Remove-Item -LiteralPath $ZipFilePath -Force
        Log-Message "Temporary ZIP removed: $ZipFilePath" -Status SUCCESS
    } catch {
        Log-Message "Error copying to network destination: $_" -Status FAILED
        throw
    }

    Log-Message "Backup run completed successfully. RunId=$RunId" -Status SUCCESS
}
catch {
    Log-Message "Backup run failed: $_" -Status FAILED
    throw
}
finally {
    if ($TranscriptStarted) {
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    }
}
