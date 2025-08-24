### Script to Check Exchange Server Build Versions
## ------------------------------------------------

<#
.SYNOPSIS
    Script to collect Exchange Server build/version information.

.DESCRIPTION
    This script queries all Exchange servers in the organization, locates the
    ExSetup.exe binary on each, and retrieves its file/product version details.
    This provides an accurate view of installed Exchange builds across the org.

.AUTHOR
	Created by	: TheITManual
	Version		: 1.0
#>

# ---------------------------
# Step 1: Get Exchange servers
# ---------------------------
# Collects a sorted list of all Exchange servers in the organization.
$servers = Get-ExchangeServer | Sort-Object Name

# ---------------------------
# Step 2: Define script block
# ---------------------------
# This block runs on each remote server to locate ExSetup.exe and read version info.
$sb = {
    # Initialize candidate paths
    $candidates = @()

    # (a) Prefer environment variable (always correct drive/path if set)
    if ($env:ExchangeInstallPath) {
        $candidates += (Join-Path $env:ExchangeInstallPath 'Bin\ExSetup.exe')
    }

    # (b) Add common default install paths as fallback
    # V15 = Exchange 2013/2016/2019
    # V14 = Exchange 2010
    $candidates += @(
        'C:\Program Files\Microsoft\Exchange Server\V15\Bin\ExSetup.exe',
        'C:\Program Files\Microsoft\Exchange Server\V14\Bin\ExSetup.exe'
    )

    # Step 3: Check each path until one is found
    foreach ($p in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $p) {
            $fi  = Get-Item -LiteralPath $p
            $fvi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fi.FullName)

            # Return useful version details
            return [pscustomobject]@{
                Path           = $fi.FullName
                FileVersion    = $fvi.FileVersion
                ProductVersion = $fvi.ProductVersion
                ProductName    = $fvi.ProductName
                CompanyName    = $fvi.CompanyName
            }
        }
    }

    # Warn if ExSetup.exe not found
    Write-Warning "ExSetup.exe not found"
}

# ---------------------------
# Step 4: Run remotely
# ---------------------------
# Executes the script block against each Exchange server, continuing on errors.
$results = Invoke-Command -ComputerName ($servers | ForEach-Object Name) -ScriptBlock $sb -ErrorAction Continue

# ---------------------------
# Step 5: Display results
# ---------------------------
# Show key details in a readable table.
$results |
    Select-Object PSComputerName, FileVersion, ProductVersion, Path |
    Format-Table -AutoSize
