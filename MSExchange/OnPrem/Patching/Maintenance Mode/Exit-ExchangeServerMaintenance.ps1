### Script to Bring an Exchange Server Out of Maintenance Mode
## -----------------------------------------------------------

<#
.SYNOPSIS
    Script to bring an Exchange Server out of maintenance mode.

.DESCRIPTION
    Reverses the maintenance process: reactivates components, ensures transport
    is restarted, and confirms the server is healthy and ready to handle mail flow again.

.AUTHOR
    Created by	: TheITManual
	Version		: 1.0
#>

# ---------------------------
# Define Exchange server variables once
# ---------------------------
$ServerFQDN = "<ExchangeServer_FQDN>"   # server being brought back online

# ---------------------------
# Step 1: Reactivate Server-Wide Component
# ---------------------------
# Takes the server out of ServerWideOffline so all roles can run again.
Set-ServerComponentState -Identity $ServerFQDN -Component ServerWideOffline -State Active -Requester Maintenance

# ---------------------------
# Step 2: Reactivate Hub Transport
# ---------------------------
# Allows this server to begin accepting and processing new mail.
Set-ServerComponentState -Identity $ServerFQDN -Component HubTransport -State Active -Requester Maintenance

# ---------------------------
# Step 3: Restart Transport Service on the target server
# ---------------------------
# Ensures Transport reloads its configuration and starts processing mail normally.
Invoke-Command -ComputerName $ServerFQDN -ScriptBlock { Restart-Service -Name MSExchangeTransport -Force }

# ---------------------------
# Step 4: Verify Components Are Active
# ---------------------------
# Confirms HubTransport and other components are marked "Active".
Get-ServerComponentState -Identity $ServerFQDN | Sort-Object Component | Format-Table Component, State, Requester -AutoSize

# ---------------------------
# Step 5: Verify Queue Status (Optional)
# ---------------------------
# Ensures server is now processing and routing messages correctly.
Get-Queue -Server $ServerFQDN | Select-Object Identity, Status, MessageCount, NextHopDomain
