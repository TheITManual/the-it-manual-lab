### Script to Place an Exchange Server into Maintenance Mode
## ---------------------------------------------------------

<#
.SYNOPSIS
    Script to Place an Exchange Server into Maintenance Mode

.DESCRIPTION
    Gracefully drains mail flow, redirects queued messages, and places an Exchange
    Server into full maintenance mode to allow patching/updates without disruption.

.AUTHOR
    Created by	: TheITManual
	Version		: 1.0
#>

# ---------------------------
# Define Exchange server variables once
# ---------------------------
$ServerFQDN = "<ExchangeServer_FQDN>"   # server being patched
$TargetFQDN = "<TargetServer_FQDN>"     # healthy server to receive redirected queues

# ---------------------------
# Step 1: Begin Transport Drain
# ---------------------------
Set-ServerComponentState -Identity $ServerFQDN -Component HubTransport -State Draining -Requester Maintenance

# ---------------------------
# Step 2: Verify HubTransport Component State
# ---------------------------
Get-ServerComponentState -Identity $ServerFQDN | Where-Object { $_.Component -eq 'HubTransport' }

# ---------------------------
# Step 3: Restart Transport Service on the target server - makes sure the server stops accepting new mail.
# (Restart-Service is local-only; invoke on the remote server)
# ---------------------------
Invoke-Command -ComputerName $ServerFQDN -ScriptBlock { Restart-Service -Name MSExchangeTransport -Force }

# ---------------------------
# Step 4: Re-Verify HubTransport State
# ---------------------------
Get-ServerComponentState -Identity $ServerFQDN | Where-Object { $_.Component -eq 'HubTransport' }

# ---------------------------
# Step 5: Redirect Queued Messages to another server - makes sure any old mail is evacuated
# ---------------------------
Redirect-Message -Server $ServerFQDN -Target $TargetFQDN

# ---------------------------
# Step 6 (Optional but recommended): Check Queue Status on the server
# ---------------------------
Get-Queue -Server $ServerFQDN | Select-Object Identity, Status, MessageCount, NextHopDomain

# (Optional) Wait for queues to drain to zero (or a small threshold) before offline
# $timeoutSeconds = 900
# $pollSeconds    = 10
# $deadline = (Get-Date).AddSeconds($timeoutSeconds)
# do {
#     $sum = (Get-Queue -Server $ServerFQDN | Measure-Object -Property MessageCount -Sum).Sum
#     Write-Host "Current queued messages on $ServerFQDN: $sum"
#     if ($sum -le 0) { break }
#     Start-Sleep -Seconds $pollSeconds
# } while (Get-Date -lt $deadline)

# ---------------------------
# Step 7: Put Server into Maintenance Mode (server-wide offline)
# ---------------------------
Set-ServerComponentState -Identity $ServerFQDN -Component ServerWideOffline -State Inactive -Requester Maintenance

# ---------------------------
# Step 8: Verify server component states
# ---------------------------
Get-ServerComponentState -Identity $ServerFQDN | Sort-Object Component | Format-Table Component, State, Requester -AutoSize
