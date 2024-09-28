<#
.SYNOPSIS
	Gets Information about disconnected mailboxes.
.DESCRIPTION
	Gets Information about disconnected mailboxes. (Amount of Disconnected mailboxes, Average Size, Average Mailbox Count per DB)
.EXAMPLE
	PS> .\Get-DisconnectedMailboxInformation.ps1
.LINK
	https://github.com/Yovel14/ExchangeScripts
.NOTES
	Author: Yovel
#>
Add-PSSnapin *xch*

$dbs = Get-MailboxDatabase | select -ExpandProperty Name | sort
$totalAmount = 0
$mailboxSizeAverage = 0
foreach($db in $dbs)
{
    Write-Host "Starting on DB: $db" -ForegroundColor Cyan
    $dbd = Get-MailboxStatistics -Database $db -Filter {DisconnectDate -ne $null}
    $mailboxSizeAverage += ($dbd.TotalItemSize.value | measure -Average).Average
    $totalAmount +=$dbd.Count
    Write-Host "Done on DB: $db" -ForegroundColor Cyan
}

Write-Host "Amount of Disconnected mailboxes: $totalAmount"
Write-Host "average mailbox size: $(($mailboxSizeAverage/$dbs.Count)/1GB)GB"
Write-Host "Average users per database: $($totalAmount/$($dbs.Count))"