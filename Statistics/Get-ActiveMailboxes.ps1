<#
.SYNOPSIS
	Gets Active Mailboxes Based on Input Date and outputs to a file.
.DESCRIPTION
	Gets Active Mailboxes Based on Input Date and outputs to a file.
.EXAMPLE
	PS> .\Get-ActiveMailboxes.ps1 -LastMailboxLoginDate (Get-Date).AddMonths(-3) -OutputCSVPath .\ActiveMailboxes.csv
.LINK
	https://github.com/Yovel14/ExchangeScripts
.NOTES
	Author: Yovel
#>
[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
	[datetime]$LastMailboxLoginDate,

	[Parameter(Mandatory=$true)]
    [ValidateScript({if(Test-Path -Path $_){ Throw "File $_ already Exists"}})]
	[String]$OutputCSVPath,

	[Parameter(Mandatory=$false)]
	[int]$JobParallelLimit = 2
)
Add-PSSnapin *xch*

# Functions
function Wait-ForAllJobs
{
    while((Get-Job -state Running).count -ne 0)
    {
        $null = (Get-Job -state Running)[0] | Wait-Job
    }
}

function Wait-ForJobsParallelLimit
{
    while((Get-Job -state Running).count -ge $JobParallelLimit)
    {
        Start-Sleep -Seconds 1
    }
}

# Actual Code
$DAGs = Get-DatabaseAvailabilityGroup | select -ExpandProperty Name | sort
$dbs = Get-MailboxDatabase | select -ExpandProperty Name | sort

foreach($dag in $DAGs)
{
    Write-Host "Starting on DAG: $dag" -ForegroundColor Cyan
    $null = Start-Job -name $dag -ScriptBlock {Add-PSSnapin *xch*; Get-MailboxDatabase | ?{$_.MasterServerOrAvailabilityGroup -eq $Using:dag} | %{Get-MailboxStatistics -Database $_.Name -Filter {lastlogontime -gt $Using:LastMailboxLoginDate} | %{[pscustomobject]@{MailboxGuid = $_.MailboxGuid.ToString(); LegacyDN = $_.LegacyDN.ToString(); TotalItemSizeMB = $_.totalitemsize.value.toMB(); lastlogontime = $_.lastlogontime.tostring()}}}}
    Wait-ForJobsParallelLimit
}

Wait-ForAllJobs

$ActiveMailboxes = Get-Job | Receive-Job | %{[pscustomobject]@{MailboxGuid = $_.MailboxGuid; LegacyDN = $_.LegacyDN; TotalItemSizeMB = $_.TotalItemSizeMB; lastlogontime = $_.lastlogontime}} # another convertion to PSCustomobject is done because receiveing data from job adds Properties like runspaceID which are not needed and the reason for not filtering out of the job is to save ram

Get-Job | Remove-Job -Force

$ActiveMailboxes | Export-Csv -Path $OutputCSVPath -NoTypeInformation