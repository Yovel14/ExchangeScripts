<#
.SYNOPSIS
	Gets average Mails size Based on sent mails.
.DESCRIPTION
	Gets average Mails size Based on sent mails.
.EXAMPLE
	PS> .\Get-AverageMailsSize.ps1 -StartDate (Get-Date).AddDays(-3)
.LINK
	https://github.com/Yovel14/ExchangeScripts
.NOTES
	Author: Yovel
#>
[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
	[datetime]$StartDate,

	[Parameter(Mandatory=$false)]
	[datetime]$EndDate = (Get-Date),

	[Parameter(Mandatory=$false)]
	[int]$JobParallelLimit = 4
)

if($StartDate -gt $EndDate)
{
    Throw "StartDate is After EndDate"
}

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
        Start-Sleep -Seconds 5
    }
}

# Actual Code
$serverList = Get-MailboxServer | select -ExpandProperty name | sort
$messageFilter = '($_.Source -eq "STOREDRIVER") -and ($_.Recipients -notlike "HealthMailbox*") -and ($_.Recipients -notlike "extest_*")  -and ($_.DisplayName -notlike "HealthMailbox*") -and ($_.DisplayName -notlike "extest_*")'

# Recipients
Write-Host -ForegroundColor Cyan "Starting Recipients data Collection"
foreach($server in $serverList)
{
    Start-Job -ScriptBlock { Add-PSSnapin *xch*; (Get-MessageTrackingLog -Server $Using:Server -ResultSize:Unlimited -Start $Using:StartDate -End $Using:EndDate  -EventID Deliver | Where-Object ([ScriptBlock]::Create($Using:messageFilter))) | measure -Property TotalBytes -Average}
    Wait-ForJobsParallelLimit
}

Wait-ForAllJobs
$AverageData = get-job | Receive-Job
Get-Job | Remove-Job
$countSum = ($AverageData | measure -Property count -Sum).Sum
$Average = ($AverageData | %{$_.count * $_.Average} | measure -Sum).Sum / $countSum
$Average | Out-File ".\AverageMailSize-Bytes$(($StartDate-$EndDate).days)Day.txt"

Write-Host -ForegroundColor Cyan "Done Recipients data Collection"