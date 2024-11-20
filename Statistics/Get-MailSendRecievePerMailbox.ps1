<#
.SYNOPSIS
	Gets Amount Of mails Each Mailbox Sends And Receives and outputs each to CSV files.
.DESCRIPTION
	Gets Amount Of mails Each Mailbox Sends And Receives and outputs each to CSV files.
.EXAMPLE
	PS> .\Get-MailSendRecievePerMailbox.ps1 -StartDate (Get-Date).AddDays(-5) -EndDate (Get-Date).AddMinutes(-30)
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
	[datetime]$endDate = (Get-Date),

	[Parameter(Mandatory=$false)]
	[String]$MailsReceiveCSVPath = "MailsReceive",

	[Parameter(Mandatory=$false)]
	[String]$MailsSendCSVPath = "MailsSend",

	[Parameter(Mandatory=$false)]
	[int]$JobParallelLimit = 4
)

if($StartDate -gt $EndDate)
{
    Throw "StartDate is After EndDate"
}

if(Test-Path -Path "$MailsReceiveCSVPath-$(($endDate-$StartDate).days)Days.csv")
{
    Throw "MailsReceiveCSVPath file Already Exists '$MailsReceiveCSVPath-$(($endDate-$StartDate).days)Days.csv'"
}

if(Test-Path -Path "$MailsSendCSVPath-$(($endDate-$StartDate).days)Days.csv")
{
    Throw "MailsSendCSVPath file Already Exists '$MailsSendCSVPath-$(($endDate-$StartDate).days)Days.csv'"
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
    Write-Host -ForegroundColor Cyan "Starting on server: $server"
    Start-Job -ScriptBlock { Add-PSSnapin *xch*; (Get-MessageTrackingLog -Server $Using:Server -ResultSize:Unlimited -Start $Using:StartDate -End $Using:endDate  -EventID Deliver | Where-Object ([ScriptBlock]::Create($Using:messageFilter))).Recipients | group -NoElement}
    Wait-ForJobsParallelLimit
}

Wait-ForAllJobs
$RecipientsData = get-job | Receive-Job
Get-Job | Remove-Job
$calculatedGroup = $RecipientsData | group -Property Name | %{[PSCustomObject](@{name = $_.name; count = (($_.group | measure -Sum -Property count)).sum})}
$calculatedGroup | Export-Csv -NoTypeInformation -path "$MailsReceiveCSVPath-$(($endDate-$StartDate).days)Days.csv"
$RecipientsData = $null
$calculatedGroup = $null

Write-Host -ForegroundColor Cyan "Done Recipients data Collection"

# Sender
Write-Host -ForegroundColor Cyan "Starting Senders data Collection"

foreach($server in $serverList)
{
    Write-Host -ForegroundColor Cyan "Starting on server: $server"
    Start-Job -ScriptBlock { Add-PSSnapin *xch*; (Get-MessageTrackingLog -Server $Using:Server -ResultSize:Unlimited -Start $Using:StartDate -End $Using:endDate -EventID Receive | Where-Object ([ScriptBlock]::Create($Using:messageFilter))).Sender | group -NoElement}
    Wait-ForJobsParallelLimit
}

Wait-ForAllJobs
$SendersData = get-job | Receive-Job
Get-Job | Remove-Job
$calculatedGroup = $SendersData | group -Property Name | %{[PSCustomObject](@{name = $_.name; count = (($_.group | measure -Sum -Property count)).sum})}
$calculatedGroup | Export-Csv -NoTypeInformation -Path "$MailsSendCSVPath-$(($endDate-$StartDate).days)Days.csv"
$SendersData = $null
$calculatedGroup = $null


Write-Host -ForegroundColor Cyan "Done Senders data Collection"
