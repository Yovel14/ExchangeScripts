<#
.SYNOPSIS
	Gets Mailbox Size Quota Deviation from DB Defaults, and outputs to a CSV file.
.DESCRIPTION
	Gets Mailbox Size Quota Deviation from DB Defaults, and outputs to a CSV file.
.EXAMPLE
	PS> .\Get-MailboxQuotaDeviationFromDBDefaults.ps1 -OutputCSVPath .\MailboxQuotaDeviations.csv
.LINK
	https://github.com/Yovel14/ExchangeScripts
.NOTES
	Author: Yovel
#>
[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
    [ValidateScript({if(Test-Path -Path $_){ Throw "File $_ already Exists"}})]
	[String]$OutputCSVPath
)

Add-PSSnapin *xch*

# could be made faster using jobs
Get-Mailbox -Filter 'UseDatabaseQuotaDefaults -eq $false' -ResultSize unlimited | Group-Object -NoElement -Property ProhibitSendReceiveQuota | %{[PSCustomObject]@{Count = $_.Count; Size = $_.Name}} | Export-Csv -Path $OutputCSVPath -NoTypeInformation
