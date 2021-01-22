# Import
Remove-Module ExCmdReport -ErrorAction SilentlyContinue
Import-Module '.\ExCmdReport.psd1'

$exoCredential = Import-CliXml -Path c:\temp\cred.xml
$tenantName = 'poshlab.onmicrosoft.com'
$organization = 'PowerShell Lab'
$today = Get-Date -Format "yyyy-MM-dd"
$ReportFile = "$($env:windir)\temp\$($tenantName)_ExCmdReport_$($today).html"

# Kill all existing Exchange PowerShell session to start fresh.
Get-PSSession -Name "Exchange*" | ForEach-Object {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}
# Clear the console
Clear-Host
Connect-ExchangeOnline -Credential $exoCredential -Organization $tenantName -ShowBanner:$false

$searchParamHash = [ordered]@{
    StartDate      = ((Get-Date).AddDays(-10))
    EndDate        = ((Get-Date))
    ExternalAccess = $false
}

$reportHash = @{
    TruncateLongValue = 50
    ReportFile   = $ReportFile
    organization = $Organization
}

$events = Get-ExCmdLog -searchParamHash $searchParamHash -InformationAction Continue
$events | Where-Object { $_.Caller -ne 'NT AUTHORITY\SYSTEM (w3wp)' } |
Write-ExCmdReport @reportHash -InformationAction Continue

$messageSplat = @{
    From       = "Mailer <Office365Mailer@poshlab.ml>"
    To         = "june@poshlab.ml"
    Subject    = "Exchange Admin Audit Log Report - $organization - $today"
    BodyAsHtml = $true
    Body       = Get-Content $ReportFile -raw
    SmtpServer = "smtp.office365.com"
    Port       = 587
    UseSSL     = $true
    Credential = $exoCredential
}
Send-MailMessage @messageSplat

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null