# Release Notes for ExCmdReport

## Version 1.1

- Required ExchangeOnlineManagement module version 2.0.3.
- Removed the custom functions below. It's up to you how you want to connect to Exchange Online or On-Premises.
  - `Connect-ExOL` - used to connect to Exchange Online
  - `Connect-ExOP` - used to connect to Exchange On-Premises
- Removed option to send the report by email from the `Write-ExCmdReport` function. This gives you control on how you want to send the HTML report by email. (eg. `Send-MailMessage`, `Graph API`, etc..)
