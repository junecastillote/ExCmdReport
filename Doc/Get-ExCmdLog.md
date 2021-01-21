# Get-ExCmdLog

The `Get-ExCmdLog` accepts a hashtable of parameters used in [`Search-AdminAuditLog`](https://docs.microsoft.com/en-us/powershell/module/exchange/policy-and-compliance-audit/search-adminauditlog?view=exchange-ps#syntax)

```PowerShell
Get-ExCmdLog
    [[-searchParamHash] <hashtable>]
    [-resolveAdminName]
    [<CommonParameters>]
```

## Get-ExCmdLog (Example 1)

```PowerShell
Get-ExCmdLog -searchParamHash @{
    StartDate      = ((Get-Date).AddHours(-24))
    EndDate        = ((Get-Date))
}
```

This example retrieves all admin audit log entries from within the last 24 hours.

## Get-ExCmdLog (Example 2)

```PowerShell
Get-ExCmdLog -searchParamHash @{
    StartDate      = ((Get-Date).AddHours(-24))
    EndDate        = ((Get-Date))
    CmdLets        = 'Set-Mailbox'
}
```

This example retrieves all admin audit log entries from within the last 24 hours that macthes the CmdLet `Set-Mailbox`.

## Get-ExCmdLog Parameters

### -searchParamHash

A hashtable of valid parameters from the `Search-AdminAuditLog` cmdlet.

```yaml
Type: Hashtable
Required: True
Default value: @{StartDate = ((Get-Date).AddHours(-24)); EndDate = ((Get-Date))}
Accept pipeline input: False
```

### -resolveAdminName

Instructs the function to resolve the Caller values from UPN to Name. (eg. bob@Contoso.com to 'Robert Parr')

```yaml
Type: switch
Required: False
Default value: None
Accept pipeline input: False
```
