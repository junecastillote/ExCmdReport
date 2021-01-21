# Write-ExCmdReport

The `Write-ExCmdReport` takes the output of the `Get-ExCmdLog` function as input. It then creates an HTML report. It can also send the report by email if specified.

```PowerShell
Write-ExCmdReport
    [-InputObject] <Object>
    [[-Organization] <string>]
    [[-TruncateLongValue] <int>]
    [[-ReportFile] <string>]
    [<CommonParameters>]
```

## Write-ExCmdReport (Example 1)

```PowerShell
Get-ExCmdLog | Write-ExCmdReport -ReportFile C:\\Temp\\auditlogreport.html
}
```

In this example, `Write-ExCmdReport` accepts the output of `Get-ExCmdLog` as a pipeline input and saves the report to *`C:\\Temp\auditlogreport.html`*.

## Write-ExCmdReport (Example 2)

```PowerShell
$logs = Get-ExCmdLog
Write-ExCmdReport -InputObject $logs -ReportFile C:\\Temp\\auditlogreport.html
}
```

In this example, the output of the `Get-ExCmdLog` is stored in the `$logs` variable. Then the `$logs` is passed to `Write-ExCmdReport` as the value of the parameter `-InputObject`.

## Write-ExCmdReport (Example 3)

```PowerShell
$logs = Get-ExCmdLog
Write-ExCmdReport -InputObject $logs -ReportFile C:\\Temp\\auditlogreport.html -SendEmail:$true -From sender@domain.com -To recipient@domain.com -smtpServer relay.domain.com -port 25
```

In this example, the output of the `Get-ExCmdLog` is stored in the `$logs` variable. Then the `$logs` is passed to `Write-ExCmdReport` as the value of the parameter `-InputObject`. Then the report is sent by email.

## Write-ExCmdReport Parameters

### -InputObject

The output of the `Get-ExCmdLog` function.

```yaml
Type: PSCustomObject
Required: True
Default value: None
Accept pipeline input: True
```

### -ReportFile

The location where the HTML report will be saved. If not specified, the file will be saved to the default location.

```yaml
Type: String
Required: False
Default value: "C:\\Windows\\Temp\\ExCmdReport_MM-dd-yyyy.html"
Accept pipeline input: False
```

### -TruncateLongValue

Some outputs have a large number of values. Use this is you want to limit the nubmer of characters displayed on the report.

```yaml
Type: Int32
Required: False
Default value: None
Accept pipeline input: False
```

### -Organization

The organization name you want to be displayed in the report. If not specified, the default value will be displayed or the script will try to query it from Exchange.

```yaml
Type: String
Required: False
Default value: None
Accept pipeline input: False
```

### -SendEmail

Switch to indicate to send the report by email.

```yaml
Type: Switch
Required: False
Default value: None
Accept pipeline input: False
```

### -From

The sender email address of the report. Required if `-SendEmail` is used.

> IMPORTANT: If you're using Office 365 SMTP Relay, the sender address must be an actual mailbox. Additionally, the credential used must have *Send As* permission to the mailbox.

```yaml
Type: mailaddress
Required: False
Default value: None
Accept pipeline input: False
```

### -To

The recipient email addresses of the report. Required if `-SendEmail` is used.

```yaml
Type: mailaddress
Required: False
Default value: None
Accept pipeline input: False
```

### -Subject

The email subject of the report. If not specified, the default value will be used.

```yaml
Type: String
Required: False
Default value: "Exchange Admin Audit Log Report"
Accept pipeline input: False
```

### -SmtpServer

This is the SMTP relay to be used for sending the email report. Required if `-SendEmail` is used.

```yaml
Type: string
Required: False
Default value: None
Accept pipeline input: False
```

### -Port

The Port number of the SMTP relay server. Required if `-SendEmail` is used.

```yaml
Type: Int32
Required: False
Default value: 25
Accept pipeline input: False
```

### -Credential

Use this if the SMTP relay server requires authentication.

```yaml
Type: PSCredential
Required: False
Default value: None
Accept pipeline input: False
```

### -UseSSL

Use if the SMTP relay server requires SSL/TLS.

```yaml
Type: Switch
Required: False
Default value: False
Accept pipeline input: False
```
