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
Get-ExCmdLog | Write-ExCmdReport -ReportFile C:\Temp\auditlogreport.html
}
```

In this example, `Write-ExCmdReport` accepts the output of `Get-ExCmdLog` as a pipeline input and saves the report to *`C:\\Temp\auditlogreport.html`*.

## Write-ExCmdReport (Example 2)

```PowerShell
$logs = Get-ExCmdLog
Write-ExCmdReport -InputObject $logs -ReportFile C:\Temp\auditlogreport.html
}
```

In this example, the output of the `Get-ExCmdLog` is stored in the `$logs` variable. Then the `$logs` is passed to `Write-ExCmdReport` as the value of the parameter `-InputObject`.

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

The location where the HTML report will be saved. Alternatively, you can just use `Out-File`.

```yaml
Type: String
Required: False
Default value: None
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

The organization name you want to be displayed in the report. If not specified, the script will try to get the organization display name from Exchange.

```yaml
Type: String
Required: False
Default value: None
Accept pipeline input: False
```
