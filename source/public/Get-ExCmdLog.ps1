Function Get-ExCmdLog {
	<#
    .SYNOPSIS
        Get Exchange Online admin audit log events
    .DESCRIPTION

    .EXAMPLE
        Get-ExCmdLog -searchParamHash @{StartDate='10/10/2018';EndDate='10/12/2018';ExternalAccess=$false} -InformationAction Continue

        The example show that you created a hashtable parameter of StartDate, EndDate, and ExternalAccess
    #>
	[cmdletbinding()]
	param (
		<#
            Hashtable of valid parameters for the cmdLet Search-AdminAuditLog.
            eg. @{StartDate='10/10/2018';EndDate='10/12/2018';ExternalAccess=$false}
        #>
		[parameter()]
		[hashtable]$searchParamHash = @{
			StartDate      = ((Get-Date).AddHours(-24))
			EndDate        = ((Get-Date))
			ExternalAccess = $false
			IsSuccess      = $true
		},
		[parameter()]
		[switch]$resolveAdminName
	)

	#Region - Is Exchange Connected?
	try {
		$null = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
	}
	catch [System.Management.Automation.CommandNotFoundException] {
		Say "> It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first before asking me to do stuff for you."
		return $null
	}
	catch {
		Say "> Something is wrong. You can see the error below. You should fix it before asking me to try again."
		Say $_.Exception.Message
		return $null
	}
	#EndRegion

	Say "........................................."
	Say "> I'm using the search parameters you gave me:" | Out-Null
	foreach ($i in $searchParamHash.Keys) {
		Say "> $($i) = $($searchParamHash.Item($i))"
	}
	Say ">........................................."
	$startIndex = 0
	Say "> Starting my search from index #$($startIndex+1)"
	$searchParamHash += @{StartIndex = $startIndex }

	$auditLogs = Search-AdminAuditLog @searchParamHash

	# Terminate script if there are no audit logs to report
	if (!$auditLogs) {
		Say "> The admin audit log between the time range you specified is empty. There's nothing to report."
		return $null
	}

	if ($auditLogs.count -eq 1000) {
		Do {
			$startIndex += 1000
			$searchParamHash.StartIndex = $startIndex
			Say "> Now I'm searching from index #$($startIndex+1)"
			$temp = Search-AdminAuditLog @searchParamHash
			if ($temp.count -gt 0) {
				$auditLogs += $temp
			}
		}
		While ($temp.count -eq 1000)
	}
	$auditLogs | Add-Member -MemberType NoteProperty -Name CallerAdminName -Value $null
	Say "> I found a total of $($auditLogs.count) events in the [unfiltered] audit log."

	if ($resolveAdminName) {
		Say "> Please wait while I try to put names on the callers' login. Remember, you asked me to do this."
		Say "> ........................................."
		$uniqueCallerAdminName = @()

		# Get unique caller id from the result
		$uniqueCaller = $auditLogs | Select-Object caller -Unique

		# Create a hashtable of Caller and CallerAdminName
		$uniqueCaller | ForEach-Object {
			try {
				$CallerAdminName = (Get-User ($_.Caller) -ErrorAction Stop).DisplayName
			}
			catch {
				$CallerAdminName = $null
			}

			$callerObj = New-Object psobject -Property @{
				Caller          = $_.Caller
				CallerAdminName = $CallerAdminName
			}
			$uniqueCallerAdminName += $callerObj
		}

		foreach ($i in $auditLogs) {
			$i.CallerAdminName = ($uniqueCallerAdminName | Where-Object { $_.Caller -eq $i.Caller }).CallerAdminName
		}
	}
	# return the results
	return $auditLogs
}