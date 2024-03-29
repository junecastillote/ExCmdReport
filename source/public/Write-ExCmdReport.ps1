Function Write-ExCmdReport {
    [CmdletBinding()]
    param (
        [parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline
        )]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        [parameter()]
        [string]
        $Organization,

        [parameter()]
        [int]
        $TruncateLongValue,

        [parameter()]
        [string]
        $ReportFile
    )
    Begin {

        Say '.........................................'
        Say "> I'm ready to start writing the report"
        Say '.........................................'

        #Region - Is Exchange Connected?
        if (!($Organization)) {
            Say "> You did not specify the name of the organization. That probably means you want me to get it for you instead."
            try {
                Say "> Ok then, I'm trying to get your organization name now. Seriously, you ask too much."
                $Organization = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
                Say "> Found it! Your organization name is $($Organization)"
            }
            catch [System.Management.Automation.CommandNotFoundException] {
                Say "> It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first before asking me to do stuff for you."
                Say "> Or you can just specify your organization name next time so that I don't have to look for it for you. The parameter is -Organization <organization name>."
                return $null
            }
            catch {
                Say "> Something is wrong. You can see the error below. I can't tell you how to fix it, but you should fix it before asking me to stuff for you."
                Say "> Or you can just specify your organization name next time so that I don't have to look for it for you. The parameter is -Organization <organization name>."
                Say $_.Exception.Message
                return $null
            }
        }

        #EndRegion

        if ($ReportFile) {
            New-Item -ItemType File -Path $ReportFile -Force | Out-Null
        }

        # For use later to determine the oldest and newest entry
        $dateCollection = @()

        $ModuleInfo = Get-Module ExCmdReport
        $tz = ([System.TimeZoneInfo]::Local).DisplayName.ToString().Split(" ")[0]
        $today = Get-Date -Format "MMMM dd, yyyy hh:mm tt"
        $css = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\public\style.css') -Raw
        $title = "Exchange Admin Audit Log Report - $($Organization) - $($today)"

        $i = 1
    }

    Process {
        foreach ($item in $InputObject) {
            if ($item.CallerAdminName) {
                # $Caller = $item.CallerAdminName
                $Caller = "$($item.CallerAdminName) [$($item.Caller)]"
            }
            else {
                $Caller = $item.Caller
            }
            $dateCollection += $item.RunDate
            # $html2 += '<tr><td>' + $i + '<td><b>Date: </b>' + (Get-Date $item.RunDate -Format "MMM-dd-yyyy hh:mm:ss tt") + '<br><b>Caller: </b>' + $Caller + '<br><b>Target: </b>' + $item.ObjectModified + '<br><b>Succeeded: </b>' + $item.Succeeded + '</td>'
            $html2 += '<tr><td><b>Date: </b>' + (Get-Date $item.RunDate -Format "MMM-dd-yyyy hh:mm:ss tt") + '<br><b>Caller: </b>' + $Caller + '<br><b>Target: </b>' + $item.ObjectModified + '<br><b>Succeeded: </b>' + $item.Succeeded + '</td>'
            $html2 += '<td><b>' + $item.CmdLetName + '</b><br><br>'
            foreach ($param in $item.CmdletParameters) {
                if ($TruncateLongValue) {
                    if ($param.Value.length -gt $TruncateLongValue) {
                        $paramValue = ((($param.Value).ToString().SubString(0, $TruncateLongValue)) + "...")
                    }
                    else {
                        $paramValue = $param.Value
                    }
                }
                else {
                    $paramValue = $param.Value
                }
                # $html2 += ('<b>' + $param.Name + '</b>' + ' = ' + $paramValue + '<br>')
                $html2 += ('<b>' + $param.Name + ':</b> ' + $paramValue + '<br>')
            }
            $html2 += '</td></tr>'
            $i = $i + 1
        }
    }
    End {

        $dateCollection = $dateCollection | Sort-Object
        $startDate = $dateCollection[0]
        $endDate = $dateCollection[-1]
        Say "> Your report covers the period of $($startDate) to $($endDate)"
        Say "> I am creating your HTML report now... in my memory."
        #$html1 = @()
        $html1 += '<html><head><title>' + $title + '</title>'
        $html1 += '<style type="text/css">'
        $html1 += $css
        $html1 += '</style></head>'
        $html1 += '<body>'
        $html1 += '<table id="tbl">'
        $html1 += '<tr><td class="head"> </td></tr>'
        $html1 += '<tr><th class="section">Exchange Admin Command Audit Report</th></tr>'
        $html1 += '<tr><td class="head"><b>' + $Organization + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
        $html1 += '<tr><td class="head"> </td></tr>'
        $html1 += '</table>'
        $html1 += '<table id="tbl">'
        $html1 += '<tr><td><b>Run Details:</b> (' + "$( Get-Date $startDate -Format "MMM-dd-yyyy HH:mm:ss") - $( Get-Date $endDate -Format "MMM-dd-yyyy HH:mm:ss")" + ')</td><td><b>Commands and Parameters</b></td></tr>'
        $html3 += '</table>'
        $html3 += '<table id="tbl">'
        $html3 += '<tr><td class="head"> </td></tr>'
        $html3 += '<tr><td class="head"> </td></tr>'
        $html3 += '<tr><td class="head"><a href="' + $ModuleInfo.ProjectURI.ToString() + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </td></a><br>'
        $html3 += '<tr><td class="head"> </td></tr>'
        $html3 += '</body></html>'

        $htmlBody = ($html1 + $html2 + $html3) -join "`n"
        if ($ReportFile) {
            try {
                $htmlBody | Out-File $ReportFile -Encoding UTF8 -Force
                Say "> I saved the HTML report to a file, your majesty."
                Say "> You can find the report at $((Resolve-Path $ReportFile).Path)."
                # return $htmlBody
            }
            catch {
                Say "> Something is wrong. You can see the error below. Because of it I cannot save your report to file. Fix it."
                Say $_.Exception.Message
                return $null
            }
        }
        else {
            Say "> I've created the report object for you, which is basically just an HTML code in my memory."
            Say "> If you wanted to save the report to an HTML file, you should use the -ReportFile <path to report.html> parameter."
            Say "> Or, you can just pipe the report out to file like ' | Out-File report.html'. But you should already know how to do that."
            return $htmlBody
        }
    }
}