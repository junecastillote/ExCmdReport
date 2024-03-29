Function Send-PSGraphMail {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)]
        [string]$Token,

        [parameter(Mandatory)]
        [string]$From,

        [parameter(Mandatory)]
        [string[]]$To,

        [parameter()]
        [string[]]$CC,

        [parameter()]
        [string[]]$BCC,

        [parameter(Mandatory)]
        [string]$Subject,

        [parameter(Mandatory)]
        [string]$Body,

        [parameter()]
        [string[]]$Attachments,

        [parameter()]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0'
    )

    Function ConvertRecipientsToJSON {
        param(
            [Parameter(Mandatory)]
            [string[]]
            $Recipients
        )
        $jsonRecipients = @()
        $Recipients | ForEach-Object {
            $jsonRecipients += @{EmailAddress = @{Address = $_ } }
        }
        return $jsonRecipients
    }

    $mailBody = @{
        message = @{
            subject                = $Subject
            body                   = @{
                content     = $Body
                contentType = "HTML"
            }
            internetMessageHeaders = @(
                @{
                    name  = "X-Mailer"
                    value = "PsGraphMail by June Castillote"
                }
            )
            attachments            = @()
        }
    }

    # To recipients
    $mailBody.message += @{
        toRecipients = @(
            $(ConvertRecipientsToJSON $To)
        )
    }

    # Cc recipients
    if ($CC) {
        $mailBody.message += @{
            ccRecipients = @(
                $(ConvertRecipientsToJSON $CC)
            )
        }
    }

    # BCC recipients
    if ($BCC) {
        $mailBody.message += @{
            ccRecipients = @(
                $(ConvertRecipientsToJSON $BCC)
            )
        }
    }

    if ($Attachments) {
        foreach ($file in $Attachments) {
            try {
                $filename = (Resolve-Path $file -ErrorAction STOP).Path

                if ($PSVersionTable.PSEdition -eq 'Core') {
                    $fileByte = $([convert]::ToBase64String((Get-Content $filename -AsByteStream)))
                }
                else {
                    $fileByte = $([convert]::ToBase64String((Get-Content $filename -Raw -Encoding byte)))
                }

                $mailBody.message.attachments += @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    "name"         = $(Split-Path $filename -Leaf)
                    "contentBytes" = $fileByte
                }
            }
            catch {
                "Attachment: $($_.Exception.Message)" | Out-Default
            }
        }
    }

    $mailBody = $mailBody | ConvertTo-Json -Depth 4
    # $mailBody
    $mailApiUri = "https://graph.microsoft.com/$ApiVersion/users/$($From)/sendmail"
    $headerParams = @{'Authorization' = "Bearer $($Token)" }

    Invoke-RestMethod -Method Post -Uri $mailApiUri -Body $mailbody -Headers $headerParams -ContentType 'application/json' -ErrorAction STOP
}