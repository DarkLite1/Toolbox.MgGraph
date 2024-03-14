#Requires -Version 7
#Requires -Modules Microsoft.Graph.Authentication

try {
    $params = @{
        ClientId              = $env:AZURE_CLIENT_ID
        TenantId              = $env:AZURE_TENANT_ID
        CertificateThumbprint = $env:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT
        NoWelcome             = $true
        ErrorAction           = 'Stop'
    }
    Connect-MgGraph @params
}
catch {
    throw "Failed authenticating to MS Graph: $_"
}

Function ConvertTo-MgUserMailRecipientHC {
    <#
        .SYNOPSIS
            Create a list of e-mail addresses to use with Send-MgUserMail:
            - -BodyParameter 'Message.ToRecipients'
            - -BodyParameter 'Message.CcRecipients'
            - -BodyParameter 'Message.BccRecipients'

        .PARAMETER MailAddress
            One or more SMTP mail addresses.

        .EXAMPLE
            $params = @{
                MailAddress = 'bob@conotoso.com', 'mike@conotoso.com'
            }
            ConvertTo-MgUserMailRecipientHC @params
    #>

    [CmdLetBinding()]
    [OutputType([hashtable[]])]
    Param(
        [Parameter(Mandatory)]
        [String[]]$MailAddress
    )

    $result = @()

    $MailAddress.ForEach(
        {
            $result += @{EmailAddress = @{Address = $_ } }
        }
    )

    , $result
}

Function ConvertTo-MgUserMailAttachmentHC {
    <#
        .SYNOPSIS
            Create a list of e-mail attachments to use with Send-MgUserMail:
            - BodyParameter 'Message.Attachments'

        .PARAMETER Path
            Full path to the file.

        .EXAMPLE
            $params = @{
                Path = 'c:\Temp\file.txt'
            }
            ConvertTo-MgUserMailAttachmentHC @params
    #>

    [CmdLetBinding()]
    [OutputType([hashtable[]])]
    Param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path -Path $_ -PathType 'Leaf' })]
        [String[]]$Path
    )

    $result = @()

    $Path.ForEach(
        {
            $result += @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                Name          = ($_ -split '\\')[-1]
                ContentBytes  = [Convert]::ToBase64String(
                    [IO.File]::ReadAllBytes($_)
                )
            }
        }
    )

    , $result
}