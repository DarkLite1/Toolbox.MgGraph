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
            Helper function for Send-MgUserMail to create a list of
            e-mail addresses that can be used with ToRecipients,
            CcRecipients and BccRecipients.

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