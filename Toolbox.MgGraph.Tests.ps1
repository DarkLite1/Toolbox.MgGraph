#Requires -Version 7
#Requires -Modules Pester

BeforeDiscovery {
    # used by inModuleScope
    $testModule = $PSCommandPath.Replace('.Tests.ps1', '.psm1')
    $testModuleName = $testModule.Split('\')[-1].TrimEnd('.psm1')

    Remove-Module $testModuleName -Force -Verbose:$false -EA Ignore
    Import-Module $testModule -Force -Verbose:$false
}
Describe 'ConvertTo-MgUserMailRecipientHC' {
    Context 'create a hashtable for one e-mail address' {
        BeforeAll {
            $params = @{
                MailAddress = 'bob@conotoso.com'
            }
            $actual = ConvertTo-MgUserMailRecipientHC @params
        }
        It 'with EmailAddress.Address' {
            $actual.EmailAddress.Address | Should -BeExactly $params.MailAddress
        }
        It 'with the correct quantity of hashtables' {
            $actual.Count | Should -Be $params.MailAddress.Count
        }
        It 'of type array' {
            $actual.getType().BaseType.Name | Should -BeExactly 'Array'
        }
    }
    Context 'create a hashtable for multiple e-mail addresses' {
        BeforeAll {
            $params = @{
                MailAddress = 'bob@conotoso.com', 'mike@conotoso.com'
            }
            $actual = ConvertTo-MgUserMailRecipientHC @params
        }
        It 'with EmailAddress.Address' {
            $params.MailAddress.foreach(
                {
                    $actual.EmailAddress.Values |
                    Should -Contain $_
                }
            )
            $actual.EmailAddress.Keys.ForEach({ $_ |
                    Should -BeExactly 'Address' })
        }
        It 'with the correct quantity of hashtables' {
            $actual.Count | Should -Be $params.MailAddress.Count
        }
        It 'of type array' {
            $actual.getType().BaseType.Name | Should -BeExactly 'Array'
        }
    }
}
Describe 'ConvertTo-MgUserMailAttachmentHC' {
    Context 'create a hashtable for one attachment' {
        BeforeAll {
            $params = @{
                Path = (New-Item -Path 'TestDrive:\file1.txt' -ItemType File).FullName
            }
            $actual = ConvertTo-MgUserMailAttachmentHC @params
        }
        It 'with the correct properties' {
            $actual.'@odata.type' |
            Should -BeExactly '#microsoft.graph.fileAttachment'
            $actual.Name | Should -BeExactly 'file1.txt'
            $actual.ContentBytes | Should -BeExactly (
                [Convert]::ToBase64String(
                    [IO.File]::ReadAllBytes($params.Path)
                )
            )
        }
        It 'with the correct quantity of hashtables' {
            $actual.Count | Should -Be $params.Path.Count
        }
        It 'of type array' {
            $actual.getType().BaseType.Name | Should -BeExactly 'Array'
        }
    }
    Context 'create a hashtable for multiple attachments' {
        BeforeAll {
            $params = @{
                Path = @(
                    (New-Item -Path 'TestDrive:\file2.txt' -ItemType File).FullName
                    (New-Item -Path 'TestDrive:\file3.txt' -ItemType File).FullName
                )
            }
            $actual = ConvertTo-MgUserMailAttachmentHC @params
        }
        It 'with the correct properties' {
            foreach ($path in $params.Path) {
                $testActual = $actual.Where(
                    { $_.Name -eq ($path -split '\\')[-1] }
                )

                $testActual.'@odata.type' |
                Should -BeExactly '#microsoft.graph.fileAttachment'
                $testActual.Name | Should -BeExactly ($path -split '\\')[-1]
                $testActual.ContentBytes | Should -BeExactly (
                    [Convert]::ToBase64String(
                        [IO.File]::ReadAllBytes($path)
                    )
                )
            }
        }
        It 'with the correct quantity of hashtables' {
            $actual.Count | Should -Be $params.Path.Count
        }
        It 'of type array' {
            $actual.getType().BaseType.Name | Should -BeExactly 'Array'
        }
    }
}