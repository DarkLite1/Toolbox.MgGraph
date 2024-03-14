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
            $actual.EmailAddress.Keys.ForEach({$_ |
                Should -BeExactly 'Address'})
        }
        It 'with the correct quantity of hashtables' {
            $actual.Count | Should -Be $params.MailAddress.Count
        }
        It 'of type array' {
            $actual.getType().BaseType.Name | Should -BeExactly 'Array'
        }
    }
}