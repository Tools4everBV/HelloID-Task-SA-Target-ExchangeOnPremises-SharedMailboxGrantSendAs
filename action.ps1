# HelloID-Task-SA-Target-ExchangeOnPremises-SharedMailboxGrantSendAs
####################################################################
# Form mapping
$formObject = @{
    DisplayName     = $form.DisplayName
    MailboxIdentity = $form.MailboxIdentity
    UsersToAdd      = [array]$form.UsersToAdd
}

[bool]$IsConnected = $false
try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Add-ADPermission'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToAdd) {
        Write-Information "Executing ExchangeOnPremises action: [SharedMailboxGrantSendAs] for: [$($formObject.DisplayName)]"

        $null = Add-ADPermission -Identity $formObject.MailboxIdentity -AccessRights ExtendedRight -ExtendedRights "Send As" -Confirm:$false -User $user.UserPrincipalName -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnPremises'
            TargetIdentifier  = $formObject.MailboxIdentity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnPremises action: [SharedMailboxGrantSendAs][$($user.UserPrincipalName)] for: [$($formObject.DisplayName)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnPremises action: [SharedMailboxGrantSendAs][$($user.UserPrincipalName)] for: [$($formObject.DisplayName)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute ExchangeOnPremises action: [SharedMailboxGrantSendAs] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [SharedMailboxGrantSendAs] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
####################################################################
