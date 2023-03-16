###########################################################################################
# DISCLAIMER:                                                                             #
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT               #
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY              #
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT      #
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR    #
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS    #
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR        #
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE  #
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS #
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)  #
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,      #
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES                   #
###########################################################################################

#Pre-Requisites:

<# Create and store a secure password and an AESKey on the device where the script will be added. Uncommend this section, `
add your password in the $credential row and run the cmdlets:

$AESKeyFile = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKeyFile)
$AESKeyFile | Out-File C:\temp\AESKey.key
$passwordFile = "C:\temp\Password.txt"
$key = Get-Content -Path C:\temp\AESKey.key
$credential = ConvertTo-SecureString -String 'PasswordHere' -AsPlainText -Force
$encryptedCredential = ConvertFrom-SecureString -SecureString $credential -Key $Key | Set-Content -Path C:\temp\Password.txt `
-OutVariable $encryptedCredential
#>

# Retrieve Credentials stored and opening connection to the Exchange Online PowerShell Endpoint
# Note: $adminUser needs to be defined and paths need to be adjusted according to the AESKey and Password file:
Start-Transcript -Path C:\temp\EnableMailboxAuditingTaskSchedulerLogs\TaskSchedulerRun-$((Get-Date).ToString('dd-MM-yyyy_HH-mm')).txt
Write-Host "Starting transcript.." -ForegroundColor Yellow

try {
    Write-Host "Attempting credential retrieval.." -ForegroundColor Yellow

    $adminUser = "user@contoso.com"
    $encryptedCredential = Get-Content -Path C:\temp\Password.txt -ErrorAction SilentlyContinue
    $key = Get-Content C:\temp\AESKey.key -ErrorAction SilentlyContinue
    $secureString = ConvertTo-SecureString -String $encryptedCredential -Key $key -ErrorAction SilentlyContinue
    $myCredentials = New-Object System.Management.Automation.PSCredential $adminUser, $secureString -ErrorAction SilentlyContinue
}
catch {
    Write-Host "Credential retrieval error encountered. Error message Below:" -ForegroundColor Red
    $Error[0].Exception | Format-List -f *
    $Error[0].Exception.SerializedRemoteException | Format-List -f
    Exit
    Stop-Transcript
}

Write-Host "Credential retrieval successful." -ForegroundColor Green

# Checking Exchange Online module installation
Write-Host "Checking ExchangeOnlineManagement module installation or attempting install.." -ForegroundColor Yellow
$requiredModule = "ExchangeOnlineManagement"

foreach ($i in $requiredModule) {

    try {
        Import-Module -Name $i -ErrorAction Stop -WarningAction SilentlyContinue
    }

    catch [System.IO.FileNotFoundException] {

        try {
            Install-Module -Name $i -Scope CurrentUser -ErrorAction Stop
        }

        catch {
            Write-Host "Install Module failed, error message below:" -ForegroundColor Red
            $Error[0].Exception | Format-List -f *
            Stop-Transcript
            Exit
        }

        Import-Module -Name $i -ErrorAction Stop -WarningAction SilentlyContinue
    }
}

Write-Host "Import of module successful." -ForegroundColor Green

# Connection to Exchange Online
Write-Host "Attempting Exchange Online Endpoint connection.." -ForegroundColor Yellow

try {
    Connect-ExchangeOnline -Credential $myCredentials -ShowBanner:$false
}    
catch {
    Write-Host "Connection to the Exchange Endpoint failed with error message below:" -ForegroundColor Red
    $Error[0].Exception | Format-List -f *
    Stop-Transcript
    Exit
}

Write-Host "Connection to the Exchange Endpoint was Successful." -ForegroundColor Green

# Checking audit configuration at organization level and enabling it if disabled
Write-Host "Checking if Audit is Enabled at Org level.." -ForegroundColor Yellow
$organizationAdminAudit = Get-AdminAuditLogConfig

if ($organizationAdminAudit.UnifiedAuditLogIngestionEnabled -eq $False) {
    Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
    Write-Host "Auditing was disabled and now enabled by script." -ForegroundColor Red
}
else {
    Write-Host "Auditing is Enabled." -ForegroundColor Green
}

# Checking mailbox audit configuration at organization level and enabling it if disabled
Write-Host "Checking if Mailbox Auditing is Enabled at Org level.." -ForegroundColor Yellow
$organizationAudit = Get-OrganizationConfig

if ($organizationAudit.AuditDisabled -eq $true) {
    Set-OrganizationConfig -AuditDisabled $false
    Write-Host "Mailbox Auditing was disabled and now enabled by script." -ForegroundColor Red
}
else {
    Write-Host "Mailbox Auditing is Enabled." -ForegroundColor Green
}

# Retrieve all users with mailbox auditing disabled and enabling it
$mailboxAuditDisabledUsers = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.AuditEnabled -eq $false } | Select-Object Name, AuditEnabled
Write-Host ("Mailbox Auditing is disabled for the following users: " + $mailboxAuditDisabledUsers.Name) -ForegroundColor Yellow

foreach ($i in $mailboxAuditDisabledUsers) {
    Write-Host ("Enabling Mailbox Auditing for " + $i.Name) -ForegroundColor Green
    Set-Mailbox $i.Name -AuditEnabled $true
}

# Retrieve all users with Audit Bypass enabled and disabling it
$mailboxAuditBypassEnabledUsers = Get-MailboxAuditBypassAssociation -ResultSize Unlimited | Where-Object { $_.AuditBypassEnabled -eq $true } | Select-Object Name, AuditBypassEnabled
Write-Host ("Mailbox Audit Bypass is enabled for the following users: " + $mailboxAuditBypassEnabledUsers.Name) -ForegroundColor Yellow

foreach ($i in $mailboxAuditBypassEnabledUsers) {
    Write-Host ("Disabling Mailbox Audit Bypass for " + $i.Name) -ForegroundColor Green
    Set-MailboxAuditBypassAssociation $i.Name -AuditBypassEnabled $false
}

Write-Host "Mailbox Audit Bypass Disabled for all users" -ForegroundColor Green

# Increasing mailbox auditing log level to maximum
Write-Host "Increasing mailbox auditing log level to maximum.." -ForegroundColor Yellow

# Defining log level to maximum (E5 Users)
$auditAdminActionsE5 = New-Object -TypeName 'System.Collections.ArrayList';
$auditAdminActionsE5 = "Update", "Copy", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "FolderBind", "SendAs", `
                       "SendOnBehalf", "Create", "UpdateFolderPermissions", "UpdateInboxRules", "UpdateCalendarDelegation", `
                       "RecordDelete", "ApplyRecord", "MailItemsAccessed", "Send";

$auditDelegateActionsE5 = New-Object -TypeName 'System.Collections.ArrayList';
$auditDelegateActionsE5 = "Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "FolderBind", "SendAs", `
                          "SendOnBehalf", "Create", "UpdateFolderPermissions", "UpdateInboxRules", "RecordDelete", `
                          "ApplyRecord", "MailItemsAccessed";

$auditOwnerActionsE5 = New-Object -TypeName 'System.Collections.ArrayList';
$auditOwnerActionsE5 = "Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "Create", "MailboxLogin", `
                       "UpdateFolderPermissions", "UpdateInboxRules", "UpdateCalendarDelegation", "RecordDelete", `
                       "ApplyRecord", "MailItemsAccessed", "Send", "SearchQueryInitiated";

# Defining log level to maximum (non E5 Users)
$auditAdminActions = New-Object -TypeName 'System.Collections.ArrayList';
$auditAdminActions = "Update", "Copy", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "FolderBind", "SendAs", `
                     "SendOnBehalf", "MessageBind", "Create", "UpdateFolderPermissions", "UpdateInboxRules", `
                     "UpdateCalendarDelegation", "RecordDelete", "ApplyRecord";

$auditDelegateActions = New-Object -TypeName 'System.Collections.ArrayList';
$auditDelegateActions = "Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "FolderBind", "SendAs", `
                        "SendOnBehalf", "Create", "UpdateFolderPermissions", "UpdateInboxRules", "RecordDelete", `
                        "ApplyRecord";

$auditOwnerActions = New-Object -TypeName 'System.Collections.ArrayList';
$auditOwnerActions = "Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "Create", "MailboxLogin", `
                     "UpdateFolderPermissions", "UpdateInboxRules", "UpdateCalendarDelegation", "RecordDelete", `
                     "ApplyRecord";

# Retrieving all users
$users = Get-Mailbox -ResultSize Unlimited
$userCount = 0

# Increasing Mailbox Audit Actions to maximum foreach user
foreach ($i in $users) {

    $outputAdminE5 = Compare-Object -ReferenceObject $auditAdminActionsE5 -DifferenceObject $i.AuditAdmin
    $outputDelegateE5 = Compare-Object -ReferenceObject $auditDelegateActionsE5 -DifferenceObject $i.AuditDelegate
    $outputOwnerE5 = Compare-Object -ReferenceObject $auditOwnerActionsE5 -DifferenceObject $i.AuditOwner

    $outputAdmin = Compare-Object -ReferenceObject $auditAdminActions -DifferenceObject $i.AuditAdmin
    $outputDelegate = Compare-Object -ReferenceObject $auditDelegateActions -DifferenceObject $i.AuditDelegate
    $outputOwner = Compare-Object -ReferenceObject $auditOwnerActions -DifferenceObject $i.AuditOwner
    
    if ($null -ne $outputAdminE5) {
        
        try {
            Write-Host "Extending Admin Auditing Actions for user $i to E5 level.." -ForegroundColor Yellow
            Set-Mailbox -Identity $i.UserPrincipalName -AuditAdmin @{Add = $auditAdminActionsE5 } -ErrorAction Stop
            $successA = "Extending Admin Auditing Actions for user $i to E5 level successful."
        }
        catch {
            
            if ($null -ne $outputAdmin) {

                try {
                    Write-Host "Unable to extend Admin Auditing Actions for user $i to E5 level, attempting non E5.." -ForegroundColor Yellow
                    Set-Mailbox -Identity $i.UserPrincipalName -AuditAdmin @{Add = $auditAdminActions } -ErrorAction Stop
                    $successA = "Extending Admin Auditing Actions for user $i to non-E5 level successful."
                }
                catch {
                    Write-Host "Unable to extend Admin Auditing Actions for user $i due to the error below:" -ForegroundColor Red
                    $Error[0].Exception | Format-List -f *
                    $successB = "Extending Admin Auditing Actions for user $i not successful."
                }
            }
            else {
                $successC = "Unable to extend Admin Audit Actions for user $i as these are already Extended to the maximum, for non-E5 level."
            }
            
        }
        if ($null -ne $successA) {

            Write-Host $successA -ForegroundColor Green

            if ($null -ne $successB) {

                Write-Host $successB -ForegroundColor Red
            }
        }
        else {

            Write-Host $successC -ForegroundColor Green
        }
    }
    else {
        Write-Host "Unable to extend Admin Audit Actions for user $i as these are already Extended to the maximum, for E5 level." -ForegroundColor Green
    }

    if ($null -ne $outputDelegateE5) {
        
        try {
            Write-Host "Extending Delegate Auditing Actions for user $i to E5 level.." -ForegroundColor Yellow
            Set-Mailbox -Identity $i.UserPrincipalName -AuditDelegate @{Add = $auditDelegateActionsE5 } -ErrorAction Stop
            $successD = "Extending Delegate Auditing Actions for user $i to E5 level successful."
        }
        catch {

            if ($null -ne $outputDelegate) {

                try {
                    Write-Host "Unable to extend Delegate Auditing Actions for user $i to E5 level, attempting non E5.." -ForegroundColor Yellow
                    Set-Mailbox -Identity $i.UserPrincipalName -AuditDelegate @{Add = $auditDelegateActions } -ErrorAction Stop
                    $successD = "Extending Delegate Auditing Actions for user $i to non-E5 level successful."
                }
                catch {
                    Write-Host "Unable to extend Delegate Auditing Actions for user $i due to the error below:" -ForegroundColor Red
                    $Error[0].Exception | Format-List -f *
                    $successE = "Extending Delegate Auditing Actions for user $i not successful."
                }
            }
            else {
                $successF = "Unable to extend Delegate Audit Actions for user $i as these are already Extended to the maximum, for non-E5 level."
            } 
        }

        if ($null -ne $successD) {

            Write-Host $successD -ForegroundColor Green

            if ($null -ne $successE) {

                Write-Host $successE -ForegroundColor Red
            }
        }
        else {

            Write-Host $successF -ForegroundColor Green
        }
    }
    else {
        Write-Host "Unable to extend Delegate Audit Actions for user $i as these are already Extended to the maximum, for E5 level." -ForegroundColor Green
    }

    if ($null -ne $outputOwnerE5) {
        
        try {
            Write-Host "Extending Owner Auditing Actions for user $i to E5 level.." -ForegroundColor Yellow
            Set-Mailbox -Identity $i.UserPrincipalName -AuditOwner @{Add = $auditOwnerActionsE5 } -ErrorAction Stop
            $successG = "Extending Owner Auditing Actions for user $i to E5 level successful."
        }
        catch {

            if ($null -ne $outputOwner) {

                try {
                    Write-Host "Unable to extend Owner Auditing Actions for user $i to E5 level, attempting non E5.." -ForegroundColor Yellow
                    Set-Mailbox -Identity $i.UserPrincipalName -AuditOwner @{Add = $auditOwnerActions } -ErrorAction Stop
                    $successG = "Extending Owner Auditing Actions for user $i to non-E5 level successful."
                }
                catch {
                    Write-Host "Unable to extend Owner Auditing Actions for user $i due to the error below:" -ForegroundColor Red
                    $Error[0].Exception | Format-List -f *
                    $successH = "Extending Delegate Auditing Actions for user $i not successful."
                }
            }
            else {
                $successI = "Unable to extend Owner Audit Actions for user $i as these are already Extended to the maximum, for non-E5 level."
            }
        }

        if ($null -ne $successG) {

            Write-Host $successG -ForegroundColor Green

            if ($null -ne $successH) {

                Write-Host $successH -ForegroundColor Red
            }
        }
        else {

            Write-Host $successI -ForegroundColor Green
        }    
    }
    else {
        Write-Host "Unable to extend Owner Audit Actions for user $i as these are already Extended to the maximum, for E5 level." -ForegroundColor Green
    }

    if (++$UserCount % 5 -eq 0) {

        Write-Host "Sleeping 15 seconds to avoid throttling.." -ForegroundColor Magenta
        Start-Sleep -Seconds 15
    }
}

Write-Host "Operation Completed." -ForegroundColor Green
Write-Host "Closing PSSession." -ForegroundColor Green
Get-PSSession | Remove-PSSession
Stop-Transcript
