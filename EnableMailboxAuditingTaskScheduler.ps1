###########################################################################################
# DISCLAIMER:                                                                             #
#                                                                                         #
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

#The purpose of this script is to automatize the enablement of MailboxAuditing, which can be added to TaskScheduler.

# Connection to the Exchange Online PowerShell Endpoint
$SecPassword = ConvertTo-SecureString "InsertPasswordHere" -AsPlainText -Force
$MyCredentials = New-Object System.Management.Automation.PSCredential ("admin@domain.com", $SecPassword)

Write-Host "Retrieving Credentials.."

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

# Retrieve a list of all users with AuditEnabled set to false
Write-Host "Retrieving list of Users/SharedMailboxes with AuditEnabled $false.."
$Users = Get-Mailbox -ResultSize Unlimited -Filter {AuditEnabled -eq $false} | Select UserPrincipalName,AuditEnabled,AuditDelegate
$UserCount = 0

# Enable mailbox auditing and setting/increasing audited activities, while setting the AuditLogAgeLimit to 180.
Write-Host "Enabling Mailbox Audit for all Users/SharedMailboxes retrieved.."
foreach ($User in $Users) 
	{    
	    Set-Mailbox -Identity $User.UserPrincipalName -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Copy,Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update,UpdateInboxRules,UpdateFolderPermissions,UpdateCalendarDelegation -AuditDelegate Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update,UpdateFolderPermissions,UpdateInboxRules -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update,UpdateFolderPermissions,UpdateInboxRules,UpdateCalendarDelegation -ErrorAction SilentlyContinue
	    Write-Host "Enabling Mailbox Audit for" $user.UserPrincipalName

        if (++$UserCount % 50 -eq 0)

       {

       Write-Host "Sleeping 15 seconds.."
       Start-Sleep -Seconds 15

       }

    }
Write-Host "Operation Completed.." -ForegroundColor Green
