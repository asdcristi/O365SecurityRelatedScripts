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

<# Create and store a secure password and an AESKey on the device where the script will be added:

$AESKeyFile = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKeyFile)
$AESKeyFile | Out-File C:\temp\AESKey.key

$PasswordFile = "C:\temp\Password.txt"
$Key = Get-Content -Path C:\temp\AESKey.key
$Credential = ConvertTo-SecureString -String 'MSTest123' -AsPlainText -Force
$EncryptedCredential = ConvertFrom-SecureString -SecureString $Credential -Key $Key | Set-Content -Path C:\temp\Password.txt `
-OutVariable $EncryptedCredential

#>

# Retrieve Credentials stored and opening connection to the Exchange Online PowerShell Endpoint
$AdminUser = "admin@mstcrrad.onmicrosoft.com"
$EncryptedCredential = Get-Content -Path C:\temp\Password.txt
$Key = Get-Content C:\temp\AESKey.key
$SecureString = ConvertTo-SecureString -String $EncryptedCredential -Key $Key
$MyCredentials = New-Object System.Management.Automation.PSCredential $AdminUser, $SecureString

Write-Host "Retrieving Credentials.." -ForegroundColor Green

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $MyCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

# Retrieve a list of all users with AuditEnabled set to false
Write-Host "Retrieving list of Mailboxes with AuditEnabled $false.." -ForegroundColor Green
$Users = Get-Mailbox -ResultSize Unlimited -Filter {AuditEnabled -eq $false}
$UserCount = 0

# Enable mailbox auditing and setting/increasing audited activities
Write-Host "Enabling Mailbox Audit for all Users/SharedMailboxes retrieved.." -ForegroundColor Green

foreach ($User in $Users) 

	{    
	    Set-Mailbox -Identity $User.UserPrincipalName -AuditEnabled $true `
        -AuditAdmin Copy, Create, FolderBind, HardDelete, Move, MoveToDeletedItems, SendAs, SendOnBehalf,`
                    SoftDelete, Update, UpdateInboxRules, UpdateFolderPermissions, UpdateCalendarDelegation `
        -AuditDelegate Create, FolderBind, HardDelete, Move, MoveToDeletedItems, SendAs, `
                       SendOnBehalf, SoftDelete, Update, UpdateFolderPermissions, UpdateInboxRules `
        -AuditOwner Create, HardDelete, MailboxLogin, Move, MoveToDeletedItems, SoftDelete, Update, `
                    UpdateFolderPermissions, UpdateInboxRules, UpdateCalendarDelegation `
        -ErrorAction SilentlyContinue

	    Write-Host "Enabling Mailbox Audit for $User" -ForegroundColor Green
        
        if (++$UserCount % 50 -eq 0)
        
           {

           Write-Host "Sleeping 15 seconds.."
           Start-Sleep -Seconds 15

           }

    }

Send-MailMessage -SmtpServer "outlook.office365.com" -From $AdminUser -To $AdminUser -Subject "AuditEnableTaskScheduler Job" `
-Body "Hello, the task has Run." -UseSsl -Credential $MyCredentials

Write-Host "Operation Completed.." -ForegroundColor Green
Write-Host "Closing PSSession.." -ForegroundColor Green

Remove-PSSession $Session

