<#  
.SYNOPSIS  
    Provides users with permissions to mailboxes 
.DESCRIPTION  
    This script was written to facilitate adding users permissions to mailboxes etc. 
.NOTES  
    File Name  : mailboxpermissions.ps1 
    Author     : Markus G - markusgustavsson21@gmail.com
    Requires   : O365 admin rights. 
.LINK  
    
#> 


Start-Transcript -Path "c:\temp\AddusersMailboxPermissions $(get-date -f yyyy-MM-dd-mm).txt" 
Add-Type -AssemblyName PresentationFramework
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.MessageBox]::Show('Sign on using Admin O365 credentials.')
#Signing on to O365 using the Admin Credentials for O365
$O365_Credentials  = Get-Credential 
# Importing necessary modules
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $O365_Credentials -Authentication Basic -AllowRedirection         
Import-PSSession $Session


$UsersList = Get-Mailbox -ResultSize unlimited 
$mailboxesList = Get-Mailbox -resultsize unlimited

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '765,510'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.text                   = "User"
$TextBox1.width                  = 100
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(48,3)
$TextBox1.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.text                   = "Mailbox to access"
$TextBox2.width                  = 150
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(304,2)
$TextBox2.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$Users                           = New-Object system.Windows.Forms.ListBox
$Users.text                      = "Users"
$Users.width                     = 217
$Users.height                    = 456
$Users.location                  = New-Object System.Drawing.Point(12,35)
foreach($u in $UsersList) {
  $Users.Items.Add($u.DisplayName) | Out-Null
}
$Users.SelectionMode = 'MultiExtended'

$Mailboxes                       = New-Object system.Windows.Forms.ListBox
$Mailboxes.text                  = "Mailboxes"
$Mailboxes.width                 = 221
$Mailboxes.height                = 455
foreach($m in $mailboxesList) {
  $mailboxes.Items.Add($m.DisplayName) | Out-Null
}
$Mailboxes.location              = New-Object System.Drawing.Point(261,35)

$SendAsRights                    = New-Object system.Windows.Forms.CheckBox
$SendAsRights.text               = "Send As Rights"
$SendAsRights.AutoSize           = $false
$SendAsRights.width              = 195
$SendAsRights.height             = 20
$SendAsRights.location           = New-Object System.Drawing.Point(514,40)
$SendAsRights.Font               = 'Microsoft Sans Serif,10'

$SendOnBehalf                    = New-Object system.Windows.Forms.CheckBox
$SendOnBehalf.text               = "Send On Behalf Rights"
$SendOnBehalf.AutoSize           = $false
$SendOnBehalf.width              = 195
$SendOnBehalf.height             = 20
$SendOnBehalf.location           = New-Object System.Drawing.Point(514,62)
$SendOnBehalf.Font               = 'Microsoft Sans Serif,10'

$fullAccess                      = New-Object system.Windows.Forms.CheckBox
$fullAccess.text                 = "Full Access"
$fullAccess.AutoSize             = $false
$fullAccess.width                = 195
$fullAccess.height               = 20
$fullAccess.location             = New-Object System.Drawing.Point(514,82)
$fullAccess.Font                 = 'Microsoft Sans Serif,10'

$Ok                              = New-Object system.Windows.Forms.Button
$Ok.text                         = "OK"
$Ok.width                        = 60
$Ok.height                       = 30
$Ok.location                     = New-Object System.Drawing.Point(514,128)
$Ok.Font                         = 'Microsoft Sans Serif,10'
$Ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

$Cancel                          = New-Object system.Windows.Forms.Button
$Cancel.text                     = "Cancel"
$Cancel.width                    = 60
$Cancel.height                   = 30
$Cancel.location                 = New-Object System.Drawing.Point(608,128)
$Cancel.Font                     = 'Microsoft Sans Serif,10'
$Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

$Form.controls.AddRange(@($Textbox1,$Textbox2,$Users,$Mailboxes,$SendAsRights,$SendOnBehalf,$fullAccess,$Ok,$Cancel))

$result = $form.ShowDialog()

if ($result -eq "Cancel") { 
    Write-Host -ForegroundColor Green "Tidying Up"
    Remove-PSSession $Session
    Exit 
    }
  if ($result -eq "Ok")
    {Write-Host "OK selected - proceeding with adding permissions"}
else {exit}
<##>
## form checking before submission
if($fullAccess.CheckState -eq "Unchecked" -and $SendOnBehalf.CheckState -eq "Unchecked" -and $SendAsRights.CheckState -eq "Unchecked") {
  [System.Windows.MessageBox]::Show('No rights selected')
  $result = $form.ShowDialog()
} 

if($Users.SelectedItems.Count -eq 0) {
  [System.Windows.MessageBox]::Show('No users selected.')
  $result = $form.ShowDialog()
}

if($mailboxes.SelectedItems.Count -eq 0) {
  [System.Windows.MessageBox]::Show('No Mailboxes selected.')
  $result = $form.ShowDialog()
}

###used for error testing - not necessary for script to run
write-host "fullaccess state: $($fullAccess.CheckState)"
write-host "send on behalf state: $($SendOnBehalf.CheckState)"
write-host "send as state: $($SendAsRights.CheckState)"

## Changes below
if ($fullaccess.CheckState -eq "Checked") {
  write-host "Fullaccess checked!"
  foreach($p in $users.SelectedItems){
    Add-MailboxPermission -identity "$($mailboxes.SelectedItem)" -User "$($p)" -AccessRights fullaccess -Verbose
    write-host "$($p) is now given full access to $($mailboxes.SelectedItem) mailbox!"
  }
  Get-MailboxPermission -identity "$($mailboxes.SelectedItem)"
}
if ($SendAsRights.CheckState -eq "Checked") {
  Write-Host "Send as status checked!"
  foreach($u in $users.SelectedItems) {
    Add-RecipientPermission -Identity "$($mailboxes.SelectedItems)" -Trustee "$($u)" -AccessRights sendas -Confirm:$false -Verbose
    write-host "$($u) has been given send as rights for $($mailboxes.SelectedItem)" 
  }
  Get-RecipientPermission -identity "$($mailboxes.SelectedItem)"
}
if ($SendOnBehalf.CheckState -eq "Checked") {
  Write-Host "Send on behalf selected!"
  foreach($q in $users.SelectedItems) {
    Set-Mailbox -identity "$($mailboxes.SelectedItem)" -GrantSendOnBehalfTo "$($q)" -Verbose
    write-host "$($q) has been given send on behalf rights to $($mailboxes.SelectedItem)."
  }
}
