<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Start-Transcript -Path "c:\temp\AddUsersToDistroList $(get-date -f yyyy-MM-dd-mm).txt" -Verbose
Add-Type -AssemblyName PresentationFramework


Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.MessageBox]::Show('Sign on using Admin O365 credentials.')
#Signing on to O365 using the Admin Credentials for O365
try{
$O365_Credentials  = Get-Credential } catch {exit}
# Importing necessary modules
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $O365_Credentials -Authentication Basic –AllowRedirection         
Import-PSSession $Session
Connect-MsolService -Credential $O365_Credentials
$groups = $null
$users = $null

$groups = Get-DistributionGroup | select Displayname,PrimarySmtpAddress | sort DisplayName
$users = get-mailbox | select Displayname | sort DisplayName

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,400'
$Form.text                       = "Form"
$Form.TopMost                    = $false

##Users
$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 193
$ListBox1.height                 = 368
$ListBox1.location               = New-Object System.Drawing.Point(10,20)
Foreach($i in $users) {
    $Listbox1.items.add($i.DisplayName) | out-Null}
$ListBox1.SelectionMode = 'MultiExtended'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(247,361)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK


##Groups
$ListBox2                        = New-Object system.Windows.Forms.ListBox
$ListBox2.text                   = "listBox"
$ListBox2.width                  = 172
$ListBox2.height                 = 331
$ListBox2.location               = New-Object System.Drawing.Point(215,21)
foreach($g in $groups) {$ListBox2.Items.Add($g.DisplayName) | Out-Null}


$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.text                   = "User"
$TextBox1.width                  = 100
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(48,3)
$TextBox1.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.text                   = "Group"
$TextBox2.width                  = 100
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(254,2)
$TextBox2.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Cancel"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(317,362)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel


$Form.controls.AddRange(@($ListBox1,$Button1,$ListBox2,$TextBox1,$TextBox2,$Button2))


$result = $form.ShowDialog() 

if ($result -eq "Cancel") { 
    Write-Host -ForegroundColor Green "Tidying Up"
    Remove-PSSession $Session
    
    Exit 
    }
 if ($result -eq "Ok")
    {Write-Host "OK selected - proceeding with adding group memberships"}


$selectedUsers = $ListBox1.SelectedItems
$selectedGroup = $ListBox2.SelectedItem


if ($selectedGroup -eq $null) {
    [System.Windows.MessageBox]::Show('No group selected')
    $result = $form.ShowDialog() 
}

##Before changes
write-host "before changes are made"
Get-DistributionGroupMember -Identity $selectedGroup | select name,primarysmtpaddress
write-host "above are in the group."

foreach($s in $selectedUsers) {
        Add-DistributionGroupMember -identity $Listbox2.selecteditem -Member "$($s)"
}

##After changes
write-host "after changes are made"
Get-DistributionGroupMember -Identity $selectedGroup | select name,primarysmtpaddress
write-host "above are in the group."


remove-pssession $Session
Stop-Transcript

