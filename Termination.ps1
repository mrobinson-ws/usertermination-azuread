#Requires -Modules AzureAD, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell

#Declarations
[CmdletBinding()]
Param()
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
$quitboxOutput = ""

##### Start Main Loop #####
#Start While Loop for Quitbox
while ($quitboxOutput -ne "NO"){
    # Test And Connect To AzureAD If Needed
    try {
        Write-Verbose -Message "Testing connection to Azure AD"
        Get-AzureAdDomain -ErrorAction Stop | Out-Null
        Write-Verbose -Message "Already connected to Azure AD"
    }
    catch {
        Write-Verbose -Message "Connecting to Azure AD"
        Connect-AzureAD
    }

    #Test And Connect To Microsoft Exchange Online If Needed
    try {
        Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
        Get-Mailbox -ErrorAction Stop | Out-Null
        Write-Verbose -Message "Already connected to Microsoft Exchange Online"
    }
    catch {
        Write-Verbose -Message "Connecting to Microsoft Exchange Online"
        Connect-ExchangeOnline
    }

    ##### Start Main Selection Form #####
    #Set Properties Of MainForm
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.width = 375
    $MainForm.height = 425
    $MainForm.MaximizeBox = $false
    $MainForm.MinimizeBox = $false
    $MainForm.StartPosition = 'CenterScreen'
    $MainForm.Text = "Please select options"
    $MainForm.Font = New-Object System.Drawing.Font("Times New Roman",12)
    $MainForm.MaximumSize = New-Object System.Drawing.Size(375,425)
    $MainForm.MinimumSize = New-Object System.Drawing.Size(375,425)

    #Create Shared Mailbox Checkbox 
    $ConvertCheckBox = new-object System.Windows.Forms.checkbox
    $ConvertCheckBox.Location = new-object System.Drawing.Size(25,25)
    $ConvertCheckBox.Size = new-object System.Drawing.Size(250,50)
    $ConvertCheckBox.Text = "Convert to Shared Mailbox?"
    $ConvertCheckBox.Checked = $true
    $MainForm.Controls.Add($ConvertCheckBox)  

    #Create License Checkbox 
    $LicenseCheckBox = new-object System.Windows.Forms.checkbox
    $LicenseCheckBox.Location = new-object System.Drawing.Size(25,75)
    $LicenseCheckBox.Size = new-object System.Drawing.Size(250,50)
    $LicenseCheckBox.Text = "Remove All Licenses?"
    $LicenseCheckBox.Checked = $true
    $MainForm.Controls.Add($LicenseCheckBox) 

    #Create Share The Mailbox Checkbox 
    $GrantMailboxCheckBox = new-object System.Windows.Forms.checkbox
    $GrantMailboxCheckBox.Location = new-object System.Drawing.Size(25,125)
    $GrantMailboxCheckBox.Size = new-object System.Drawing.Size(250,50)
    $GrantMailboxCheckBox.Text = "Share the Mailbox?"
    $GrantMailboxCheckBox.Checked = $true
    $MainForm.Controls.Add($GrantMailboxCheckBox) 
    $GrantMailboxCheckBox.Add_CheckStateChanged({ 
        $OneDriveSame.Enabled = $GrantMailboxCheckBox.Checked 
        $OneDriveNo.Checked = $true
    }) 

    # Create Group For Radio Buttons
    $OneDriveGroupBox = New-Object System.Windows.Forms.GroupBox
    $OneDriveGroupBox.Location = '25,175'
    $OneDriveGroupBox.size = '300,150'
    $OneDriveGroupBox.text = "Share OneDrive?"

    # Create The Radio Buttons
    $OneDriveNo = New-Object System.Windows.Forms.RadioButton
    $OneDriveNo.Location = '10,25'
    $OneDriveNo.size = '350,20'
    $OneDriveNo.Checked = $false
    $OneDriveNo.Text = "No."
     
    $OneDriveSame = New-Object System.Windows.Forms.RadioButton
    $OneDriveSame.Location = '10,75'
    $OneDriveSame.size = '350,20'
    $OneDriveSame.Checked = $true
    $OneDriveSame.Text = "To Same User As Shared Mailbox"
     
    $OneDriveDiff = New-Object System.Windows.Forms.RadioButton
    $OneDriveDiff.Location = '10,125'
    $OneDriveDiff.size = '350,20'
    $OneDriveDiff.Checked = $false
    $OneDriveDiff.Text = "To Different User As Shared Mailbox"
    $OneDriveGroupBox.Controls.AddRange(@($OneDriveNo,$OneDriveSame,$OneDriveDiff))
    $MainForm.Controls.Add($OneDriveGroupBox)

    #Add An OK Button
    Clear-Variable OKButton -ea SilentlyContinue
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.AutoSize = $true
    $OKButton.Location = new-object System.Drawing.Size(137,350)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$MainForm.Close();$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK})
    $OKButton.Enabled = $true
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::None
    $MainForm.Controls.Add($OKButton)

    #Activate The MainForm
    $MainForm.Add_Shown({$MainForm.Activate()})
    [void] $MainForm.ShowDialog()
    ##### End Main Selection Form #####
    
    #If OK Was Clicked On Main Form, Else Exit Script If Selection Box Closed
    if ($OKButton.DialogResult -eq 'OK') {
        #Pull All Azure AD Users and Store Ib Hash Table Instead Of Calling Get-AzureADUser Multiple Times
        $allUsers = @{}    
        foreach ($user in Get-AzureADUser -All $true){
            $allUsers[$user.UserPrincipalName] = $user
        }

        #Request Username(s) To Be Terminated From Script Runner (Hold Ctrl To Select Multiples)
        $usernames = $allUsers.Values | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -Passthru -Title "Please select the user(s) to be terminated" | Select-Object -ExpandProperty UserPrincipalName
        #Kill Script If Ok Button Not Clicked
        if ($null -eq $usernames) {
            Throw
        }
        ##### Start User(s) Loop #####
        foreach ($username in $usernames) {
            $UserInfo = $allusers[$username]
            #Request User(s) To Share Mailbox With When Grant Access Is Selected
                if ($GrantMailboxCheckBox.Checked -eq $true) {
                $sharedMailboxUser = $allUsers.Values | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the $username Shared Mailbox with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
                #Kill Script If Ok Button Not Clicked
                if ($null -eq $sharedMailboxUser) {
                    Throw
                }
            }
            
            #Block Sign In Of User/Force Sign Out Within 60 Minutes
            Write-Verbose -Message "Blocking Sign In to force log out on all sessions within 60 minutes"
            Set-AzureADUser -ObjectID $UserInfo.ObjectId -AccountEnabled $false
            Write-Verbose -Message "Sign in Blocked"

            #Remove All Group Memberships
            Write-Verbose -Message "Removing all group memberships, skipping Dynamic groups as they cannot be removed this way"
            $memberships = $SharedMailboxUserInfo.ObjectId | Get-AzureADUserMembership | Where-Object {$_.ObjectType -ne "Role"}  | ForEach-Object {Get-AzureADGroup -ObjectId $_.ObjectId | Select-Object DisplayName,ObjectId}
            foreach ($membership in $memberships) { Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $UserInfo.ObjectId }
            Write-Verbose -Message "All non-dynamic groups removed, please check your Downloads folder for the file, it will also open automatically at end of user termination"

            #Convert To Shared Mailbox And Hide From GAL When Convert Is Selected, Must Be Done Before Removing Licenses
            if ($ConvertCheckBox.Checked -eq $true) {
                Write-Verbose -Message "Converting $username to Shared Mailbox and Hiding from GAL"
                Set-Mailbox $username -Type Shared -HiddenFromAddressListsEnabled $true
                Write-Verbose -Message "Mailbox for $username converted to Shared, address hidden from GAL"
            }

            #Grant Access To Shared Mailbox When Grant CheckBox Is Selected
            $SharedMailboxUserInfo = $allusers[$sharedMailboxUser]
            if ($GrantMailboxCheckBox.Checked -eq $true) {
                Write-Verbose -Message "Granting access to the $username Shared Mailbox to $sharedMailboxUser"
                Add-MailboxPermission -Identity $username -User $SharedMailboxUser -AccessRights FullAccess -InheritanceType All
                Add-RecipientPermission -Identity $username -Trustee $SharedMailboxUser -AccessRights SendAs -Confirm:$False
                Write-Verbose -Message "Access granted to the $username Shared Mailbox to $sharedMailboxUser"
            }

            #Remove All Licenses When Remove Licenses Is Selected
            if ($LicenseCheckBox.Checked -eq $true) {
                Write-Verbose -Message "Removing all licenses"
                $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                if($UserInfo.assignedlicenses){
                $licenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $licenses}
                Write-Verbose -Message "Licenses have all been removed"
            }
            
            ##### Start OneDrive Block #####
            #Test And Connect To Sharepoint Online If Needed
            if ($OneDriveNo.Checked -ne $true) {
                $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
                $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
                try {
                    Write-Verbose -Message "Testing connection to SharePoint Online"
                    Get-SPOSite -ErrorAction Stop | Out-Null
                    Write-Verbose -Message "Already connected to SharePoint Online"
                }
                catch {
                    Write-Verbose -Message "Connecting to SharePoint Online"
                    Connect-SPOService -Url $AdminSiteURL
                } 
            }
            
            #Share OneDrive With Same User as Shared Mailbox
            if ($OneDriveSame.Checked -eq $true) {
                #Pull OneDriveSiteURL Dynamically And Grant Access
                $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

                #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
                Write-Verbose -Message "Adding $SharedMailboxUser to OneDrive folder for access to files"
                Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedMailboxUser -IsSiteCollectionAdmin $True
                Write-Verbose "OneDrive Data Shared with $SharedMailboxUser successfully, link to copy and give to Manager is $OneDriveSiteURL"
            }
            #Share OneDrive With Different User(s) than Shared Mailbox
            elseif ($OneDriveDiff.Checked -eq $true) {
                $SharedOneDriveUser = $allusers.Values | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the Mailbox and OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
                
                #Pull Object ID Needed For User Receiving Access To OneDrive And OneDriveSiteURL Dynamically
                $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

                #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
                Write-Verbose -Message "Adding $SharedOneDriveUser to OneDrive folder for access to files"
                Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedOneDriveUser -IsSiteCollectionAdmin $True
                Write-Verbose "OneDrive Data Shared with $SharedOneDriveUser successfully, link to copy and give to Manager is $OneDriveSiteURL"
            }
            ##### End OneDrive Block #####
            
            #Export Groups Removed and OneDrive URL to CSV
            [pscustomobject]@{
                GroupsRemoved    = $memberships.DisplayName -join ';'
                OneDriveSiteURL = $OneDriveSiteURL
            } | Export-Csv -Path c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv -NoTypeInformation

            #Open Created CSV File At End Of Loop For Ease Of Copying OneDrive URL To Give
            Start-Process c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv
        }
        ##### End User(s) Loop #####    
    }
    #Kill Script If "OK" Is Not Clicked On Main Form
    else { Throw }
#Create Quit Prompt and Close While Loop
$quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to terminate another user?" , "User Termination(s) Complete" , 4)
#Need to clear all variables before starting loop over - I think
}
##### End Main Loop #####