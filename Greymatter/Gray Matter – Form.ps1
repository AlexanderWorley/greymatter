function Start-Logging(){
                param(
                                $log
                )
}
Function Main{
                #File Locations for Modules, root location.
                $Global:rootLocation= "C:\Program Files\Grey Matter"
                #Erase logs older than 30 days and starts the new log after that
                $Script:limit= (Get-Date).AddDays(-30)
                #Get-ChildItem -Path "$Global:rootLocation\Logs" -Recurse -Force | Where-Object { !$_.PSIsContainer-and $_.CreationTime-lt$limit } | Remove-Item -Force
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Grey Matter Core: `"Executing Transcript`""
                #Start-Transcript -Path "$Global:rootLocation\Logs\Log_$(Get-Date -UFormat"%Y-%m-%d_%H-%M").log"
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Grey Matter Core: `"Importing Modules`""
                Import-Module -name "$Global:rootlocation\Grey Matter - Modules\Gray Matter_EmployeeData Table Source.psm1"
                Import-Module -name "$Global:rootlocation\Grey Matter - Modules\Gray Matter_StoreData Table Source.psm1"
                Import-Module -name "$Global:rootlocation\Grey Matter - Modules\PingDevices.psm1"
                Import-Module -name "$Global:rootlocation\Grey Matter - Modules\printQueue.psm1"
                #Pulls the config file
                [xml]$Global:storeConfig= Get-Content "$Global:rootlocation\config.xml"
                #Gets the HR file for the teller ID
                #Import-HRFiles
                #This needs ran ASAP to allow the other functions to work. Pulls the DHCP List then executes the form.
                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                Get-DHCPList-single
                }
                Elseif($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Multi'){
                                Get-DHCPList-Multi
                }
                #imports AD module and assemblies.
                Import-Module ActiveDirectory
               
                Open-GUI
}

Function Open-GUI{
                Add-Type -AssemblyNameMicrosoft.VisualBasic
                Add-Type -AssemblyNameSystem.Windows.Forms
                [System.Windows.Forms.Application]::EnableVisualStyles()
                #Launches the main form               
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Grey Matter Core: `"Initializing Form`""
                $Image = [system.drawing.image]::FromFile("$Global:rootlocation\Images\bg.jpg")
                $Script:Server= 'corp.checksmart.com'
                $Form = New-Object system.Windows.Forms.Form
                $Form.ClientSize= '818,520' 
                $Form.text= "CCFI Grey Matter(Beta v2.0)"
                $Form.backgroundImage= $Image
                $form.Icon= "$Global:rootLocation\Images\BlankIcon.ico"
                $Form.TopMost= $false
                $Form.Add_FormClosing({
                                if($Global:pingScreen-ne $null){
                                                $Global:pingScreen.Close()
                                                $Global:pingScreen.Dispose()
                                }
                                if($Global:PrintQueue-ne $null){
                                                If($Global:PrintQueueHash.ThreadID-ne $null){
                                                                $Kill = Get-Runspace-ID $Global:PrintQueueHash.ThreadID
                                                                $Kill.Close()
                                                }
                                                $Global:PrintQueue.Close()
                                                $Global:PrintQueue.Dispose()
                                }
                })
                #region MenuBar
                #Creating the menu strip and the varuiousdrop down options and their respective sub options
                $greyMenuMS= New-Object System.Windows.Forms.MenuStrip
                $greyFileDropdownMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyRefreshConfigFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyRefreshDHCPFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyEemailerFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyExitFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyHelpDropdownMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $greyHelpDocHelpSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                #Adding the drop down options to the menu strip
                $greyMenuMS.Items.AddRange(@($greyFileDropdownMI,$greyHelpDropdownMI)) | Out-Null
                $greyMenuMS.Location= '0,0'
                $greyMenuMS.Size= '475,24'
                $greyMenuMS.TabIndex= '0'
                $Form.Controls.Add($greyMenuMS)
                #Adding the sub options under the file drop down
$greyFileDropdownMI.DropDownItems.AddRange(@($greyRefreshConfigFileSubMI,$greyRefreshDHCPFileSubMI,$greyEemailerFileSubMI,$greyExitFileSubMI)) | Out-Null
                $greyFileDropdownMI.Size= '35,20'
                $greyFileDropdownMI.Name= 'fileDropdownMI'
                $greyFileDropdownMI.Text= "&File"
                $greyRefreshConfigFileSubMI.Size= '182,20'
                $greyRefreshConfigFileSubMI.Text= "Refresh &Config"
                $greyRefreshConfigFileSubMI.Add_Click({
                                $Global:storeConfig= Get-Content "$Global:rootlocation\config.xml"
                                [System.Windows.Forms.MessageBox]::Show("Config file was refreshed.","ConfigRefresh","Ok","Information")
                })
                $greyRefreshDHCPFileSubMI.Size= '182,20'
                $greyRefreshDHCPFileSubMI.Text= "Refresh &DHCP"
                $greyEemailerFileSubMI.Size= '100,20'
                $greyEemailerFileSubMI.Text= 'E&mailer'
                $greyRefreshDHCPFileSubMI.Add_Click({
                                $Form.Controls.Add($Refresh_label)
                                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                                Get-DHCPList-single
                                }
                                Elseif($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Multi'){
                                                Get-DHCPList-Multi
                                }
                                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                                While($Global:data.IsCompleted-ne $true){}
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Closing and Disposing DHCP Job`""
                                                $Global:command.EndInvoke($Global:data)
                                                $Global:command.Dispose()
                                }
                                Start-Logging $syncHash.Bug
                                [System.Windows.Forms.MessageBox]::Show("DHCP master list(s) refresh complete.","DHCPRefresh","Ok","Information")
                                $Form.Controls.Remove($Refresh_label)
                })
                $greyExitFileSubMI.Size= '182,20'
                $greyExitFileSubMI.Text= "&Exit"
                $greyExitFileSubMI.Add_Click({
                                $Form.Close()
                })
                #Adding the sub option under the help drop down
                $greyHelpDropdownMI.DropDownItems.Add($greyHelpDocHelpSubMI) | Out-Null
                $greyHelpDropdownMI.Size= '51,20'
                $greyHelpDropdownMI.Text= "&Help"
                $greyHelpDocHelpSubMI.Size= '250,20'
                $greyHelpDocHelpSubMI.Text= "&User Manual"
                $greyHelpDocHelpSubMI.Add_Click({
                                Start-Process -FilePath"$Global:rootlocation\Documents\Grey Matter_User-Guide.pdf"
                })
                #endregion MenuBar
                #Only to be added if File>Refresh DHCP was selected
                $Refresh_label= New-Object system.Windows.Forms.Label
                $Refresh_label.text= "Refreshing DHCP List(s)...."
                $Refresh_label.AutoSize= $true
                $Refresh_label.width= 25
                $Refresh_label.height= 10
                $Refresh_label.backcolor= [System.Drawing.Color]::FromName("Transparent")
                $Refresh_label.location= New-Object System.Drawing.Point(1,23)
                $Refresh_label.forecolor= "white"
                $Refresh_label.font= New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]::Bold)
                #Gridview for Employee Lookup (Right most Grid)            
                $users_Lookup_Grid= New-Object system.Windows.Forms.DataGridView
                $users_Lookup_Grid.AllowUserToAddRows=$false
                $users_Lookup_Grid.width= 264
                $users_Lookup_Grid.height= 408
                $users_Lookup_Grid.ReadOnly= $True
                $users_Lookup_Grid.ColumnHeadersVisible= $true
                $users_Lookup_Grid.AutoGenerateColumns= $true;
                $users_Lookup_Grid.AllowUserToAddRows=$false
                foreach ($row in $users_Lookup_GridData){
                                $users_Lookup_Grid.Rows.Add($row)
                }
                $users_Lookup_Grid.location= New-Object System.Drawing.Point(555,120)
 
 
                #Filter drop box for employee Name/EmpID
                $Employee_Selection_Filter= New-Object system.Windows.Forms.ComboBox
                $Employee_Selection_Filter.text= "Employee ID"
                #$Employee_Selection_Filter.Dropdownstyle = 'DropDownList'
                $Employee_Selection_Filter.width= 100
                $Employee_Selection_Filter.height= 20 
                $Employee_Selection_Filter.location= New-Object System.Drawing.Point(450,45)
                $Employee_Selection_Filter.Font= 'Calibri,10'
                #Employee Filter Array - Add to the array then call the string in the if-Elseif statement listed in the Employee Lookup Module
                @('Employee ID','Name') | ForEach-Object {[void] $Employee_Selection_Filter.Items.Add($_)}
                $Employee_Selection_Filter.SelectedIndex= 0
 
 
                #Employee Lookup Label in Black.
                $Employee_Lookup_label= New-Object system.Windows.Forms.Label
                $Employee_Lookup_label.text= "Employee LookUp"
                $Employee_Lookup_label.AutoSize= $true
                $Employee_Lookup_label.width= 25
                $Employee_Lookup_label.height= 10
                $Employee_Lookup_label.backcolor= [System.Drawing.Color]::FromName("Transparent")
                $Employee_Lookup_label.location= New-Object System.Drawing.Point(555,25)
                $Employee_Lookup_label.forecolor= "Black"
                $Employee_Lookup_label.font= New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]::Bold)
 
                #Employee Look up text box and enter key logic to pass through to the module referedabove
                $Lookup = {
                                if ($_.KeyCode-eq "Enter") {
                                                #If statement: pulls the array values for the filter
                                                $_.SuppressKeyPress= $true
                                                <#EmployeeID#>
                                                if($Employee_Selection_Filter.SelectedIndex-eq 0){
                                                                $hold = updateData-Employee_Lookup_Value$Employee_Lookup_TextBox.Text
                                                                $users_Lookup_Grid.datasource= $hold.Tables[0]
                                                                $users_Lookup_Grid.Update()
                                                               
                                                <#FullName#>
                                                }
                                                Elseif($Employee_Selection_Filter.SelectedIndex-eq 1){
                                                                <#Normal Pass through with cn#>
                                                                #counts the number of values
                                                                $Search = "*"+$Script:Employee_Lookup_TextBox.text+"*"
                                                                $Employee_Selection_Filter.SelectedIndex= 0
                                                                $Script_Duplicate= get-ADUser-server corp.checksmart.com -filter {(givenName-like $Search) -or (sn-like $Search) -or (cn-like $Search)} -properties cn, EmployeeID, Title, Department, Manager
                                                                $Script_Duplicate_Measure= $Script_Duplicate| measure
                                                 
                                                                Start-Logging $Script_Duplicate_Measure
                                                                if($Script_Duplicate_Measure.count-gt1){
                                                                                Add-Type -AssemblyNameSystem.Windows.Forms
                                                                                [System.Windows.Forms.Application]::EnableVisualStyles()
 
                                                                                $Duplicate_Form= New-Object system.Windows.Forms.Form
                                                                                $Duplicate_Form.ClientSize= '634,320'
                                                                                $Duplicate_Form.text= "Form"
                                                                                $Duplicate_Form.TopMost= $false
                                                                                $Duplicate_Form.backgroundImage= $Image
 
 
                                                                                $Duplicate_Employee_Value= New-Object system.Windows.Forms.DataGridView
                                                                                $Duplicate_Employee_Value.width  = 551
                                                                                $Duplicate_Employee_Value.height  = 282
                                                                                $Duplicate_Employee_Value.ReadOnly= $true
                                                                                $Duplicate_Employee_Value.MultiSelect= $false
                                                                                $Duplicate_Employee_Value.AllowUserToAddRows=$false
                                                                                $Duplicate_Employee_Value.location  = New-Object System.Drawing.Point(2,37)
 
                                                                                $Duplicate_Select_Button= New-Object system.Windows.Forms.Button
                                                                                $Duplicate_Select_Button.text= "Select"
                                                                                $Duplicate_Select_Button.width= 75
                                                                                $Duplicate_Select_Button.height= 20
                                                                                $Duplicate_Select_Button.location= New-Object System.Drawing.Point(477,15)
                                                                                $Duplicate_Select_Button.Font= 'Calibri,10'
                                                                                $Duplicate_Select_Button.Add_Click({
                                                                                                $Script:Employee_Lookup_TextBox.Text= $Duplicate_Employee_Value.CurrentRow.Cells[1].Value
                                                                                                                #Checks if the account is enabled, assigns a global variable
                                                                                                #Pulls the full name of the assoicate, assigns as a global variable. Used for the message boxes that convert empIDto CN. 1
                                                                                                $global:Employee_Lookup_TextBox_Name= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertycn
                                                                                                if($Employee_Lookup_TextBox_Name_EnabledCheck-eq $False){
                                                                                                                #Disabled account message.Pullsthe account information anyways. Will no password reset or account unlock.
                                                                                                                [System.Windows.Forms.MessageBox]::Show($Employee_Lookup_TextBox_Name + " Is currently Disabled. Unable to make changes to the account.","AccountLookup Reset","Ok")
                                                                                                                $hold = updateData-Employee_Lookup_Value  $Script:Employee_Lookup_Textbox.text
                                                                                                                $users_Lookup_Grid.datasource= $hold.Tables[0]
                                                                                                                $users_Lookup_Grid.Update()
                                                                                                }Else{
                                                                                                                #Updates the gridviewwhen you hit enter.
                                                                                                                $hold = updateData-Employee_Lookup_Value  $Script:Employee_Lookup_Textbox.text
                                                                                                                $users_Lookup_Grid.datasource= $hold.Tables[0]
                                                                                                                $users_Lookup_Grid.Update()
                                                                                                }
                                                                                                $Duplicate_Form.Close()
                                                                                })
                                                                                                #Label above the Duplicate users gridview.
                                                                                $Store_Duplicate_Label= New-Object system.Windows.Forms.Label
                                                                                $Store_Duplicate_Label.text= "Select the correct user"
                                                                                $Store_Duplicate_Label.AutoSize  = $true
                                                                                $Store_Duplicate_Label.width= 25
                                                                                $Store_Duplicate_Label.height= 10
                                                                                $Store_Duplicate_Label.location= New-Object System.Drawing.Point(6,18)
                                                                                $Store_Duplicate_Label.forecolor= "White"               
                                                                                $Store_Duplicate_Label.backcolor= [System.Drawing.Color]::FromName("Transparent")
                                                                                $Store_Duplicate_Label.Font= $Store_Duplicate_Label.Font= 'Microsoft Sans Serif,10,style=Bold'
                                                                $Duplicate_Form.controls.AddRange(@($Duplicate_Employee_Value,$Duplicate_Select_Button,$Store_Duplicate_Label))
                                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateData: `"Exporting Data - $Script:Employee_Lookup_TextBox`""
                                                                                  
                                                                                # Duplicate GridviewMain Columns
                                                                                $Duplicate_GridDataTable= New-Object System.Data.DataTable
                                                                                #Column A
                                                                                $Duplicate_GridDataTable.Columns.Add('Name', [string]) | Out-Null
                                                                                #Column B
                                                                                $Duplicate_GridDataTable.Columns.Add('EmployeeID', [string]) | Out-Null
                                                                                #Column C
                                                                                $Duplicate_GridDataTable.Columns.Add('Title', [string]) | Out-Null
                                                                                #Column D
                                                                                $Duplicate_GridDataTable.Columns.Add('Department', [string]) | Out-Null
                                                                                #Column E
                                                                                $Duplicate_GridDataTable.Columns.Add('Manager', [string]) | Out-Null
 
                                                                                #Passes the table as a dataset. Returns the data set. This is what processes the datagridview.
                                                                                $Duplicate_ds= New-Object System.Data.DataSet
                                                                                $Duplicate_ds.Tables.Add($Duplicate_GridDataTable)
                                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateData: `"Stop - $Script:Employee_Lookup_TextBox`""
                                                               
                                                                                ForEach($possibleEmployeein $Script_Duplicate){
                                                                                                #Passes the column sets as horizontal. Passes value for Multiple Users using the For loop
                                                                                                $Name = $possibleEmployee| Select -ExpandPropertycn
                                                                                                $EmployeeID= $possibleEmployee| Select -ExpandPropertyEmployeeID
                                                                                                $Title = $possibleEmployee| Select -ExpandPropertyTitle
                                                                                                $Department=  $possibleEmployee| Select -ExpandPropertyDepartment
                                                                                                $Manager= $possibleEmployee| Select @{N='Manager';E={(Get-ADUser $_.Manager).Name}}
                                                                                                $Manager = $Manager.Manager
                                                                                                $Duplicate_ds.Tables[0].Rows.Add($Name, $EmployeeID, $Title, $Department, $Manager) | Out-Null
                                                                                }
                                                               
                                                                                $Duplicate_Employee_Value.DataSource= $Duplicate_ds.Tables[0]
 
                                                                                $Duplicate_Form.ShowDialog()
                                                                }
                                                                elseif($Script_Duplicate_Measure.count-eq 1){
                                                                                $Script:Employee_Lookup_TextBox.Text= $Script_Duplicate.EmployeeID
                                                                                #Checks if the account is enabled, assigns a global variable                                                           
                                                                                #Pulls the full name of the assoicate, assigns as a global variable. Used for the message boxes that convert empIDto CN. 1
                                                                                $global:Employee_Lookup_TextBox_Name= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Script:Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertycn
                                                                                if($Employee_Lookup_TextBox_Name_EnabledCheck-eq $False){
                                                                                                $hold = updateData-Employee_Lookup_Value  $Script:Employee_Lookup_Textbox.text
                                                                                                $users_Lookup_Grid.datasource= $hold.Tables[0]
                                                                                                $users_Lookup_Grid.Update()
                                                                                }Else{
                                                                                                #Updates the gridviewwhen you hit enter.
                                                                                                $hold = updateData-Employee_Lookup_Value  $Employee_Lookup_Textbox.text
                                                                                                $users_Lookup_Grid.datasource= $hold.Tables[0]
                                                                                                $users_Lookup_Grid.Update()
                                                                                }
                                                                }
                                                                else{
                                                                                [System.Windows.Forms.MessageBox]::Show("Invalid Employee, please retry your action.", 'Employee Look Up','Ok',"Error")
                                                                }
                                                }
                                }
                }
                $Script:Employee_Lookup_TextBox= New-Object system.Windows.Forms.TextBox
                $Script:Employee_Lookup_TextBox.multiline= $false
                $Script:Employee_Lookup_TextBox.width= 146
                $Script:Employee_Lookup_TextBox.height= 20
                $Script:Employee_Lookup_TextBox.location= New-Object System.Drawing.Point(554,45)
                $Script:Employee_Lookup_TextBox.Font= 'Microsoft Sans Serif,10'
                $Script:Employee_Lookup_TextBox.Add_KeyDown($Lookup)
 
                #Script that displays the Account Unlock button, and run checks for if the account is disabled.
                $AD_Account_Unlock= New-Object system.Windows.Forms.Button
                $AD_Account_Unlock.text= "Unlock"
                $AD_Account_Unlock.width= 100
                $AD_Account_Unlock.height= 26
                $AD_Account_Unlock.location= New-Object System.Drawing.Point(715,60)
                $AD_Account_Unlock.Font= 'Calibri,10'
                $AD_Account_Unlock.Add_Click({      
                $Global:GetSAM_FullName= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertycn
                $global:Employee_Lookup_TextBox_Name_EnabledCheck=  get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertyenabled       
                                #Checks to see if the account is enabled or disabled.Ifdisabled message box stating so will popupand nothing else will activate.
                                if($global:Employee_Lookup_TextBox_Name_EnabledCheck-eq $False){
                                                [System.Windows.Forms.MessageBox]::Show($Global:GetSAM_FullName+ " Is currently Disabled, Unable to unlock the account!","AccountUnlock","Ok","Error")
                                }Else{
                                                #Unlocks the account if Yes is selected. If no is selected the statement else statement will launch.
                                                if ([System.Windows.Forms.MessageBox]::Show("Are you sure you wish to unlock "+ $Global:GetSAM_FullName,"Account Unlock","YesNo") -eq [System.Windows.Forms.DialogResult]::Yes){
                                                                Start-Logging "Unlocked"           
                                                                $GetSAM= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertysAMAccountName
                                                                Unlock-ADAccount-Identity $GetSAM
                                                                [System.Windows.Forms.MessageBox]::Show($Global:GetSAM_FullName+' has been unlocked', 'Account Unlock','Ok')  
                                                }else{
                                                                Start-Logging "locked"           
                                                                [System.Windows.Forms.MessageBox]::Show($Global:GetSAM_FullName  +' has been skipped', 'Account Unlock','Ok')
                                                }
                                }
                })
                #Password Reset button and logic
                $AD_Account_Password_Reset= New-Object system.Windows.Forms.Button
                $AD_Account_Password_Reset.text= "Password Reset"
                $AD_Account_Password_Reset.width= 100
                $AD_Account_Password_Reset.height= 26
                $AD_Account_Password_Reset.location= New-Object System.Drawing.Point(715,90)
                $AD_Account_Password_Reset.Font= 'Calibri,10'
                #Password Reset Logic
                $AD_Account_Password_Reset.Add_Click({
                                $Employee_Lookup_TextBox_Name= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertycn
                                #checks to see if the account is disabled. No action will occur other than the messagebox.
                                $Employee_Lookup_TextBox_Name_EnabledCheck=  get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertyenabled       
 
                                if($Employee_Lookup_TextBox_Name_EnabledCheck-eq $False){
                                                [System.Windows.Forms.MessageBox]::Show($Employee_Lookup_TextBox_Name + " Is currently Disabled. Unable to reset the password.","AccountPassword Reset","Ok","Error")
 
                                }Else{
 
                                                #Message box requesting confirmation for whom you are resetting the account for. If Yes, it will ask for the new password and confirm in plain text. If no it will state "User has been skipped".
                                               
                                                if ([System.Windows.Forms.MessageBox]::Show("Are you sure you wish to Reset the password for: "+ $Employee_Lookup_TextBox_Name,"AccountPassword Reset","YesNo") -eq [System.Windows.Forms.DialogResult]::Yes){
                                                  $newPass= [Microsoft.VisualBasic.Interaction]::InputBox('Enter Password ', 'Password Reset')
                                                  #Tests to see if the string for the new password is Null or Empty. If True it will throw an error and cancel the operation.
                                                                if ([string]::IsNullOrEmpty($newPass)){               
                                                                                [System.Windows.Forms.MessageBox]::Show("Unable to reset the password for:"+ $Employee_Lookup_TextBox_Name+ " , New password field is empty!","AccountPassword Reset","Ok","Error")
                                                                                #continues the operation, asking to force the password to be reset upon login, and confirmed the new password.
                                                                }else{
                                                               
                                                                $GetSAM= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertysAMAccountName
                                                                Start-Logging "Unlocked $GetSAM"
                                                                Set-ADAccountPassword-Identity $GetSAM-Reset -NewPassword(ConvertTo-SecureString-AsPlainText"$newPass" -Force)
                                                                if([System.Windows.Forms.MessageBox]::Show("Would you like to force a password reset when the user logs in?", 'Account Password Reset','YesNo') -eq [System.Windows.Forms.DialogResult]::Yes){
                                                                Set-aduser$GetSAM-changepasswordatlogon$true
                                                                [System.Windows.Forms.MessageBox]::Show("Password has been reset to " + $newPass, 'Account Password Reset','Ok')
                                                                }else{
                                                                [System.Windows.Forms.MessageBox]::Show("They will not be prompted to change their password. Password has been changed to: " + $newPass, 'Account Password Reset','Ok')
                                                                } 
                                                  }          
                                                }else{
                                                                Start-Logging "Skipped $GetSAM"
                                                                #Exit for the if function that returns you to the main form.
                                                                $GetSAM_No= get-aduser-server $Script:Server-Filter{EmployeeID-eq $Employee_Lookup_TextBox.Text} -Properties * | Select -ExpandPropertycn
                                                                [System.Windows.Forms.MessageBox]::Show($GetSAM_no  +' has been skipped', 'Account Password Reset','Ok')
                                                }
                                }
                })
                #Manual Remote Control Label
                $Manual_RemoteControl_Label= New-Object system.Windows.Forms.Label
                $Manual_RemoteControl_Label.text= "Manual Remote Control"
                $Manual_RemoteControl_Label.AutoSize= $true
                $Manual_RemoteControl_Label.width= 25
                $Manual_RemoteControl_Label.height= 10
                $Manual_RemoteControl_Label.backcolor= [System.Drawing.Color]::FromName("Transparent")
                $Manual_RemoteControl_Label.location= New-Object System.Drawing.Point(554,69)
                $Manual_RemoteControl_Label.forecolor= "Black"
                $Manual_RemoteControl_Label.font= New-Object System.Drawing.Font('Calibri',10,[System.Drawing.FontStyle]::Bold)
 
                #manual Remote Control button and execution. Mainly for remoting into computers that are not listed in the store machine view.
                $Manual_RemoteControl= New-Object system.Windows.Forms.TextBox
                $Manual_RemoteControl.multiline= $false
                $Manual_RemoteControl.width= 140
                $Manual_RemoteControl.height= 10
                $Manual_RemoteControl.location= New-Object System.Drawing.Point(555,90)
                $Manual_RemoteControl.Font= 'Microsoft Sans Serif,10'
                $Manual_RemoteControl.Add_KeyDown({
                                if ($_.KeyCode-eq "Enter") {
                                                $_.SuppressKeyPress= $true
                                                #Setup a trap for none number and too large or too small       
                                                                  $CorpComputer= $Manual_RemoteControl.text
                                                                  Start-Process -FilePath"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe" -ArgumentList$CorpComputer
                                }
                })
 
                #Store look up text box.
                $Global:Store_Lookup_textBox= New-Object system.Windows.Forms.TextBox
                $Global:Store_Lookup_textBox.multiline= $false
                $Global:Store_Lookup_textBox.width= 150
                $Global:Store_Lookup_textBox.height= 20
                $Global:Store_Lookup_textBox.location= New-Object System.Drawing.Point(2,87)
                $Global:Store_Lookup_textBox.Font= 'Microsoft Sans Serif,10'
                #Passes the store number value requires store number count to equal 3 or 4.
                $Global:Store_Lookup_textBox.Add_KeyDown({
                                if ($_.KeyCode-eq "Enter") {
                                                $_.SuppressKeyPress= $true
                                                #Setup a trap for none number and too large or too small
                                If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Global:Store_Lookup_textBox.Text)']") -eq $null){
                                                                [System.Windows.Forms.MessageBox]::Show("$($Global:Store_Lookup_textBox.Text) does not meet the store number standard. `nPleasecheck the text box and try again.", "Invalid Store Number Entry","OK","Error")
                                                }
                                                else{
                                                                $hold = updateStoreView-Store_Lookup_Value$Global:Store_Lookup_textBox.text
                                                                $hold = $hold[$hold.length-1]
                                                                $Store_Users_GridView.datasource= $hold.Tables[0]
                                                                $Store_Users_GridView.Update()
                                                                $scope = Get-StoreScope-StoreNumber$Global:Store_Lookup_textBox.text
                                                                $temp = Search-StoreComputers-StoreNumber$Global:Store_Lookup_textBox.text-ScopeID$scope
                                                                $Store_Mach_Grid.datasource= $temp.Tables[0]
                                                                $Store_Mach_Grid.Update()
                                                }
                                }
                })
                #Store label in White
                $Store_Lookup_Label= New-Object system.Windows.Forms.Label
                $Store_Lookup_Label.text= "Store Number"
                $Store_Lookup_Label.AutoSize= $true
                $Store_Lookup_Label.width= 15
                $Store_Lookup_Label.height= 10
                $Store_Lookup_Label.location= New-Object System.Drawing.Point(1,68)
                $Store_Lookup_Label.backcolor= [System.Drawing.Color]::FromName("Transparent")
                $Store_Lookup_Label.forecolor= "#f9fcfc"
                $Store_Lookup_Label.Font= New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]::Bold)
 
                #Remote control button with logic to pass through the arugmentof the machine name/IP Address.
                $Remote_Control_Button= New-Object system.Windows.Forms.Button
                $Remote_Control_Button.text= "Remote Control"
                $Remote_Control_Button.width= 102
                $Remote_Control_Button.height= 30
                $Remote_Control_Button.location= New-Object System.Drawing.Point(160,85)
                $Remote_Control_Button.Font= 'Calibri,10'
                $Remote_Control_Button.Add_Click({
                                $Computer = $Store_Mach_Grid.SelectedCells.Value
                                Start-Process -FilePath"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe" -ArgumentList$Computer
                })
                $Print_Queue_button= New-Object system.Windows.Forms.Button
                $Print_Queue_button.text= "Print Queue"
                $Print_Queue_button.width= 102
                $Print_Queue_button.height= 30
                $Print_Queue_button.location= New-Object System.Drawing.Point(358,85)
                $Print_Queue_button.Font= 'Calibri,10'
                $Print_Queue_button.Add_Click({
                                Open-PrintQueueGUI
                })
                #pinger application launcher button.
                $Ping_Button= New-Object system.Windows.Forms.Button
                $Ping_Button.text= "Pinger"
                $Ping_Button.width= 90
                $Ping_Button.height= 30
                $Ping_Button.location= New-Object System.Drawing.Point(265,85)
                $Ping_Button.Font= 'Calibri,10'
                $Ping_Button.Add_Click({
                                Open-PingGUI-Background $Image
                })
 
                #Store users grid view (Top most gridview)
                $Store_Users_GridView= New-Object system.Windows.Forms.DataGridView
                $Store_Users_GridView.width= 543
                $Store_Users_GridView.height= 190
                $Store_Users_GridView.ReadOnly= $True
                $Store_Users_GridView.ColumnHeadersVisible= $true
                $Store_Users_GridView.AutoGenerateColumns= $true;
                $Store_Users_GridView.AllowUserToAddRows=$false
                foreach ($row in $store_Lookup_GridData){
                                $Store_Users_GridView.Rows.Add($row)
                }
                $Store_Users_GridView.location= New-Object System.Drawing.Point(-1,120)
                #Store machine gridview
                $Store_Mach_Grid= New-Object system.Windows.Forms.DataGridView
                $Store_Mach_Grid.width= 543
                $Store_Mach_Grid.height= 215
                $Store_Mach_Grid.ReadOnly= $true
                $Store_Mach_Grid.AutoSizeColumnsMode= 16
                $Store_Mach_Grid.AllowUserToDeleteRows= $false;
                $Store_Mach_Grid.AllowUserToResizeColumns= $false;
                $Store_Mach_Grid.MultiSelect= $false;
                $Store_Mach_Grid.AllowUserToAddRows=$false
                $Store_Mach_Grid.location= New-Object System.Drawing.Point(-1,313)
                                                                                                  
                #Load form variables to show all functions, set as an array; if you create/add a variable that needs to be visiableto the form (i.etextboxes, buttons, comboboxes, labels, etc) please add that variable to this list to show it.
$Form.controls.AddRange(@($Print_Queue_button,$Manual_RemoteControl_Label,$Manual_RemoteControl,$Duplicate_GridDataTable,$Script_Duplicate,$Store_Duplicate_Label,$Duplicate_Select_Button,$Duplicate_Employee_Value,$Employee_Selection_Filter,$Global:Employee_Lookup_Valueform,$QuickLinks_ComboBox,$Website_Drop,$Store_Mach_Grid,$users_Lookup_Grid,$Remote_Control_Button,$Ping_Button,$Employee_Lookup_TextBox,$Employee_Lookup_label,$AD_Account_Unlock,$AD_Account_Password_Reset,$Label2,$Quick_Links_Label,$Global:Store_Lookup_textBox,$Store_Lookup_Label,$Store_Users_GridView))
                #Sets the form border style.
                $Form.FormBorderStyle= 'Fixed3D'
                #disables the main form to be resized. This prevents scaling issues.
                $Form.MaximizeBox= $false
                #Console log that pulls the date and time when the applicaitonis launch and gives statuses on the DHCP import. Only way to view this besides PowershellISE or Powershellapp. is through the transcript logs.
                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Grey Matter Core: `"Warn user of wait time`""
                                [System.Windows.Forms.MessageBox]::Show("Please wait for DHCP loading to complete. `nThemain form will show up shortly.", "Please Wait...","OK","Information")
                                While($Global:data.IsCompleted-ne $true){}
                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Closing and Disposing DHCP Job`""
                                $Global:command.EndInvoke($Global:data)
                                $Global:command.Dispose()
                }
                Start-Logging $syncHash.Bug
 
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Grey Matter Core: `"Opening Form`""
                #Launches the form.
                $form.ShowDialog()
}
Main
 
Modules
 

 
 

 
 
 


 
 


