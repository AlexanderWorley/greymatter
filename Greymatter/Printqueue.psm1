Function Open-PrintQueueGUI(){
                $Global:PrintQueueHash= [hashtable]::Synchronized(@{})
                $Global:PrintQueue= New-Object system.Windows.Forms.Form
                $Global:PrintQueue.ClientSize= '712,290'
                $Global:PrintQueue.text= "Store Print Queue Manager"
                $Global:PrintQueue.TopMost= $false
                $Global:PrintQueue.MaximizeBox= $false
                $Global:PrintQueue.BackgroundImage= [system.drawing.image]::FromFile("$Global:rootlocation\Images\bg.jpg")
                $Global:PrintQueue.BackgroundImageLayout= "Stretch"
                $Global:PrintQueue.Icon= "$Global:rootLocation\Images\BlankIcon.ico"
                $Global:PrintQueue.MaximizeBox= $false
                $Global:PrintQueue.FormBorderStyle= 'Fixed3D'
 
                $Script:StoreNumberTB= New-Object system.Windows.Forms.TextBox
                $Script:StoreNumberTB.multiline= $false
                $Script:StoreNumberTB.width= 100
                $Script:StoreNumberTB.height= 20
                $Script:StoreNumberTB.location= New-Object System.Drawing.Point(14,34)
                $Script:StoreNumberTB.Font= 'Microsoft Sans Serif,10'
 
                $Script:StoreTotalPB= New-Object system.Windows.Forms.ProgressBar
                $Script:StoreTotalPB.width= 690
                $Script:StoreTotalPB.height= 14
                $Script:StoreTotalPB.location= New-Object System.Drawing.Point(10,268)
                $Script:StoreTotalPB.Step= 1
 
                $Global:StoreMachinePrintDGV= New-Object system.Windows.Forms.DataGridView
                $Global:StoreMachinePrintDGV.width= 690
                $Global:StoreMachinePrintDGV.height= 180
                $Global:StoreMachinePrintDGV.AutoSizeColumnsMode= 16
                $Global:StoreMachinePrintDGV.location= New-Object System.Drawing.Point(10,71)
                $Global:StoreMachinePrintDGV.SelectionMode= 1
                $Global:StoreMachinePrintDGV.MultiSelect= $false
                $Global:StoreMachinePrintDGV.AllowUserToAddRows=$false
                $Global:StoreMachinePrintDGV.ReadOnly= $true
 
                $Script:ClearStationBTN= New-Object system.Windows.Forms.Button
                $Script:ClearStationBTN.text= "Clear Station"
                $Script:ClearStationBTN.width= 99
                $Script:ClearStationBTN.height= 20
                $Script:ClearStationBTN.location= New-Object System.Drawing.Point(159,34)
                $Script:ClearStationBTN.Font= 'Microsoft Sans Serif,10'
 
                $Script:ClearStoreBTN= New-Object system.Windows.Forms.Button
                $Script:ClearStoreBTN.text= "Clear Store"
                $Script:ClearStoreBTN.width= 99
                $Script:ClearStoreBTN.height= 20
                $Script:ClearStoreBTN.location= New-Object System.Drawing.Point(276,34)
                $Script:ClearStoreBTN.Font= 'Microsoft Sans Serif,10'
 
                $StoreNumberLB= New-Object system.Windows.Forms.Label
                $StoreNumberLB.text= "Store Number"
                $StoreNumberLB.AutoSize= $true
                $StoreNumberLB.width= 25
                $StoreNumberLB.height= 10
                $StoreNumberLB.location= New-Object System.Drawing.Point(14,15)
                $StoreNumberLB.Font= New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]::Bold)
                $StoreNumberLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
 
                $StoreTotalLB= New-Object system.Windows.Forms.Label
                $StoreTotalLB.text= "Store Total Progress"
                $StoreTotalLB.AutoSize= $true
                $StoreTotalLB.width= 25
                $StoreTotalLB.height= 10
                $StoreTotalLB.location= New-Object System.Drawing.Point(14,254)
                $StoreTotalLB.Font= 'Microsoft Sans Serif,10'
                $StoreTotalLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
 
$Global:PrintQueue.controls.AddRange(@($Script:StoreNumberTB,$StoreNumberLB,$Global:StoreMachinePrintDGV,$Script:StoreTotalPB,$StoreTotalLB,$Script:ClearStationBTN,$Script:ClearStoreBTN))
 
                #region GUIEvents
                $ClearStoreBTN_Click={
                                if($Global:PrintQueueHash.Store-ne $null){
                                                if([System.Windows.Forms.MessageBox]::Show("Do you wish to clear store $($Global:PrintQueueHash.Store) entire store's print queue?","ClearWhole Store Queue","YesNo","Question") -eq 'Yes'){
                                                                Clear-StoreQueue
                                                }
                                                else{
                                                                [System.Windows.Forms.MessageBox]::Show("Store $($Global:PrintQueueHash.Store) has been skipped. `nNoactions were taken against the stores print queue.", "Clear Store Skipped","OK","Information")
                                                }
                                }
                                else{
                                                [System.Windows.Forms.MessageBox]::Show("Please enter a store number.", "No Store Detected","OK","Information")
                                }
                }
                $ClearStationBTN_Click={
                                $station = $Global:StoreMachinePrintDGV.SelectedRows.Cells[0].Value
                                If([System.Windows.Forms.MessageBox]::Show("Do you wish to clear $($station)'s print queue?","ClearSingle Station Queue","YesNo","Question") -eq 'Yes'){
                                                Clear-StationQueue
                                }
                                else{
                                                [System.Windows.Forms.MessageBox]::Show("$($station) has been skipped. `nNoactions were taken against the stores print queue.", "Clear Store Skipped","OK","Information")
                                }
                }
                $StoreNumberTB_KeyDown={
                                If ($_.KeyCode-eq "Return" -or $_.KeyCode-eq "Enter") {
                                                $_.SuppressKeyPress= $true
                                                #Setup a trap for none number and too large or too small
                                           If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Script:StoreNumberTB.Text)']") -eq $null){
                                                                [System.Windows.Forms.MessageBox]::Show("$($Script:StoreNumberTB.Text) does not meet the store number standard. `nPleasecheck the text box and try again.", "Invalid Store Number Entry","OK","Error")
                                                }
                                                else{
                                                                Get-StorePrintQueues-StoreNumber$($Script:StoreNumberTB.Text)
                                                }
                                }
                }
                #endregion GUIEvents
   
    $Script:ClearStoreBTN.Add_Click($ClearStoreBTN_Click)
                $Script:ClearStationBTN.Add_Click($ClearStationBTN_Click)
                $Script:StoreNumberTB.Add_KeyDown($StoreNumberTB_KeyDown)
 
                If(-not ($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Global:Store_Lookup_textBox.Text)']") -eq $null)){
                                If(-not ($Global:Store_Lookup_textBox.Text-eq "" -or $Global:Store_Lookup_textBox.Text-eq $null)){
                                                Get-StorePrintQueues-StoreNumber$($Global:Store_Lookup_textBox.Text)
                                                $Script:StoreNumberTB.Text= $Global:Store_Lookup_textBox.Text
                                }
                }
                $Global:PrintQueueHash.Form= $Global:PrintQueue
                $Global:PrintQueueHash.Button= $Script:ClearStoreBTN
                $Global:PrintQueue.Show()
}
 
Function Get-StorePrintQueues(){
Param(
    [Parameter(Mandatory=$true, Position=1)]
    [String]$StoreNumber
)
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Start - $StoreNumber`""
#Variable declaration
$Global:PrintQueueHash.StoreMachDS= New-Object System.Data.DataSet
$stationsDt= New-Object System.Data.DataTable
$Global:PrintQueueHash.Store= $StoreNumber
$Script:StoreTotalPB.Value= 0
#Adding columns to the datatable
$stationsDt.Columns.Add("Station Name") | Out-Null
$stationsDt.Columns.Add("Step") | Out-Null
$stationsDt.Columns.Add("Status") | Out-Null
$stationsDt.Columns.Add("Error") | Out-Null
#Pull the OU string from the config file, all stores will have an OU setup for the store
$tSearchbase= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").StationOU
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: OU Stations Searchbase= $tSearchbase"
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Executing stations search - $StoreNumber`""
#Getting all of the stations listed in the primary OU and storeingit in an array
$tStations= Get-ADComputer-SearchBase$tSearchbase-Filter * | Sort-Object -Property Name
#Getting the MWS OU from the config file and checking its value
If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").MwsOU -eq ""){
                #If there is nothing in the value but "", which for JSON is not null, No separate MWS OU exists for that store
                $mSearchbase= "No MWS OU Found"
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: OU MWS Searchbase= $mSearchbase"
    $Global:PrintQueueHash.StoreMachList= $tStations
}
else{
                #If any value returns, pull the value and sets the search base value
                $mSearchbase= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").MwsOU
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: OU MWS Searchbase= $mSearchbase"
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Executing MWS search - $StoreNumber`""
                #Pulls the mwsbased on the OU provided by the config file
                $mStations= Get-ADComputer-SearchBase$mSearchbase-Filter * | Sort-Object -Property Name
    $Global:PrintQueueHash.StoreMachList= $tStations+ $mStations
}
ForEach($station in $Global:PrintQueueHash.StoreMachList){
    $stationsDt.Rows.Add($station.Name, "N/A", "Not Started","") | Out-Null
}
 
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Exporting Data - $StoreNumber`""
#Adds the datatableto a dataset for interaction with the datagridview
$Global:PrintQueueHash.StoreMachDS.Tables.Add($stationsDt) | Out-Null
                               
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Stop - $StoreNumber`""
#Returns the dataset
$Global:StoreMachinePrintDGV.DataSource= $Global:PrintQueueHash.StoreMachDS.Tables[0]
$Global:StoreMachinePrintDGV.Update()
}
 
Function Clear-StoreQueue(){
                $Global:PrintQueue.Controls.Remove($Script:ClearStoreBTN)
                #This is meant to close the previous thread, try to prevent the hogging of resources
                If($Global:PrintQueueHash.ThreadID-ne $null){
                                $Kill = Get-Runspace-ID $Global:PrintQueueHash.ThreadID
                                $Kill.Close()
                }
                #Set the max value of store total progress
                $Script:StoreTotalPB.Value= 0
                $Script:StoreTotalPB.Maximum= $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows.Count
                #Creates a single new runspace
                $StorePrintRunspace=[runspacefactory]::CreateRunspace()
                #Sets the runspaceto use a single threaded apartment
                $StorePrintRunspace.ApartmentState= "STA"
                #Sets the runspaceto create a new thread that can be re-used for any subsequent invocations
                $StorePrintRunspace.ThreadOptions= "ReuseThread"      
                #Opens the new runspace
                $StorePrintRunspace.Open()
                #Storing the ID for later
                $Global:PrintQueueHash.ThreadID= $StorePrintRunspace.ID
                $Global:PrintQueueHash.ProgressBar= $Script:StoreTotalPB
                #Sets the synchronized hash table that we created earlier to a new variable called syncHashthat is accessible in the new runspace
                $StorePrintRunspace.SessionStateProxy.SetVariable("PrintQueueHash",$Global:PrintQueueHash)
                #Created a new powershellscript to invoke in the new runspace
                $StorePrintScript= [Powershell]::Create().AddScript({
                                $Sessionstate= [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                                $pool = [RunspaceFactory]::CreateRunspacePool(1, 3, $Sessionstate, $Host)
                                #Sets the pool to be multithreaded
                                $pool.ApartmentState= "MTA"
                                $pool.Open()
                                $runspaces= @()
 
                                #Create reusable scriptblock. This is the workhorse of the runspace
                                $RunStation= {
                                                Param (
                                                                $station,
                                                                $hash
                                                )
                                                $x = $hash.StoreMachDS.Tables[0].'Station Name'.IndexOf($station)
                                                #Connect to the station and stop the service
                                                $Status = "Running"
                                                $printError= ""
                                                # STOPPING SPOOLER SERVICE
                                                invoke-command -computername$station {
 
                                                                Set-Service spooler -StartupTypeDisabled
                                                                Stop-Service spooler -force
 
                                                } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
                                                if($errortext-match "WinRM"){
                                                                $printError= "Unable to connect"
                                                                $Status = "Failed"
                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                                                Break
                                                }
 
                                                else{
                                                                If($errortext){
                                                                                $printError= "Failed to stop Spooler"
                                                                                $Status = "Failed"
                                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                                                $hash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                                                                Break
                                                                }
                                                }
                                               
                                                $Step = "Stopping Service"
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                               
                                                #Remove the print jobs file
                                                $Status = "Running"
                                                $printError= ""
                                                # CLEAR PRINT QUEUE
                                                invoke-command -computername$station {
                                                               
                                                                Remove-Item C:\Windows\System32\spool\PRINTERS\*
                                                               
                                                } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
                                                if($errortext){
                                                                $Step = "Deleting Job File"
                                                                $Status = "Failed"
                                                                $printError= "Could not delete print jobs from the spooler folder"
                                                }
                                                else{
                                                                $Step = "Print Job Queue cleared"
                                                }
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                               
                                                #Start the print spooler service
                                                $Step = "Starting Service"
                                                $printError= ""
                                                # START SPOOLER SERVICE
                                                invoke-command -computername$station {
 
                                                                set-service spooler -StartupTypeAutomatic
                                                                start-service spooler
 
                                                } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
                                                if($errortext){
                                                                $Status = "Failed"
                                                                $printError= "Could not start or enable the spooler"
                                                }
                                                else{
 
                                                                $Step = "Spooler service enabled and started"
                                                                $Status = "Complete"
 
                                                }
                                                #Start the print spooler service
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                $hash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                }
 
                                #Create runspaceper ipand add to runspacepool
                                ForEach($machine in $PrintQueueHash.StoreMachDS.Tables[0]){
 
                                                $runspace= [PowerShell]::Create()
                                                #Adde the script block defiled earlier to the new runspace
                                                $null = $runspace.AddScript($RunStation)
                                                $null = $runspace.AddArgument($machine.'StationName')
                                                $null = $runspace.AddArgument($Global:PrintQueueHash)
                                                #Tie the runspacepool defined earlier to the scripts runspacepool
                                                $runspace.RunspacePool= $pool
 
                                                #Add runspaceto runspacescollection and "start" it
                                                #Asynchronously runs the commands of the PowerShell object pipeline
                                                $runspaces+= [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
                                }
 
                                while ($runspaces.Status-ne $null)
                                {
                                                #Clean up, any time a runspacefinishes it's script, add it to an array to be cycled through, end the invoke, retrieve the returned arrays, and do this over and over till there are no runspacesleft in the pool
                                                $completed = $runspaces| Where-Object { $_.Status.IsCompleted-eq $true }
                                                foreach ($runspacein $completed)
                                                {
                                                                #Collect the returned arrays
                                                                $results = $runspace.Pipe.EndInvoke($runspace.Status)
                                                                $PrintQueueHash.ProgressBar.PerformStep()
                                                                $runspace.Status= $null
                                                }
                                }
                                #Adds the button back after all of this is done
                                $Global:PrintQueueHash.Form.Controls.Add($Global:PrintQueueHash.Button)
                                #Close and dispose of the pool
                                $pool.Dispose()
                })
                #Ties the script to the new runspaceand begins the invoke
                $StorePrintScript.Runspace= $StorePrintRunspace
                $Global:StorePrintData= $StorePrintScript.BeginInvoke()
}
 
Function Clear-StationQueue(){
                $station = $Global:StoreMachinePrintDGV.SelectedRows.Cells[0].Value
                $x = $Global:PrintQueueHash.StoreMachDS.Tables[0].'Station Name'.IndexOf($station)
                $Script:StoreTotalPB.Value= 0
                $Script:StoreTotalPB.Maximum= 1
                #Connect to the station and stop the service
                $Status = "Running"
                $printError= ""
    # STOPPING SPOOLER SERVICE
    invoke-command -computername$station {
 
        Set-Service spooler -StartupTypeDisabled
        Stop-Service spooler -force
 
    } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
    if($errortext-match "WinRM"){
        $printError= "Unable to connect"
                                $Status = "Failed"
    }
 
    else{
        If($errortext){
            $printError= "Failed to stop Spooler"
                                                $Status = "Failed"
                                                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                                                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                                                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                                                $Script:StoreTotalPB.PerformStep()
                                                Break
        }
    }
   
    $Step = "Stopping Service"
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
               
                #Remove the print jobs file
                $Status = "Running"
                $printError= ""
    # CLEAR PRINT QUEUE
    invoke-command -computername$station {
       
        Remove-Item C:\Windows\System32\spool\PRINTERS\*
       
    } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
    if($errortext){
                                $Step = "Deleting Job File"
                                $Status = "Failed"
        $printError= "Could not delete print jobs from the spooler folder"
    }
    else{
        $Step = "Print Job Queue cleared"
    }
               
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
               
                #Start the print spooler service
    $Step = "Starting Service"
                $printError= ""
    # START SPOOLER SERVICE
    invoke-command -computername$station {
 
        set-service spooler -StartupTypeAutomatic
        start-service spooler
 
    } -ErrorVariableerrortext-ErrorActionSilentlyContinue
 
    if($errortext){
                                $Status = "Failed"
        $printError= "Could not start or enable the spooler"
    }
    else{
 
        $Step = "Spooler service enabled and started"
                                $Status = "Complete"
 
    }
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Step = $Step
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Status = $Status
                $Global:PrintQueueHash.StoreMachDS.Tables[0].Rows[$x].Error = $printError
                $Script:StoreTotalPB.PerformStep()
}