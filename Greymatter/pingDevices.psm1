Function Get-DHCPList(){
<#
.DESCRIPTION
Gathers all avauilablescopes from the DHCP server
 
.SYNOPSIS
Creates and invokes a new thread for gathering all of the available scopes on the DHCP server named l1wpdhcp01.corp.checksmart.com
#>
Param (
    [switch]$Single,
    [switch]$Multi
)
    If($Single.IsPresent){
                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating DHCP Job`""
        #$job = Get-DhcpServerv4Scope -ComputerNamel1wpdhcp01.corp.checksmart.com
 
        $ServerPing= New-Object System.Data.DataTable
        $ServerPing.Columns.Add("Status") | Out-Null
        $ServerPing.Columns.Add("Server") | Out-Null
        $ServerPing.Columns.Add("Order") | Out-Null
   
        #Create and open runspacepool, setup runspacesarray with min and max set to the number of machine processors plus 1
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $env:NUMBER_OF_PROCESSORS+1)
        #Sets the pool to be multithreaded
        $pool.ApartmentState= "MTA"
        $pool.Open()
        $runspaces= @()
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating repeatable ping script`""
        #Create reusable scriptblock. This is the workhorse of the runspace
        $scriptblock= {
            Param (
            [string]$ipAddress,
            [string]$order
            )
            #Ping the DHCP server
            #Return true or false and server name as array
            $pingResult= ping $ipAddress-n 2 -w 1000
                        #If we get a reply, output that data, other wiseit could not be pinged
            if($pingResult[3].Contains("Reply")){
                                        $Ping = $true
                $log = "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Server $($ipAddress) is online`""
            }
            else{
                $Ping = $false
                $log = "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Server $($ipAddress) is offline`""
                [System.Windows.Forms.MessageBox]::Show("The DHCP Server $ipAddressis offline. `nPleaserestart the app and try again. `nIfthis persists, reach out to a lead.", "DHCP Server Offline","OK","Error")
            }
                        #Take all of the data and put it into the array defined earlier and return it
            $pingHold= ($ping,$ipAddress,$order,$log)
            Return $pingHold
        }
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Giving script specific server data to run and starting script`""
        #Create runspaceper ipand add to runspacepool
        ForEach($server in $Global:storeConfig.SelectNodes("//Settings/Servers/DHCPServer")) {
 
            $runspace= [PowerShell]::Create()
                        #Adde the script block defiled earlier to the new runspace
            $null = $runspace.AddScript($scriptblock)
                        #Tie $pingIPto $server name in the script block
            $null = $runspace.AddArgument($Server.Name)
            #Adding in the order for keeping track of the pings
            $null = $runspace.AddArgument($Server.Order)
                        #Tie the runspacepool defined earlier to the scripts runspacepool
            $runspace.RunspacePool= $pool
 
                        #Add runspaceto runspacescollection and "start" it
            #Asynchronously runs the commands of the PowerShell object pipeline
            $runspaces+= [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Waiting for all scripts to stop and exporting`""
        #Wait for runspacesto finish
        while ($runspaces.Status-ne $null)
        {
            #Clean up, any time a runspacefinishes it's script, add it to an array to be cycled through, end the invoke, retrieve the returned arrays, and do this over and over till there are no runspacesleft in the pool
            $completed = $runspaces| Where-Object { $_.Status.IsCompleted-eq $true }
            foreach ($runspacein $completed)
            {
                                        #Collect the returned arrays
                $results = $runspace.Pipe.EndInvoke($runspace.Status)
                $ServerPing.Rows.Add($results[0],$results[1],$results[2]) | Out-Null
                Start-Logging $results[3]
                #Set the runspacestatus to null
                $runspace.Status= $null
            }
        }
        #Close and dispose of the pool
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Closing pool`""
        $pool.Dispose()
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Sorting Server order and finding first available server`""
        #Sorts the data table and saves it
        $ServerPing= $ServerPing| Sort-Object Order
        ForEach($Server in $ServerPing){
            #Takes the current row, called server and checks the status returned from ping, true or false, and breaks the for loop so the thelast known pingable server is returned
            If($Server.Status){
               Break
            }
        }
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating DHCP Job`""
        #Creates a synchronized hash table to be used to pass variables between the threads
        $Global:syncHash= [hashtable]::Synchronized(@{})
        #Sets the server target given the last known pingable server's name
        $Global:syncHash.Server= $Server.Server
        #Creates a single new runspace
                    $Global:newDHCPRunspace=[runspacefactory]::CreateRunspace()
        #Sets the runspaceto use a single threaded apartment
                    $Global:newDHCPRunspace.ApartmentState= "STA"
                    #Sets the runspaceto create a new thread that can be re-used for any subsequent invocations
        $Global:newDHCPRunspace.ThreadOptions= "ReuseThread"      
                    #Opens the new runspace
        $Global:newDHCPRunspace.Open()
                    #Tying the json settings to the sync has table
                    $Global:syncHash.Settings= $Global:storeConfig.Settings
        #Sets the synchronized hash table that we created earlier to a new variable called syncHashthat is accessible in the new runspace
                    $Global:newDHCPRunspace.SessionStateProxy.SetVariable("syncHash",$Global:syncHash)
                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating DHCP Script`""
        #Created a new powershellscript to invoke in the new runspace
                    $Global:command= [Powershell]::Create().AddScript({
            $syncHash.List= Get-DhcpServerv4Scope -ComputerName$syncHash.Server
        })
                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Executing DHCP Job`""
        #Ties the script to the new runspaceand begins the invoke
                    $Global:command.Runspace= $Global:newDHCPRunspace
        $Global:data= $Global:command.BeginInvoke()
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Multi server complete`""
    }
    ElseIf($Multi.IsPresent){
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating DHCP Job`""
        #$job = Get-DhcpServerv4Scope -ComputerNamel1wpdhcp01.corp.checksmart.com
        
        $Global:syncHash= [hashtable]::Synchronized(@{})
        #$Global:storeConfig.Settings.Servers.DHCPServer.Region will list all of the regions and will have duplicates, the foreach below will remove all duplicates
        $Temp = $Global:storeConfig.Settings.Servers.DHCPServer.Region
        $Sections = @()
        #Used to get a starting region than other wise compares case insensativelythe regions
        ForEach($reg in $Temp){
            If($Sections.Count-eq 0){
                $Sections += $reg
            }
            else{
                $present = $false
                ForEach($Region in $Sections){
                    if($reg -ieq$Region){
                        $present = $true
                        Break
                    }
                }
                If(-not $present){
                    $Sections += $reg
                }
            }
        }
        ForEach($Region in $Sections){
            $Global:syncHash."$($Region)" = New-Object System.Data.DataTable
            $Global:syncHash."$($Region)".Columns.Add("Status") | Out-Null
            $Global:syncHash."$($Region)".Columns.Add("Server") | Out-Null
            $Global:syncHash."$($Region)".Columns.Add("Order") | Out-Null
 
        }
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Repatingscript in new thread per DHCP server`""
        #Create runspaceper ipand add to runspacepool
        ForEach($server in $Global:storeConfig.SelectNodes("//Settings/Servers/DHCPServer")) {
                                                #Add each server to each region list
                                                $Global:syncHash."$($Server.Region)".Rows.Add($true,$Server.Name,$Server.Order) | Out-Null
        }
        #Goes through each region and gets the first available server
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Sorting Server order and finding first available server per region`""
        ForEach($Region in $Sections){
            #Takes each regions datatableand sorts it and saves it
            $Global:syncHash."$($Region)" = $Global:syncHash."$($Region)" | Sort-Object Order
            #Taks the current row, called server and checks the status returned from ping, true or false
            ForEach($DHCP in $Global:syncHash."$($Region)"){
                If($DHCP.Status){
                    #Sets the first pingable server and sets it name assideto be ran later
                    $Global:syncHash."$($Region)Run" = $DHCP.Server
                    Break
                }
            }
        }
 
        #Create a new runspacesarray
                                #Create and open runspacepool, setup runspacesarray with min and max set to the number of machine processors plus 1
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $env:NUMBER_OF_PROCESSORS+1)
       #Sets the pool to be multithreaded
        $pool.ApartmentState= "MTA"
        $pool.Open()
        $runspaces= @()
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Creating repeatable ping script per region`""
        #Create reusable scriptblock. This is the workhorse of the runspace
        $scriptblockdhcp= {
            Param (
            [string]$dhcpServer,
            [string]$region
            )
 
            $temp = Get-DhcpServerv4Scope -ComputerName$dhcpServer
            $hold = ($temp,$region)
 
            Return $hold
        }
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Getting DHCP list per region`""
        #Create runspaceper ipand add to runspacepool
        ForEach($Region in $Sections) {
 
            $runspace= [PowerShell]::Create()
                        #Adde the script block defiled earlier to the new runspace
            $null = $runspace.AddScript($scriptblockdhcp)
                        #Tie regions server to be run to the
            $null = $runspace.AddArgument($Global:syncHash."$($Region)Run")
            #Tie the region to tothe argument to keep track of it all later
            $null = $runspace.AddArgument($Region)
                        #Tie the runspacepool defined earlier to the scripts runspacepool
            $runspace.RunspacePool= $pool
 
                        #Add runspaceto runspacescollection and "start" it
            #Asynchronously runs the commands of the PowerShell object pipeline
            $runspaces+= [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }
 
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Waiting for all scripts to stop and exporting`""
        #Wait for runspacesto finish
        while ($runspaces.Status-ne $null)
        {
            #Clean up, any time a runspacefinishes it's script, add it to an array to be cycled through, end the invoke, retrieve the returned arrays, and do this over and over till there are no runspacesleft in the pool
            $completed = $runspaces| Where-Object { $_.Status.IsCompleted-eq $true }
            foreach ($runspacein $completed)
            {
                                        #Collect the returned arrays
                $results = $runspace.Pipe.EndInvoke($runspace.Status)
                #Output the DHCP list to each region
                $Global:syncHash."$($results[1])List" = $results[0]
                #Set the runspacestatus to null
                $runspace.Status= $null
            }
        }
 
        #Close and dispose of the pool
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Closing pool`""
        $pool.Dispose()
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-DHCPLIST: `"Multi server complete`""
    }
}
 
Function Get-StoreScope(){
<#
.DESCRIPTION
Search for a stores IPv4 IP scope
 
.SYNOPSIS
Searches and returns the stores IPv4 scope that is used by DHCP to hand out IPv4 address's
 
.PARAMETER - StoreNumber
Required string value to be passed to the function that will be used to search for the stores IP scope
#>
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]$StoreNumber
    )
    process {
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Start - $StoreNumber`""
                                #Gets the stores DHCP scope name from the config file
                                $storeName= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPScopeName
        $region = $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").Region
        If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
            Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Searching - $StoreNumber`""
                                    #Goes through the syncHashlist made from the Get-DHCPListfunction of all of the active DHCP scopes and returns the DHCP scope once found.
                                    ForEach($store in $Global:syncHash.List){
                If($store.Name.Contains($storeName))
                {
                                                                    #Checks to see if any exceptions exist for this store, if none exist, return the first found ipscope, otherwise check the exception list
                                                                    If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count -eq 0){
                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                    Return $store.ScopeID
                                                                    }
                                                                    else{
                                                                                    #Checks to see if only one or more exceptions exist, if one we just access the bare exception, otherwise we go through the for loop
                                                                                    If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count -eq 1){
                                                                                                    If(-not ($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.Exception -eq $store.ScopeID)){
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                                                    Return $store.ScopeID
                                                                                                    }
                                                                                    }
                                                                                    else{
                                                                                                    #Sets a variable to be used as false at the beginning
                                                                                                    $exc= $false
                                                                                                    #For loop setup to run through the list of exceptions
                                                                                                    For($x=0;$x -lt$Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count;$x++){
                                                                                                                    #If the exceptions variable is still false, keep looking to see if it is true
                                                                                                                    if($exc-eq $false){
                                                                                                                                    #Replace the booleanvariable with a logical statement that compares the ipaddress of the DHCP scope to the DHCP exception
                                                                                                                                    #If they equal, true will be returned and this variable will not bachanged again no matter how many exceptions are left to check, otherwise the variable is "changed" to false
                                                                                                                                    $exc= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.Exception[$x] -eq $store.ScopeID
                                                                                                                    }
                                                                                                    }
                                                                                                    #So long as the variable is still false, return the current DHCP scope, otherwise we go on to the next scope
                                                                                                    If($exc-eq $false){
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                                                    Return $store.ScopeID
                                                                                                    }
                                                                                    }
                                                                    }
                }
            }
                                    Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Not Found - $StoreNumber`""
        }
        Elseif($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Multi'){
            Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Searching - $StoreNumber`""
                                    #Goes through the syncHashlist made from the Get-DHCPListfunction of all of the active DHCP scopes and returns the DHCP scope once found.
                                    ForEach($store in $Global:syncHash."$($region)List"){
                If($store.Name.Contains($storeName))
                {
                                                                    #Checks to see if any exceptions exist for this store, if none exist, return the first found ipscope, otherwise check the exception list
                                                                    If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count -eq 0){
                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                    Return $store.ScopeID
                                                                    }
                                                                    else{
                                                                                    #Checks to see if only one or more exceptions exist, if one we just access the bare exception, otherwise we go through the for loop
                                                                                    If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count -eq 1){
                                                                                                    If(-not ($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.Exception -eq $store.ScopeID)){
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                                                    Return $store.ScopeID
                                                                                                    }
                                                                                    }
                                                                                    else{
                                                                                                    #Sets a variable to be used as false at the beginning
                                                                                                    $exc= $false
                                                                                                    #For loop setup to run through the list of exceptions
                                                                                                    For($x=0;$x -lt$Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.ChildNodes.Count;$x++){
                                                                                                                    #If the exceptions variable is still false, keep looking to see if it is true
                                                                                                                    if($exc-eq $false){
                                                                                                                                    #Replace the booleanvariable with a logical statement that compares the ipaddress of the DHCP scope to the DHCP exception
                                                                                                                                    #If they equal, true will be returned and this variable will not bachanged again no matter how many exceptions are left to check, otherwise the variable is "changed" to false
                                                                                                                                    $exc= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").DHCPexception.Exception[$x] -eq $store.ScopeID
                                                                                                                    }
                                                                                                    }
                                                                                                    #So long as the variable is still false, return the current DHCP scope, otherwise we go on to the next scope
                                                                                                    If($exc-eq $false){
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Found - $StoreNumber`""
                                                                                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Stop - $StoreNumber`""
                                                                                                                    Return $store.ScopeID
                                                                                                    }
                                                                                    }
                                                                    }
                }
            }
                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - Get-StoreScope: `"Not Found - $StoreNumber`""
                                                Return '169.254.255.255'
        }
    }
}
 
Function Get-StoreDHCPList(){
<#
.DESCRIPTION
Gathers the list of DHCP leases tied to a stores IP scope
 
.SYNOPSIS
Gathers the list of current DHCP leases per the store scope, stores the needed information in a data table, and outputs to a datagridview
 
.PARAMETER - StoreNumber
Required string value to be passed to the function that will be used to search for the stores IP scope
#>
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]$StoreNumber
    )
                #Variable declaration
                $leaseDT= New-Object System.Data.DataTable
                $leaseDS= New-Object System.Data.DataSet
                #Adding the table to the dataset and adding columns to the table
                $leaseDS.Tables.Add($leaseDT)
                $leaseDT.Columns.Add("IP Address") | Out-Null
                $leaseDT.Columns.Add("MAC Address") | Out-Null
                $leaseDT.Columns.Add("Host Name") | Out-Null
                $leaseDT.Columns.Add("Lease State") | Out-Null
                $leaseDT.Columns.Add("Lease Experation") | Out-Null
               
                #Get the stores IP scope
                $ScopeID= Get-StoreScope-StoreNumber$StoreNumber
                #Lists the stores IP leases
                #Getting the IP addresses setup on the DHCP server with the stores specific scope
                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                $storeLeases= Get-DhcpServerv4Lease -ComputerName$Global:syncHash.Server-ScopeId$ScopeID
                }
                Elseif($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Multi'){
                                $region = $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").Region
                                $storeLeases= Get-DhcpServerv4Lease -ComputerName$Global:syncHash."$($region)Run" -ScopeId$ScopeID
                }
                #For loop that goes through each IP lease listed and get the needed info
                ForEach($lease in $storeLeases){
                                $ip= $lease.IPAddress.IPAddressToString
                                $mac = $lease.ClientId.Replace("-", " ")
                                $hostName= $lease.HostName
                                $leaseState= $lease.AddressState
                                $leaseExpMonth= $lease.LeaseExpiryTime.Month
                                $leaseExpDay= $lease.LeaseExpiryTime.Day
                                $leaseExpYear= $lease.LeaseExpiryTime.Year
                                If($lease.LeaseExpiryTime.TimeOfDay.Hours-eq $null){
                                                $leaseExp= ""
                                }
                                ElseIf($lease.LeaseExpiryTime.TimeOfDay.Hours-gt12){
                                                $leaseExp= "$($leaseExpMonth)-$($leaseExpDay)-$($leaseExpYear) $($lease.LeaseExpiryTime.TimeOfDay.Hours- 12):$($lease.LeaseExpiryTime.TimeOfDay.Minutes):$($lease.LeaseExpiryTime.TimeOfDay.Seconds) PM"
                                }
                                Else{
                                                $leaseExp= "$($leaseExpMonth)-$($leaseExpDay)-$($leaseExpYear) $($lease.LeaseExpiryTime.TimeOfDay.Hours):$($lease.LeaseExpiryTime.TimeOfDay.Minutes):$($lease.LeaseExpiryTime.TimeOfDay.Seconds) AM"
                                }
                                $leaseDT.Rows.Add($ip,$mac,$hostName,$leaseState,$leaseExp) | Out-Null
                }
                $Script:groupPingDGV.DataSource= $leaseDS.Tables[0]
}
 
Function Search-StoreComputers(){
<#
.DESCRIPTION
Search for a stores computers
 
.SYNOPSIS
Searches and returns the stores computers and their respective last known leased IP address according to the DHCP server
 
.PARAMETER - StoreNumber
Required string value to be passed to the function that will be used to search for the stores stations based on the OU name equal to the store number
 
.PARAMETER - ScopeID
Required string value to be passed to the function that will be used to search for the stations IP address based on the scope provided
#>
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        $StoreNumber,
        [Parameter(Mandatory=$true, Position=2)]
        $ScopeID
    )
    Process{
                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Start - $StoreNumber`""
        #Variable declaration
                                $storeDs= New-Object System.Data.DataSet
        $stationsDt= New-Object System.Data.DataTable
                                #Adding columns to the datatable
        $stationsDt.Columns.Add("Station Name") | Out-Null
        $stationsDt.Columns.Add("Assigned IP") | Out-Null
                                $stationsDt.Columns.Add("Lease Expiration Date") | Out-Null
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
                                }
                                else{
                                                #If any value returns, pull the value and sets the search base value
                                                $mSearchbase= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").MwsOU
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: OU MWS Searchbase= $mSearchbase"
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Executing MWS search - $StoreNumber`""
                                                #Pulls the mwsbased on the OU provided by the config file
                                                $mStations= Get-ADComputer-SearchBase$mSearchbase-Filter * | Sort-Object -Property Name
                                }
                                If($ScopeID-eq $Null -or $ScopeID-eq '169.254.255.255'){
                                                If($mSearchbase-ne ""){
                                                                ForEach($comp in $mStations){
                                                                                #Giving up the IPs array as the scope was not found
                                                                                $assignip= "No IP Scope Found"
                                                                                $leaseExp= "Please Investigate"
                                                                                $stationsDt.Rows.Add($comp.Name, $assignip, $leaseExp) | Out-Null
                                                                }
                                                }
                                                #Goes through the array if stations and retrieves the IP address assigned to the stations DNS name and gets the lease expire time
                                                ForEach($comp in $tStations){
                                                                #Giving up the IPs array as the scope was not found
                                                                $assignip= "No IP Scope Found"
                                                                $leaseExp= "Please Investigate"
                                                                $stationsDt.Rows.Add($comp.Name, $assignip, $leaseExp) | Out-Null
                                                }
                                }
                                else{
                                                #Getting the IP addresses setup on the DHCP server with the stores specific scope
                                                If($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Single'){
                                                                $ips= Get-DhcpServerv4Lease -ComputerName$Global:syncHash.Server-ScopeId$ScopeID
                                                }
                                                Elseif($Global:storeConfig.Settings.ServerSettings.Mode-eq 'Multi'){
                                                                $region = $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$StoreNumber']").Region
                                                                $ips= Get-DhcpServerv4Lease -ComputerName$Global:syncHash."$($region)Run" -ScopeId$ScopeID
                                                }
                                                #Goes through the array if manager stations, most often 1 or 0, and retrieves the IP address assigned to the stations DNS name and gets the lease expire time
                                                If($mSearchbase-ne ""){
                                                                ForEach($comp in $mStations){
                                                                                #Because the $ipsvariable is an array and not a string, indexofcan notbe case insensitive
                                                                                #Setup a second for each loop to go through each IP in the $ipsarray and compare the string to case insensitive data and if found, break the for loop and add the relevant data, otherwise spit out no found IP
                                                                                ForEach($ipin $ips){
                                                                                                if($ip.HostName-eq $comp.DNSHostName){
                                                                                                                $stationsDt.Rows.Add($comp.Name, $ip.IPAddress.IPAddressToString, $ip.LeaseExpiryTime) | Out-Null
                                                                                                                $x= -0
                                                                                                                break
                                                                                                }
                                                                                                else{
                                                                                                                $x= -1
                                                                                                }
                                                                                }
                                                                                If($x -lt0){
                                                                                                #If we are still at an index of less than 0, station has no assigned IP and has not talked to the DHCP in over 8 days
                                                                                                $assignip= "No IP Assigned"
                                                                                                $leaseExp= "Please Investigate"
                                                                                                $stationsDt.Rows.Add($comp.Name, $assignip, $leaseExp) | Out-Null
                                                                                }
                                                                }
                                                }
                                                #Goes through the array if stations and retrieves the IP address assigned to the stations DNS name and gets the lease expire time
                                                ForEach($comp in $tStations){
                                                                ForEach($ipin $ips){
                                                                                if($ip.HostName-ieq$comp.DNSHostName){
                                                                                                #For each computer IP info found, the info is stored into temp variables and stored in a new row of the datatable
                                                                                                $stationsDt.Rows.Add($comp.Name, $ip.IPAddress.IPAddressToString, $ip.LeaseExpiryTime) | Out-Null
                                                                                                $x= -0
                                                                                                break
                                                                                }
                                                                                else{
                                                                                                $x= -1
                                                                                }
                                                                }
                                                                If($x -lt0){
                                                                                #If we are still at an index of less than 0, station has no assigned IP and has not talked to the DHCP in over 8 days
                                                                                $assignip= "No IP Assigned"
                                                                                $leaseExp= "Please Investigate"
                                                                                $stationsDt.Rows.Add($comp.Name, $assignip, $leaseExp) | Out-Null
                                                                }
                                                }
                                }
 
                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Exporting Data - $StoreNumber`""
        #Adds the datatableto a dataset for interaction with the datagridview
                                $storeDs.Tables.Add($stationsDt) | Out-Null
                               
                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreComputers: `"Stop - $StoreNumber`""
        #Returns the dataset
                                Return $storeDs
    }
}
 
Function Search-StoreScope(){
<#
.DESCRIPTION
Search the entiresstores IP range with a cidrof /24
 
.SYNOPSIS
Given the store scope, it assumes a cidrof /24 and pings the entire subnet from 1-254
 
.PARAMETER - storeScope
The store number that you are wanting to search
#>
Param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline= $true)]
        [String]$storeScope
    )
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Start - $storeScope`""
    #Creating the data set and createinga new datatable. In the new datatable, creating the columns
    $Script:syncPingDS= New-Object System.Data.DataSet
    $Script:syncPingDS.Tables.Add([System.Data.DataTable]::new())
    $Script:syncPingDS.Tables[0].Columns.Add("Status") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("IP_Address") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("Host Name") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("Ping") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("TTL") | Out-Null
    #Tying the new datatableto the datagridviewon the pingergui
    $Script:groupPingDGV.DataSource= $Script:syncPingDS.Tables[0]
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Launching store scope search - $storeScope`""
    #Converting the store number given in $storeScopeand replace it with the network address of that store
    $storeScope= Get-StoreScope-StoreNumber$storeScope
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Setting progressbarstart and stop - $storeScope`""
    #Set the progress bar start and stop values
    $Script:searchRangePB.Maximum= 254
    $Script:searchRangePB.Value= 0
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Creating runspacepool for multithreading - $storeScope`""
    #Create and open runspacepool, setup runspacesarray with min and max set to 100
    $pool = [RunspaceFactory]::CreateRunspacePool(1, 100)
    #Sets the pool to be multithreaded
    $pool.ApartmentState= "MTA"
    $pool.Open()
    $runspaces= @()
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Creating repeatable ping script - $storeScope`""
    #Create reusable scriptblock. This is the workhorse of the runspace
    $scriptblock= {
        Param (
        [string]$ipAddress,
        [string]$scope
        )
                    #This is done this way for proper sorting of ipaddress' when interacting with the datagridviewbecause the IP address is treated as a string and sorts different than numbers
        If($ipAddress.Length-eq 1){
        $ipAddress2 = "00"+$ipAddress
        }
        Elseif($ipAddress.Length-eq 2){
        $ipAddress2 = "0"+$ipAddress
        }
        Else{
        $ipAddress2 = $ipAddress
        }
                    #Create an array used for exporting later. Because the runspaceis apartof a pool, setting a session state variable is not an option
        $pingHold= @()
                    #Creating two ipvariables, $pingIP2 is meant to be the "corrected" string used for sorting later while $pingIPis the unmodified ipaddress used  to make the actual ping cause adding leading 0's when running ping does not work
        $pingIP= $scope.Substring(0,$scope.Length-1) + $ipAddress
        $pingIP2 = $scope.Substring(0,$scope.Length-1) + $ipAddress2
                    #Use ping over Test-connection due to being faster and having more control in powershell5.1, results return an array of strings and through the use of substring and index of, I get the needed data
                    #Ping twice per IP to confirm that the device is truly down and not just the router does not notthe route for the IP
        $pingResult= ping $pingIP-n 2 -w 1000
                    #If we get a reply, output that data, other wiseit could not be pinged
        if($pingResult[3].Contains("Reply")){
                                    #Get the response time of the ping
            $ping = $pingResult[3].Substring($pingResult[3].LastIndexOf("time")+5,$pingResult[3].LastIndexOf("ms")-$pingResult[3].LastIndexOf("time")-3)
            #Get the time to live of the ping
                                    $ttl= $pingResult[3].Substring($pingResult[3].LastIndexOf("TTL")+4,$pingResult[3].Length-$pingResult[3].LastIndexOf("TTL")-4)
            $status = "OK"
            #Get the host name of the ipfrom DNS if one exists, and NSlookupis faster than Resolve-DNSName
                                    $hostNameLookup= nslookup$pingIP
            if($hostNameLookup[3] -eq $null){
                $hostNameLookup= ""
            }
            else{
                $hostNameLookup= $hostNameLookup[3].Substring(9,$hostNameLookup[3].Length-9)
            }
        }
        else{
            $ping = $pingIP
            $ttl= "N/A"
            $status = "Fail"
            $hostNameLookup= "N/A"
        }
                    #Take all of the data and put it into the array defined earlier and return it
        $pingHold= ($status,$pingIP2,$hostNameLookup,$ping,$ttl)
        Return $pingHold
    }
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Giving script specific data to run and starting - $storeScope`""
    #Create runspaceper ipand add to runspacepool
    for($pingIP= 1; $pingIP-lt255; $pingIP++) {
 
        $runspace= [PowerShell]::Create()
                    #Adde the script block defiled earlier to the new runspace
        $null = $runspace.AddScript($scriptblock)
                    #Tie $pingIPto $ipAddressin the script block
        $null = $runspace.AddArgument($pingIP)
                    #Tie the $storeScoperedefined earlier to $scope in the script block
        $null = $runspace.AddArgument($storeScope)
                    #Tie the runspacepool defined earlier to the scripts runspacepool
        $runspace.RunspacePool= $pool
 
                    #Add runspaceto runspacescollection and "start" it
        #Asynchronously runs the commands of the PowerShell object pipeline
        $runspaces+= [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
    }
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Waiting for all scripts to stop and exporting - $storeScope`""
    #Wait for runspacesto finish
    while ($runspaces.Status-ne $null)
    {
        #Clean up, any time a runspacefinishes it's script, add it to an array to be cycled through, end the invoke, retrieve the returned arrays, and do this over and over till there are no runspacesleft in the pool
        $completed = $runspaces| Where-Object { $_.Status.IsCompleted-eq $true }
        foreach ($runspacein $completed)
        {
                                    #Collect the returned arrays
            $results = $runspace.Pipe.EndInvoke($runspace.Status)
                                    #If the option from the menu in pingerto ignore failed pings is, any ping that did not have a reply would not be added to the data table, otherwise all results are added to the table
            if($Script:failedPingsOptionsSubMI.Checked){
                if($results[0].Contains("OK")){
                    $Script:syncPingDS.Tables[0].Rows.Add($results[0],$results[1],$results[2],$results[3],$results[4]) | Out-Null
                }
            }
            else{
                If($results[0] -eq $null){
                    $Script:syncPingDS.Tables[0].Rows.Add("Fail",$results[1],"N/A","N/A","N/A") | Out-Null
                }
                else{
                    $Script:syncPingDS.Tables[0].Rows.Add($results[0],$results[1],$results[2],$results[3],$results[4]) | Out-Null
                }
            }
                                    #Set the runspacestatus to null and increase the progress bar
            $runspace.Status= $null
            $Script:searchRangePB.PerformStep() | Out-Null
        }
    }
    #Close and dispose of the pool
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-StoreScope: `"Stop - $storeScope`""
    $pool.Dispose()
}
 
Function Search-IPRange(){
<#
.DESCRIPTION
Searches the range of IP's given
 
.SYNOPSIS
Takes the start IP and the last octet of the ending IP and uses that to ping the IPs inbetween
 
.PARAMETER - startIP
The start of the iprand and profivedthe net mask, can be no larger than a cidrof /24
 
.PARAMETER - endIP
The final IP to be pinged and takes the final octet of the ipaddress
#>
Param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline= $true)]
        [String]$startIP,
        [Parameter(Mandatory=$true, Position=2, ValueFromPipeline= $true)]
        [String]$endIP
    )
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Start - $startIP`""
    #Creating the data set and createinga new datatable. In the new datatable, creating the columns
    $Script:syncPingDS= New-Object System.Data.DataSet
    $Script:syncPingDS.Tables.Add([System.Data.DataTable]::new())
    $Script:syncPingDS.Tables[0].Columns.Add("Status") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("IP Address") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("Host Name") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("Ping") | Out-Null
    $Script:syncPingDS.Tables[0].Columns.Add("TTL") | Out-Null
    #Tying the new datatableto the datagridviewon the pingergui
    $Script:groupPingDGV.DataSource= $Script:syncPingDS.Tables[0]
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Creating runspacepool for multithreading - $startIP`""
    #Create and open runspacepool, setup runspacesarray with min and max set to 100
    $pool = [RunspaceFactory]::CreateRunspacePool(1, 100)
    #Sets the pool to be multithreaded
    $pool.ApartmentState= "MTA"
    $pool.Open()
    $runspaces= @()
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Creating repeatable ping script - $startIP`""
    #Create reusable scriptblock. This is the workhorse of the runspace
    $scriptblock= {
        Param (
        [string]$ipAddress,
        [string]$scope
        )
                    #This is done this way for proper sorting of ipaddress' when interacting with the datagridviewbecause the IP address is treated as a string and sorts different than numbers
        If($ipAddress.Length-eq 1){
        $ipAddress2 = "00"+$ipAddress
        }
        Elseif($ipAddress.Length-eq 2){
        $ipAddress2 = "0"+$ipAddress
        }
        Else{
        $ipAddress2 = $ipAddress
        }
                    #Create an array used for exporting later. Because the runspaceis apartof a pool, setting a session state variable is not an option
        $pingHold= @()
                    #Creating two ipvariables, $pingIP2 is meant to be the "corrected" string used for sorting later while $pingIPis the unmodified ipaddress used  to make the actual ping cause adding leading 0's when running ping does not work
        $pingIP= $scope + $ipAddress
        $pingIP2 = $scope + $ipAddress2
                    #Use ping over Test-connection due to being faster and having more control in powershell5.1, results return an array of strings and through the use of substring and index of, I get the needed data
                    #Ping twice per IP to confirm that the device is truly down and not just the router does not notthe route for the IP
        $pingResult= ping $pingIP-n 2 -w 1000
                    #If we get a reply, output that data, other wiseit could not be pinged
        if($pingResult[3].Contains("Reply")){
                                    #Get the response time of the ping
            $ping = $pingResult[3].Substring($pingResult[3].LastIndexOf("time")+5,$pingResult[3].LastIndexOf("ms")-$pingResult[3].LastIndexOf("time")-3)
            #Get the time to live of the ping
                                    $ttl= $pingResult[3].Substring($pingResult[3].LastIndexOf("TTL")+4,$pingResult[3].Length-$pingResult[3].LastIndexOf("TTL")-4)
            $status = "OK"
            #Get the host name of the ipfrom DNS if one exists, and NSlookupis faster than Resolve-DNSName
                                    $hostNameLookup= nslookup$pingIP
            if($hostNameLookup[3] -eq $null){
                $hostNameLookup= ""
            }
            else{
                $hostNameLookup= $hostNameLookup[3].Substring(9,$hostNameLookup[3].Length-9)
            }
        }
        else{
            $ping = $pingIP
            $ttl= "N/A"
            $status = "Fail"
            $hostNameLookup= "N/A"
        }
                    #Take all of the data and put it into the array defined earlier and return it
        $pingHold= ($status,$pingIP2,$hostNameLookup,$ping,$ttl)
        Return $pingHold
    }
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Setting progressbarstart and stop - $startIP`""
    #Get the scope, start and stop points
    $scope = $startIP.Substring(0,$startIP.LastIndexOf('.')+1)
    [int]$start = $startIP.Substring($startIP.LastIndexOf('.')+1,$startIP.Length-$startIP.LastIndexOf('.')-1)
    [int]$end = $endIP.Substring($endIP.LastIndexOf('.')+1,$endIP.Length-$endIP.LastIndexOf('.')-1)
 
    #Set the progress bar
    $Script:searchRangePB.Maximum= $end - $start
    $Script:searchRangePB.Value= 0
 
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Giving script specific data to run and starting - $startIP`""
    #Create runspaceper ipand add to runspacepool
    for($pingIP= $start; $pingIP-le $end; $pingIP++) {
 
        $runspace= [PowerShell]::Create()
                    #Adde the script block defiled earlier to the new runspace
        $null = $runspace.AddScript($scriptblock)
                    #Tie $pingIPto $ipAddressin the script block
        $null = $runspace.AddArgument($pingIP)
                    #Tie the $scope redefined earlier to $scope in the script block
        $null = $runspace.AddArgument($scope)
                    #Tie the runspacepool defined earlier to the scripts runspacepool
        $runspace.RunspacePool= $pool
 
                    #Add runspaceto runspacescollection and "start" it
        #Asynchronously runs the commands of the PowerShell object pipeline
        $runspaces+= [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
    }
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Waiting for all scripts to stop and exporting - $startIP`""
    #Wait for runspacesto finish
    while ($runspaces.Status-ne $null)
    {
        #Clean up, any time a runspacefinishes it's script, add it to an array to be cycled through, end the invoke, retrieve the returned arrays, and do this over and over till there are no runspacesleft in the pool
        $completed = $runspaces| Where-Object { $_.Status.IsCompleted-eq $true }
        foreach ($runspacein $completed)
        {
                                    #Collect the returned arrays
            $results = $runspace.Pipe.EndInvoke($runspace.Status)
                                    #If the option from the menu in pingerto ignore failed pings is, any ping that did not have a reply would not be added to the data table, otherwise all results are added to the table
            if($Script:failedPingsOptionsSubMI.Checked){
                if($results -eq $null){}
                elseif($results[0].Contains("OK")){
                    $Script:syncPingDS.Tables[0].Rows.Add($results[0],$results[1],$results[2],$results[3],$results[4]) | Out-Null
                }
            }
            else{
                If($results[0] -eq $null){
                    $Script:syncPingDS.Tables[0].Rows.Add("Fail",$results[1],"N/A","N/A","N/A") | Out-Null
                }
                else{
                    $Script:syncPingDS.Tables[0].Rows.Add($results[0],$results[1],$results[2],$results[3],$results[4]) | Out-Null
                }
            }
                                    #Set the runspacestatus to null and increase the progress bar
            $runspace.Status= $null
            $Script:searchRangePB.PerformStep() | Out-Null
        }
    }
    #Close and dispose of the pool
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Search-IPRange: `"Start - $startIP`""
    $pool.Dispose()
}
 
Function Submit-DNSSearch(){
<#
.DESCRIPTION
Attempts to look up the IP address of a computer name
 
.SYNOPSIS
Either takes the IP address otthe name of the computer and tried to convert it to the other of a DNS record exists
 
.PARAMETER - targetAddress
This is the IP address or the name of the computer that we are searching against
#>
    Param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline= $true)]
        [String]$targetAddress
    )
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"Start - $targetAddress`""
    #Creating the dataset and datatable
                $dnsDS= New-Object System.Data.DataSet
    $dnsDS.Tables.Add([System.Data.DataTable]::new())
    #Adding the columns
    $dnsDS.Tables[0].Columns.Add("Attribute") | Out-Null
    $dnsDS.Tables[0].Columns.Add("Value") | Out-Null
    #DNS querrybased on the target
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"QuerryDNS - $targetAddress`""
                $result = Resolve-DnsName$targetAddress-ErrorActionSilentlyContinue-QuickTimeout
    IF($result -eq $null){
        #will only happen if the address has no DNS record, wheitherby searching by IP or by computer name
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"No Record Found - $targetAddress`""
        $dnsDS.Tables[0].Rows.Add("Target","$targetAddress")
        $dnsDS.Tables[0].Rows.Add("Error","NoRecord Exists")
    }
    ElseIf($result.IPAddress-eq $null){
        #Will only happen if the user searched by the IP address and a record was found
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"Record Found - $targetAddress`""
        $dnsDS.Tables[0].Rows.Add("Target","$targetAddress")
        $dnsDS.Tables[0].Rows.Add("Host Name","$($result.NameHost)")
    }
    else{
        #Will happen if the user searched by the computer name and a record was found
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"Record Found - $targetAddress`""
        $dnsDS.Tables[0].Rows.Add("Target","$targetAddress")
        $dnsDS.Tables[0].Rows.Add("IP Address","$($result.IPAddress)")
    }
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Submit-DNSSearch: `"Stop - $targetAddress`""
    #Set the datasourceon the datagridview
                $Script:nsLookupDGV.DataSource= $dnsDS.Tables[0]
}
 
Function Send-ConstantPing(){
<#
.DESCRIPTION
Pings a single ipuntillit is told to stop
 
.SYNOPSIS
This takes a single IP address and pings it constantly untilltold otherwise and updates a data table. This is done in a sepraterun space to allow continuous action with the GUI
 
.PARAMETER - targetAddress
This the IP address that is the target of the constant ping
#>
    Param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline= $true)]
        [String]$targetAddress
    )
 
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Start - $targetAddress`""
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Create new runspace- $targetAddress`""
                #Creating a synchronized hash table
                $Global:syncConstantPingHash= [hashtable]::Synchronized(@{})
    #Create the single run space
                $pingRunspace=[runspacefactory]::CreateRunspace()
                #Set the run space to run in single thread
    $pingRunspace.ApartmentState= "STA"
                #Sets the runspaceto create a new thread that can be re-used for any subsequent invocations
    $pingRunspace.ThreadOptions= "ReuseThread"
                #Opens the runspace
    $pingRunspace.Open()
                #Sets the variable syncConstantPingHashin the script ran in the new runspaceto allow for communication between the two runspaces
    $pingRunspace.SessionStateProxy.SetVariable("syncConstantPingHash",$syncConstantPingHash)
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Set shared variables - $targetAddress`""
    #Set the singleIPPingBTN(Button) variables text to the .trigger variable in the sync hash table
                $Global:syncConstantPingHash.trigger= $Script:singleIPPingBTN.Text
                #Sets the target address to be used in the ping to the .target variable in the sync hash table
    $Global:syncConstantPingHash.target= $targetAddress
                #Creates a new dataset and ties it to the .pingDSvariable in the sync hash table
    $Global:syncConstantPingHash.pingDS= New-Object System.Data.DataSet
                #Creates a new datatableand adds it to the dataset in the $Global:syncConstantPingHash.pingDS
    $Global:syncConstantPingHash.pingDS.Tables.Add([System.Data.DataTable]::new())
                #Creates the columns of the new datatablecreated earlier
    $Global:syncConstantPingHash.pingDS.Tables[0].Columns.Add("Status") | Out-Null
    $Global:syncConstantPingHash.pingDS.Tables[0].Columns.Add("IP Address") | Out-Null
    $Global:syncConstantPingHash.pingDS.Tables[0].Columns.Add("Ping") | Out-Null
    $Global:syncConstantPingHash.pingDS.Tables[0].Columns.Add("TTL") | Out-Null
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Set dataset to datagridview- $targetAddress`""
                #Sets the datagridviewsdatasourceto the new data table
    $Script:singleIPPingDGV.Datasource= $Global:syncConstantPingHash.pingDS.Tables[0]
    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Create script to run untillstop - $targetAddress`""
                #Creating the powershellscript to be ran in the new runspace
                $Global:pingScript= [Powershell]::Create().AddScript({
                                #The singleIPPingBTNtext is used to start and stop this while loop. When the button is pressed, the text changes to "Stop Ping" and can start this while loop
        while($syncConstantPingHash.trigger-eq "Stop Ping"){
                                                #Pings the target ipaddress and the results are stored in an array of strings
            $result = ping $syncConstantPingHash.Target-n 2 -w 1000
                                                #Depending on the result of the ping, the data table will be updated with one of the 3 responses
            if($result.Contains("Ping request could not find host")){
                $ip= ""
                $ping = "Could Not Resolve"
                $ttl= "N/A"
                $status = "DNS Error"
            }
            elseif($result[3].Contains("Reply")){
                $ip= $result[3].Substring(11,$result[3].IndexOf(":")-11)
                $ping = $result[3].Substring($result[3].LastIndexOf("time")+5,$result[3].LastIndexOf("ms")-$result[3].LastIndexOf("time")-3)
                $ttl= $result[3].Substring($result[3].LastIndexOf("TTL")+4,$result[3].Length-$result[3].LastIndexOf("TTL")-4)
                $status = "OK"
            }
            else{
                $ip= $result[1].Substring($result[1].IndexOf(" ")+1,$result[1].IndexOf("with")-1-$result[1].IndexOf(" ")-1)
                $ping = "Time Out"
                $ttl= "N/A"
                $status = "Fail"
            }
                                                #Prevents the data table from getting too long. If it gets tolong, the GUI becomes unresponsive, idk why
            if($syncConstantPingHash.pingDS.Tables[0].Rows.Count-eq 13){
                $syncConstantPingHash.pingDS.Tables[0].Rows[0].Delete()
            }
                                                #Add the latest result to the end of the conatantping data table
            $syncConstantPingHash.pingDS.Tables[0].Rows.Add($status,$ip,$ping,$ttl) | Out-Null
                                                #Sleeps for half a second between ping attempts
            Start-Sleep -Milliseconds 500
        }
    })
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Start constant ping in seperaterunspace- $targetAddress`""
                #Sets the runespaceto the script
                $Global:pingScript.Runspace= $pingRunspace
    #Starts the ping attempts
                $Global:data= $Global:pingScript.BeginInvoke()
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Send-ConstantPing: `"Stop - $targetAddress`""
}
 
Function Open-PingGUI(){
Param(
        [Parameter(Mandatory=$true, Position=1)]
        $Background
    )
begin{
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Initialize GUI`""
    #variable declaration
    $Script:pingDataSet= New-Object System.Data.DataSet
    $Script:pingDataTable= New-Object System.Data.DataTable
               
                #Everything below is the creation of the ping GUI
    #region Main GUI
                #Pinger Form
    $Global:pingScreen= New-Object System.Windows.Forms.Form
    $Global:pingScreen.Text= "Pinger"
    $Global:pingScreen.Name= "Pinger"
    $Global:pingScreen.Size= "1110,465"
    $Global:pingScreen.BackgroundImage= $Background
    $Global:pingScreen.BackgroundImageLayout= "Stretch"
    $Global:pingScreen.Icon= "$Global:rootLocation\Images\BlankIcon.ico"
    $Global:pingScreen.MaximizeBox= $false
    $Global:pingScreen.FormBorderStyle= 'Fixed3D'
    $Global:pingScreen.Add_Shown({$Global:pingScreen.Activate()})
               
                #region MenuBar
    $menuMS= New-Object System.Windows.Forms.MenuStrip
    $fileDropdownMI= New-Object System.Windows.Forms.ToolStripMenuItem
    $exitFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $resetFileSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsDropdownMI= New-Object System.Windows.Forms.ToolStripMenuItem
    $Script:failedPingsOptionsSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
                $DHCPServerOptionsSubMI= New-Object System.Windows.Forms.ToolStripMenuItem
   
    $menuMS.Items.AddRange(@($fileDropdownMI,$optionsDropdownMI))
    $menuMS.Location= '0,0'
    $menuMS.Size= '475,24'
    $menuMS.TabIndex= '0'
    $Global:pingScreen.Controls.Add($menuMS)
 
    $fileDropdownMI.DropDownItems.AddRange(@($resetFileSubMI,$exitFileSubMI))
    $fileDropdownMI.Size= '35,20'
    $fileDropdownMI.Name= 'fileDropdownMI'
    $fileDropdownMI.Text= "&File"
                $resetFileSubMI.Size= '182,20'
    $resetFileSubMI.Text= "&Reset Broken DataGridView"
    $exitFileSubMI.Size= '182,20'
   $exitFileSubMI.Text= "&Exit"
 
    $optionsDropdownMI.DropDownItems.AddRange(@($Script:failedPingsOptionsSubMI,$DHCPServerOptionsSubMI))
    $optionsDropdownMI.Size= '51,20'
    $optionsDropdownMI.Text= "&Options"
    $Script:failedPingsOptionsSubMI.Size= '250,20'
    $Script:failedPingsOptionsSubMI.Text= "&Ignore Failed Pings"
    $Script:failedPingsOptionsSubMI.CheckOnClick= $true
                $DHCPServerOptionsSubMI.Size= '250,20'
                $DHCPServerOptionsSubMI.Text= "&Change DHCP Server"
               
    [hashtable]$servers =
                $i=1
                ForEach($dhcpServer in $Global:storeConfig.Settings.Servers.ChildNodes){
                               
                }
                #endregion MenuBar
               
                #Textbox for starting IP in IP range search
    $Script:startIPTB= New-Object System.Windows.Forms.Textbox
    $Script:startIPTB.Size= "135,23"
    $Script:startIPTB.Location= "17,61"
    $Script:startIPTB.TabIndex= "1"
    $Global:pingScreen.Controls.Add($Script:startIPTB)
               
                #Textbox for ending IP in IP range search
    $Script:endIPTB= New-Object System.Windows.Forms.Textbox
    $Script:endIPTB.Size= "135,23"
    $Script:endIPTB.Location= "157,61"
    $Script:endIPTB.TabIndex= "2"
    $Global:pingScreen.Controls.Add($Script:endIPTB)
               
                #Textbox for entire store IP range search
    $Script:storeNumberTB= New-Object System.Windows.Forms.Textbox
    $Script:storeNumberTB.Size= "68,16"
    $Script:storeNumberTB.Location= "113,85"
    $Script:storeNumberTB.TabIndex= "4"
    $Global:pingScreen.Controls.Add($Script:storeNumberTB)
               
                #Textbox for DNS lookup
    $Script:nsLookupTB= New-Object System.Windows.Forms.Textbox
    $Script:nsLookupTB.Size= "108,23"
    $Script:nsLookupTB.Location= "535,85"
    $Script:nsLookupTB.TabIndex= "7"
    $Global:pingScreen.Controls.Add($Script:nsLookupTB)
               
                #Textbox for constant ping
    $Script:singleIPTB= New-Object System.Windows.Forms.Textbox
    $Script:singleIPTB.Size= "108,23"
    $Script:singleIPTB.Location= "844,58"
    $Script:singleIPTB.TabIndex= "9"
    $Global:pingScreen.Controls.Add($Script:singleIPTB)
               
                #Label for $Script:startIPTB
    $startIPLB= New-Object System.Windows.Forms.Label
    $startIPLB.Text= "Starting IP"
    $startIPLB.Size= "69,14"
    $startIPLB.Location= "17,48"
    $startIPLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $Global:pingScreen.Controls.Add($startIPLB)
   
                #Label for $Script:endIPTB
    $endIPLB= New-Object System.Windows.Forms.Label
    $endIPLB.Text= "Ending IP"
    $endIPLB.Size= "69,14"
    $endIPLB.Location= "157,48"
    $endIPLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $Global:pingScreen.Controls.Add($endIPLB)
               
                #Label for $Script:storeNumberTB
    $storeNumLB= New-Object System.Windows.Forms.Label
    $storeNumLB.Text= "Store Number:"
    $storeNumLB.Size= "91,14"
    $storeNumLB.Location= "30,90"
    $storeNumLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $Global:pingScreen.Controls.Add($storeNumLB)
               
                #Label for $Script:nsLookupTB
    $nsLookupLB= New-Object System.Windows.Forms.Label
    $nsLookupLB.Text= "IP Address/Domain Name"
    $nsLookupLB.Size= "140,14"
    $nsLookupLB.Location= "535,65"
    $nsLookupLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $Global:pingScreen.Controls.Add($nsLookupLB)
               
                #Label for $Script:singleIPTB
    $ipLB= New-Object System.Windows.Forms.Label
    $ipLB.Text= "IP Address:"
    $ipLB.Size= "72,14"
    $ipLB.Location= "775,60"
    $ipLB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $Global:pingScreen.Controls.Add($ipLB)
   
                #Button for ranged IP search
    $Script:searchRangeBTN= New-Object System.Windows.Forms.Button
    $Script:searchRangeBTN.Text= "Search Range"
    $Script:searchRangeBTN.Size= "90,22"
    $Script:searchRangeBTN.Location= "430,60"
    $Script:searchRangeBTN.TabIndex= "3"
    $Global:pingScreen.Controls.Add($Script:searchRangeBTN)
   
                #Button for store range IP search
    $Script:searchStoreBTN= New-Object System.Windows.Forms.Button
    $Script:searchStoreBTN.Text= "Search Store"
    $Script:searchStoreBTN.Size= "90,22"
    $Script:searchStoreBTN.Location= "430,84"
    $Script:searchStoreBTN.TabIndex= "5"
    $Global:pingScreen.Controls.Add($Script:searchStoreBTN)
               
                #Button for store range IP leases
    $Script:getStoreLeasesBTN= New-Object System.Windows.Forms.Button
    $Script:getStoreLeasesBTN.Text= "Get Store Leases"
    $Script:getStoreLeasesBTN.Size= "110,22"
    $Script:getStoreLeasesBTN.Location= "310,84"
    $Script:getStoreLeasesBTN.TabIndex= "6"
    $Global:pingScreen.Controls.Add($Script:getStoreLeasesBTN)
               
                #Button for DNS lookup
    $Script:nsLookupBTN= New-Object System.Windows.Forms.Button
    $Script:nsLookupBTN.Text= "Lookup"
    $Script:nsLookupBTN.Size= "97,22"
    $Script:nsLookupBTN.Location= "660,84"
    $Script:nsLookupBTN.TabIndex= "8"
    $Global:pingScreen.Controls.Add($Script:nsLookupBTN)
               
                #Button for constant ping
    $Script:singleIPPingBTN= New-Object System.Windows.Forms.Button
    $Script:singleIPPingBTN.Text= "Start Ping"
    $Script:singleIPPingBTN.Size= "97,22"
    $Script:singleIPPingBTN.Location= "978,57"
    $Script:singleIPPingBTN.TabIndex= "10"
    $Global:pingScreen.Controls.Add($Script:singleIPPingBTN)
   
                #Progress bar for the rang/store range search
    $Script:searchRangePB= New-Object System.Windows.Forms.ProgressBar
    $Script:searchRangePB.Size= "100,18"
    $Script:searchRangePB.Location= "192,86"
    $Script:searchRangePB.Step= 1
    $Script:searchRangePB.Value= 0
    $Global:pingScreen.Controls.Add($Script:searchRangePB)
               
                #DataGridView for the ranged pings
    $Script:groupPingDGV= New-Object System.Windows.Forms.DataGridView
    $Script:groupPingDGV.Size= '500,295'
    $Script:groupPingDGV.Location= '17,112'
    $Script:groupPingDGV.ReadOnly= $true
    $Script:groupPingDGV.AutoSizeColumnsMode= 6
                $Script:groupPingDGV.AllowUserToAddRows=$false
    $Global:pingScreen.Controls.Add($Script:groupPingDGV)
               
                #DataGridView for the DNS lookup
    $Script:nsLookupDGV= New-Object System.Windows.Forms.DataGridView
    $Script:nsLookupDGV.Size= '221,295'
    $Script:nsLookupDGV.Location= '535,112'
    $Script:nsLookupDGV.ReadOnly= $true
    $Script:nsLookupDGV.AutoSizeColumnsMode= 6
                $Script:nsLookupDGV.AllowUserToAddRows=$false
    $Global:pingScreen.Controls.Add($Script:nsLookupDGV)
               
                #DataGridView for the constant ping
    $Script:singleIPPingDGV= New-Object System.Windows.Forms.DataGridView
    $Script:singleIPPingDGV.Size= '300,324'
    $Script:singleIPPingDGV.Location= '775,83'
    $Script:singleIPPingDGV.ReadOnly= $true
    $Script:singleIPPingDGV.AutoSizeColumnsMode= 16
                $Script:singleIPPingDGV.AllowUserToAddRows=$false
    $Global:pingScreen.Controls.Add($Script:singleIPPingDGV)
               
                #Group box used to group the various controls for ranged ping
    $pingRangeGB= New-Object System.Windows.Forms.GroupBox
    $pingRangeGB.Size= '516,381'
    $pingRangeGB.Location= '10,32'
    $pingRangeGB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $pingRangeGB.Text= "Search IP Range/Store Number"
    $Global:pingScreen.Controls.Add($pingRangeGB)
               
                #Group box used to group the various controls for NSlookup
    $nsLookupGB= New-Object System.Windows.Forms.GroupBox
    $nsLookupGB.Size= '241,381'
    $nsLookupGB.Location= '525,32'
    $nsLookupGB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $nsLookupGB.Text= "DNS Lookup"
    $Global:pingScreen.Controls.Add($nsLookupGB)
               
                #Group box used to group the various controls for constant ping
    $pingSingleGB= New-Object System.Windows.Forms.GroupBox
    $pingSingleGB.Size= '312,381'
    $pingSingleGB.Location= '765,32'
    $pingSingleGB.BackColor= [System.Drawing.Color]::FromName("Transparent")
    $pingSingleGB.Text= "Ping Single IP Address"
    $Global:pingScreen.Controls.Add($pingSingleGB)
#endregion Main GUI
 
#region events
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Initialize GUI Events`""
    $Script:searchRangeBTN.Add_MouseEnter({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Hand
    })
    $Script:searchRangeBTN.Add_MouseLeave({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Arrow
    })
                #Click event, checks the lenthof the IP anteredand so long as the length is long enough, it will try and run
                $searchRangeBTN_Click= {
                                If($Script:startIPTB.Text.Length-lt8){
                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchRange: `"The IP address $($Script:startIPTB.Text) is too short`""
                                                [System.Windows.Forms.MessageBox]::Show("The IP address $($Script:startIPTB.Text) is too short. `nPleasecheck the starting IP text box and try again.", "Invalid IP Address","OK","Error")
                                }
                                elseif($Script:endIPTB.Text.Length-lt8){
                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchRange: `"The IP address $($Script:endIPTB.Text) is too short`""
                                                [System.Windows.Forms.MessageBox]::Show("The IP address $($Script:endIPTB.Text) is too short. `nPleasecheck the starting IP text box and try again.", "Invalid IP Address","OK","Error")
                                }
                                else{
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start IP Range Search`""
                                                Search-IPRange-startIP$Script:startIPTB.Text-endIP$Script:endIPTB.Text
                                }
    }
                $startEndIPTB_Return= {
                                If ($_.KeyCode-eq "Enter" -or $_.KeyCode-eq "Return") {
                                                $_.SuppressKeyPress= $true
                                                If($Script:startIPTB.Text.Length-lt8){
                                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchRange: `"The IP address $($Script:startIPTB.Text) is too short`""
                                                                [System.Windows.Forms.MessageBox]::Show("The IP address $Script:startIPTB.Textis too short. `nPleasecheck the starting IP text box and try again.", "Invalid IP Address","OK","Error")
                                                }
                                                elseif($Script:endIPTB.Text.Length-lt8){
                                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchRange: `"The IP address $($Script:endIPTB.Text) is too short`""
                                                                [System.Windows.Forms.MessageBox]::Show("The IP address $Script:endIPTB.Textis too short. `nPleasecheck the starting IP text box and try again.", "Invalid IP Address","OK","Error")
                                                }
                                                else{
                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start IP Range Search`""
                                                                Search-IPRange-startIP$Script:startIPTB.Text-endIP$Script:endIPTB.Text
                                                }
                                }
    }
                #The button and the enter button will trigger the same event
                $Script:searchRangeBTN.Add_Click($searchRangeBTN_Click)
                $Script:startIPTB.Add_KeyDown($startEndIPTB_Return)
                $Script:endIPTB.Add_KeyDown($startEndIPTB_Return)
    $Script:searchStoreBTN.Add_MouseEnter({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Hand
    })
    $Script:searchStoreBTN.Add_MouseLeave({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Arrow
    })
                #Triggers the IP ping of the entire store scope so long as the store number is the right length
                $searchStoreBTN_Click={
                           If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Script:storeNumberTB.Text)']") -eq $null){
                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchStoreRange: `"The store number $($Script:storeNumberTB.Text) does not exist`""
                                                [System.Windows.Forms.MessageBox]::Show("$($Script:storeNumberTB.Text) does not meet the store number standard. `nPleasecheck the text box and try again.", "Invalid Store Number Entry","OK","Error")
                                }
                                else{
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start Store Search`""
                                                Search-StoreScope-storeScope$Script:storeNumberTB.Text
                                }
    }
                $storeNumberTB_Return={
                                if ($_.KeyCode-eq "Enter" -or $_.KeyCode-eq "Return") {
                                                $_.SuppressKeyPress= $true
                                           If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Script:storeNumberTB.Text)']") -eq $null){
                                                    Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchStoreRange: `"The store number $($Script:storeNumberTB.Text) does not exist`""
                                                    [System.Windows.Forms.MessageBox]::Show("$($Script:storeNumberTB.Text) does not meet the store number standard. `nPleasecheck the text box and try again.", "Invalid Store Number Entry","OK","Error")
                                    }
                                    else{
                                                    Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start Store DHCP Lease Retrieval`""
                                                                Get-StoreDHCPList-StoreNumber$Script:storeNumberTB.Text
                                    }
                                }
    }
                #The button and the enter button will trigger the same event
    $Script:searchStoreBTN.Add_Click($searchStoreBTN_Click)
                $Script:storeNumberTB.Add_KeyDown($storeNumberTB_Return)
                $Script:nsLookupBTN.Add_MouseEnter({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Hand
    })
    $Script:nsLookupBTN.Add_MouseLeave({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Arrow
    })
                #Triggers the IP ping of the entire store scope so long as the store number is the right length
                $nsLookupBTN_Click={
        Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start DNS Querry`""
                                Submit-DNSSearch-targetAddress$Script:nsLookupTB.Text
    }
                $nsLookupTB_Return={
                                if ($_.KeyCode-eq "Enter" -or $_.KeyCode-eq "Return") {
                                                $_.SuppressKeyPress= $true
            Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start DNS Querry`""
                                                Submit-DNSSearch-targetAddress$Script:nsLookupTB.Text
                                }
    }
                #The button and the enter button will trigger the same event
    $Script:nsLookupBTN.Add_Click($nsLookupBTN_Click)
                $Script:nsLookupTB.Add_KeyDown($nsLookupTB_Return)
                $Script:getStoreLeasesBTN.Add_MouseEnter({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Hand
    })
    $Script:getStoreLeasesBTN.Add_MouseLeave({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Arrow
    })
                $getStoreLeasesBTN_Click={
                           If($Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$($Script:storeNumberTB.Text)']") -eq $null){
                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - SearchStoreLeases: `"The store number $($Script:storeNumberTB.Text) does not exist`""
                                                [System.Windows.Forms.MessageBox]::Show("$($Script:storeNumberTB.Text) does not meet the store number standard. `nPleasecheck the text box and try again.", "Invalid Store Number Entry","OK","Error")
                                }
                                else{
                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start Store DHCP Lease Retrieval`""
                                                Get-StoreDHCPList-StoreNumber$Script:storeNumberTB.Text
                                }
    }
                #This will trigger the DHCP lease retrieval
    $Script:getStoreLeasesBTN.Add_Click($getStoreLeasesBTN_Click)
    $Script:singleIPPingBTN.Add_MouseEnter({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Hand
    })
    $Script:singleIPPingBTN.Add_MouseLeave({
                                #Change the cursor, cosmetic effect
        $Global:pingScreen.Cursor= [System.Windows.Forms.Cursors]::Arrow
    })
                $singleIPPingBTN_Click={
                                #Takes the buttons text value and uses it like a switch, it will either start the ping and lock the text box, or stop the pin and unlock the text box
                                #Each click switches the text value to the opposite switch case trigger, a make shift toggle switch
        switch ($Script:singleIPPingBTN.Text){
            "Start Ping"{
                If($Script:singleIPTB.Text.Length-ge5){
                    $Script:singleIPPingBTN.Text= "Stop Ping"
                    $Script:singleIPTB.ReadOnly= $true
                                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start Constant Ping`""
                    Send-ConstantPing-targetAddress$Script:singleIPTB.Text
                }
                                                                else{
                                                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"The address $($Script:singleIPTB.Text) is too short`""
                                                                                [System.Windows.Forms.MessageBox]::Show("The address $($Script:singleIPTB.Text) is too short. `nPleasecheck the text box and try again.", "Invalid Address Entry","OK","Error")
                                                                }
            }
            "Stop Ping"{
                $Script:singleIPPingBTN.Text= "Start Ping"
                $Script:singleIPTB.ReadOnly= $false
                $Global:syncConstantPingHash.trigger= "Start Ping"
                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Stop Constant Ping`""
                                                                $Global:pingScript.EndInvoke($Global:data)
                                                                $Global:pingScript.Dispose()
            }
        }
    }
                $singleIPTB_Return={
                                if ($_.KeyCode-eq "Enter" -or $_.KeyCode-eq "Return") {
                                                $_.SuppressKeyPress= $true
                                                If($Script:singleIPTB.Text.Length-ge5){
                                                                $Script:singleIPPingBTN.Text= "Stop Ping"
                                                                $Script:singleIPTB.ReadOnly= $true
                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start Constant Ping`""
                                                                Send-ConstantPing-targetAddress$Script:singleIPTB.Text
                                                }
                                                else{
                                                                Start-Logging "ERROR: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"The address $($Script:singleIPTB.Text) is too short`""
                                                                [System.Windows.Forms.MessageBox]::Show("The address $($Script:singleIPTB.Text) is too short. `nPleasecheck the text box and try again.", "Invalid Address Entry","OK","Error")
                                                }
                                }
    }
                #The button and the enter button will trigger the same event
    $Script:singleIPPingBTN.Add_Click($singleIPPingBTN_Click)
                $Script:singleIPTB.Add_KeyDown($singleIPTB_Return)
    $exitFileSubMI.Add_Click({
                                #Closes the window, does the same as the x button
        $Global:pingScreen.Close()
   })
                $resetFileSubMI.Add_Click({
                                $Global:pingScreen.Controls.Remove($Script:singleIPPingDGV)
                                $Script:singleIPPingDGV= $null
                                $Script:singleIPPingDGV= New-Object System.Windows.Forms.DataGridView
                                $Script:singleIPPingDGV.Size= '300,324'
                                $Script:singleIPPingDGV.Location= '775,83'
                                $Script:singleIPPingDGV.ReadOnly= $true
                                $Script:singleIPPingDGV.AutoSizeColumnsMode= 16
                                $Script:singleIPPingDGV.AllowUserToAddRows=$false
                                $Global:pingScreen.Controls.Add($Script:singleIPPingDGV)
                                $Script:singleIPPingDGV.BringToFront()
                })
    $Global:pingScreen.Add_FormClosing({
                                #Sets the GUI to get disposed to free up resources
                                switch ($Script:singleIPPingBTN.Text){
            "Stop Ping"{
                $Script:singleIPPingBTN.Text= "Start Ping"
                $Script:singleIPTB.ReadOnly= $false
                $Global:syncConstantPingHash.trigger= "Start Ping"
                                                                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Stop Constant Ping`""
                                                                $Global:pingScript.EndInvoke($Global:data)
                                                                $Global:pingScript.Dispose()
            }
                                }
        $Global:pingScreen.Dispose()
    })
#endregion events
    }
    Process{
                Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - Open-PingGUI: `"Start GUI`""
                #Show the gui
    $Global:pingScreen.Show()
    }
}