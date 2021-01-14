function updateStoreView{
<#
.DESCRIPTION
Gets a list of employees in a store OU
 
.SYNOPSIS
Given a store number as a string, gather all of the employees in the stores OU and output the releventdata to a dataset
 
.PARAMETER - Store_Lookup_Value
The store number, this is used in the OU search
#>
                param(
                                $Store_Lookup_Value
                )
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Start - $Store_Lookup_Value`""
#Primary store lookup
$Server = 'corp.checksmart.com'
# Main Columns
$Store_Lookup_GridDataTable= New-Object System.Data.DataTable
#Column A
$Store_Lookup_GridDataTable.Columns.Add('Tellers', [string]) | Out-Null
#Column B
$Store_Lookup_GridDataTable.Columns.Add('EmployeeID', [string]) | Out-Null
#Column C
$Store_Lookup_GridDataTable.Columns.Add('Title', [string]) | Out-Null
#Column D
$Store_Lookup_GridDataTable.Columns.Add('Account Status', [string]) | Out-Null
#Column E
$Store_Lookup_GridDataTable.Columns.Add('Locked', [string]) | Out-Null
#Store Search
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Setting AD Scope - $Store_Lookup_Value`""
#Gets the stores users OU based on the config file
$Searchbase= $Global:storeConfig.SelectSingleNode("//Settings/Stores/Store[@Number='$Store_Lookup_Value']").UsersOU
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Pulling AD Objects - $Store_Lookup_Value`""
#Pulls all store users in a OU.            
$Script = get-ADuser-server $server -filter * -searchbase$searchbase-properties *
#Runs $Script through and pulls all attributes
#Runs through all users in a store under the foreach loop.
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Going through AD Users - $Store_Lookup_Value`""
                if($Script.GetType().BaseType.Name-eq 'Array'){
                    ForEach($Employee in $Script){
                                                #Passes the column sets as horizontal.
                                                $Teller = $Employee | Select -ExpandPropertycn
                                                $EmployeeID= $Employee| Select -ExpandPropertyEmployeeID
                                                $Description = $Employee | Select -ExpandPropertyTitle
                                                $EmployeePin=  $Employee | Select -ExpandPropertyenabled
                                                $EmployeeDoB= $Employee | Select -ExpandPropertyLockedout
                                                $Store_Lookup_GridDataTable.rows.Add($Teller, $EmployeeID,$Description,$EmployeePin,$employeeDoB)
                                }
    }else{
                                $Teller = $Script | Select -ExpandPropertycn
                                $EmployeeID= $Script| Select -ExpandPropertyEmployeeID
                                $Description = $Script | Select -ExpandPropertyDescription
                                $EmployeePin=  $Script | Select -ExpandPropertyenabled
                                $EmployeeDoB= $Script | Select -ExpandPropertyLockedout
                                $Store_Lookup_GridDataTable.rows.Add($Teller, $EmployeeID,$Description,$EmployeePin,$employeeDoB)
                }
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Exporting Data - $Store_Lookup_Value`""
#Creates the DataTableand pushes it to the form.
$dataset = New-Object System.Data.DataSet
$dataset.Tables.Add($Store_Lookup_GridDataTable)
Start-Logging "INFO: $(Get-Date -UFormat"%Y-%m-%d %r") - updateStoreView: `"Stop - $Store_Lookup_Value`""
Return $dataset
}