<#   
.SYNOPSIS   
    Interactive menu that allows a user to connect to a remote WSUS server to pull information.
     
.DESCRIPTION 
    Presents an interactive menu for the user to first make a connection to a remote WSUS Server.  After making connection to the machine,  
    the user is asked which group you would like topull inofrmation from.The sript pulls select inofrmation from the WSUS server and exports it to excel
    placed in a table for further sorting. There are 4 reports generated (Update summary, Pending reboot, Needed updates (excludes feature), failed updates)
        
.NOTES   
    Name: WSUS_Report 
    Author: StrayTripod 
    Modifier: StrayTripod
    DateCreated: 7/9//2021
    DateModifed: 7/12/2021
    Updates needed report filters out feature updates.
          
.EXAMPLE     
#> 
Write-Host "WSUS Update Reports"
Write-Host "############################"
Write-Host ""

$wsusserver = Read-Host "Enter the server name to connect to"
$usessl = $False
$Port = 8530
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$usessl,$Port)
##Get updates summary per computer##
$TargetGroup = read-host 'Enter the group name'

$computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$subgroups = $False
$groupid = ($wsus.GetComputerTargetGroups() | ? { $_.Name -eq $TargetGroup}).Id #Lookup the group ID based on name
$group = $wsus.GetComputerTargetGroup($groupid) #Get the group object based on ID
$computerScope.ComputerTargetGroups.Add($group) #Add the target group to the scope
$computerScope.IncludeSubgroups = $subgroups #Set subgroups to $true or $false

Write-Host ""
Write-Host Generating Update Summary for $TargetGroup
Write-Host "################################################"
Write-Host ""
$updatesum=$wsus.GetSummariesPerComputerTarget($updatescope,$computerscope) |
Select-Object @{L='ComputerTarget';E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}}, 
@{L='NeededCount';E={($_.DownloadedCount + $_.NotInstalledCount)}},DownloadedCount,NotInstalledCount,InstalledCount,FailedCount,LastUpdated,@{L="InstalledOrNotApplicablePercentage";e={(($_.NotApplicableCount + $_.InstalledCount) / ($_.NotApplicableCount + $_.InstalledCount + $_.NotInstalledCount+$_.FailedCount))*100}} 


##### Pending reboot on what updates ####
Write-Host ""
Write-Host Generating Devices Pending a reboot for $TargetGroup
Write-Host "####################################################"
Write-Host ""
$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updateScope.IncludedInstallationStates = 'InstalledPendingReboot'
$computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$computerScope.IncludedInstallationStates = 'InstalledPendingReboot'
$pendreboot=($wsus.GetComputerTargetGroups() | Where {
    $_.Name -eq $TargetGroup 
}).GetComputerTargets($computerScope) | ForEach {
        $Computername = $_.fulldomainname
        $_.GetUpdateInstallationInfoPerUpdate($updateScope) | ForEach {
            $update = $_.GetUpdate()
            [pscustomobject]@{
                Computername = $Computername
                TargetGroup = $TargetGroup
                UpdateTitle = $Update.Title 
                IsApproved = $update.IsApproved
            }
    }
} 

##### Need what updates ####
Write-Host ""
Write-Host "Generating Devices needing updates for $TargetGroup"
Write-Host "(Devices with feature updates have been excepted)"
Write-Host "################################################"
Write-Host ""
$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updateScope.IncludedInstallationStates = 'NotInstalled'
$computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$computerScope.IncludedInstallationStates = 'NotInstalled'
$needed=($wsus.GetComputerTargetGroups() | Where {
    $_.Name -eq $TargetGroup 
}).GetComputerTargets($computerScope) | ForEach {
        $Computername = $_.fulldomainname
        $_.GetUpdateInstallationInfoPerUpdate($updateScope) | ForEach {
            $update = $_.GetUpdate()
            [pscustomobject]@{
                Computername = $Computername
                TargetGroup = $TargetGroup
                UpdateTitle = $Update.Title 
                IsApproved = $update.IsApproved
            } 
    } | where-object { $_.UpdateTitle -notlike "*Feature update*" }
    
} 

##### Need what updates ####
Write-Host ""
Write-Host Generating Devices failed updates for $TargetGroup
Write-Host "################################################"
Write-Host ""

$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updateScope.IncludedInstallationStates = 'Failed'
$computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$computerScope.IncludedInstallationStates = 'NotInstalled'
$failed=($wsus.GetComputerTargetGroups() | Where {
    $_.Name -eq $TargetGroup 
}).GetComputerTargets($computerScope) | ForEach {
        $Computername = $_.fulldomainname
        $_.GetUpdateInstallationInfoPerUpdate($updateScope) | ForEach {
            $update = $_.GetUpdate()
            [pscustomobject]@{
                Computername = $Computername
                TargetGroup = $TargetGroup
                UpdateTitle = $Update.Title 
                IsApproved = $update.IsApproved
            } 
    }}


### Export information to Execl

$filename = "$wsusserver $TargetGroup report.xlsx"
$updatesum | Export-excel -Path .\$filename  -AutoSize -TableName 'Device_update_summary' -WorksheetName 'Update Summary'
$pendreboot | Export-excel -Path .\$filename  -AutoSize -TableName 'Devices_needing_reboot' -WorksheetName 'Pending Reboot' -append
$needed | Export-excel -Path .\$filename  -AutoSize -TableName 'Devices_needing_updates' -WorksheetName 'Need updates' -append
$failed | Export-excel -Path .\$filename  -AutoSize -TableName 'Devices_failed_updates' -WorksheetName 'Failed updates' -append
