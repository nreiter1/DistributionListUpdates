#Updated October 2023

#Sets default variables to 1

$DistributionGroupList = 1
$Dis

#Ask user which distribution group you want to make changes to and echos list.

$DistributionGroup = Read-Host -Prompt "Distribution Group Name"
Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host

#Adjust Distribution Groups

while($DistributionGroupList -eq 1)
{
    #Default Variables
    
    $DistributionGroupUpdate = 1
    
    #Add Users
    
    $DistributionGroupMemAdd = Read-Host -Prompt "Do we need to add a user/contact to this distribution list? (Y: Yes, N: No)"
    if($DistributionGroupMemAdd -eq "Y")
    {
        $Host.ui.rawui.foregroundcolor = "Green"
        while($DistributionGroupUpdate -eq 1)
        {
            $DistributionGroupMember = Read-Host - Prompt "What is the name of the user/contact needing to be added to the list?"
            $DistributionGroupMemberTrue = Read-Host - Prompt "The user/contact: $DistributionGroupMember will be added to $DistributionGroup (Y: Yes, N: No)"
            if($DistributionGroupMemberTrue -eq "Y")
            {
                Add-DistributionGroupMember -Identity $DistributionGroup -Member $DistributionGroupMember
                Read-Host -Prompt "If there are no errors, user added successfully!"
            }
            $AddAdditionalUsers = Read-Host -Prompt "Do we need to add any additional users/contacts to this distribution group? (Y: Yes, N: No)"
            if($AddAdditionalUsers -eq "N")
            {$DistributionGroupUpdate = 0}
        }
    }
    
    $Host.ui.rawui.foregroundcolor = "White"
    $DistributionGroupUpdate = 1 #Reset so the 2nd part of this scripts works just the like the 1st part.
     
    #Remove Users
    
    $DistributionGroupMemRemove = Read-Host -Prompt "Do we need to remove a user/contact from the list? (Y: Yes, N: No)"
    if($DistributionGroupMemRemove -eq "Y")
    {
        $Host.ui.rawui.foregroundcolor = "Yellow"
        while($DistributionGroupUpdate -eq 1)
        {
            $DistributionGroupMember = Read-Host - Prompt "What is the name of the user/contact needing to be removed from the list?"
            $DistributionGroupMemberTrue = Read-Host - Prompt "The user/contact: $DistributionGroupMember will be removed from $DistributionGroup (Y: Yes, N: No)"
            if($DistributionGroupMemberTrue -eq "Y")
            {
                Remove-DistributionGroupMember -Identity $DistributionGroup -Member $DistributionGroupMember
                Read-Host -Prompt "If there are no errors, user removed successfully!"
            }
            $RemoveAdditionalUsers = Read-Host -Prompt "Do we need to remove any additional users/contacts to this distribution group? (Y: Yes, N: No)"
            if($RemoveAdditionalUsers -eq "N")
            {$DistributionGroupUpdate = 0}
        }
    }

    $Host.ui.rawui.foregroundcolor = "White"
    Write-Host "Distribution Group Name: $DistributionGroup"
    Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host #Generates updated distribution list
    $Finished = Read-Host "Do we need to make any additional changes? (Y: Yes, N: No)"
    if($Finished -eq "N")
    {$DistributionGroupList = 0}
}
   $Host.ui.rawui.foregroundcolor = "white"