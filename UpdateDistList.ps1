#Updated 1/15/2024
 
#Sets default variables to 1

$DistributionGroupList = 1
# $Dis Not sure if this variable is still needed.

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
        while($DistributionGroupUpdate -eq 1)
        {
            Write-Host -ForegroundColor Green "What is the name of the user/contact needing to be added to the list?: " -NoNewline
            $DistributionGroupMember = Read-Host
            Write-Host -ForegroundColor Green "The user/contact: $DistributionGroupMember will be added to $DistributionGroup (Y: Yes, N: No): " -NoNewline
            $DistributionGroupMemberTrue = Read-Host
            
            if($DistributionGroupMemberTrue -eq "Y") #Add contact to distribution list
            {
                $DistributionGroupMemberCheck = $null
                $DistributionGroupMemberCheck = Get-MailContact -Identity $DistributionGroupMember
                if($DistributionGroupMemberCheck -eq $null) #Create mail contact if not already in AD
                {
                    Write-Host -ForegroundColor Yellow "Going to add. What is the name of the contact? (Leave blank if you DON'T want to create contact): " -NoNewline
                    $DistributionGroupMemberName = Read-Host
                    New-MailContact -Name $DistributionGroupMemberName -ExternalEmailAddress $DistributionGroupMember
                }
                Add-DistributionGroupMember -Identity $DistributionGroup -Member $DistributionGroupMember
                Read-Host -Prompt "If there are no errors, user added successfully! Press any key to continue....."
            }
            $AddAdditionalUsers = Read-Host -Prompt "Do we need to add any additional users/contacts to this distribution group? (Y: Yes, N: No): "
            if($AddAdditionalUsers -eq "N")
            {$DistributionGroupUpdate = 0}
        }
    }
    
    $DistributionGroupUpdate = 1 #Reset so the 2nd part of this scripts works just the like the 1st part.
     
    #Remove Users
    
    $DistributionGroupMemRemove = Read-Host -Prompt "Do we need to remove a user/contact from the list? (Y: Yes, N: No): "
    if($DistributionGroupMemRemove -eq "Y")
    {
        $Host.ui.rawui.foregroundcolor = "Yellow"
        while($DistributionGroupUpdate -eq 1)
        {
            Write-Host -ForegroundColor Yellow "What is the name of the user/contact needing to be removed from the list?: " -NoNewline
            $DistributionGroupMember = Read-Host
            Write-Host -ForegroundColor Yellow "The user/contact: $DistributionGroupMember will be removed from $DistributionGroup (Y: Yes, N: No): " -NoNewline
            $DistributionGroupMemberTrue = Read-Host
            if($DistributionGroupMemberTrue -eq "Y")
            {
                Remove-DistributionGroupMember -Identity $DistributionGroup -Member $DistributionGroupMember
                Read-Host -Prompt "If there are no errors, user removed successfully! Press any key to continue....."
            }
            $RemoveAdditionalUsers = Read-Host -Prompt "Do we need to remove any additional users/contacts to this distribution group? (Y: Yes, N: No): "
            if($RemoveAdditionalUsers -eq "N")
            {$DistributionGroupUpdate = 0}
        }
    }

    Write-Host "Distribution Group Name: $DistributionGroup"
    Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host #Generates updated distribution list
    $Finished = Read-Host "Do we need to make any additional changes? (Y: Yes, N: No): "
    if($Finished -eq "N")
    {$DistributionGroupList = 0}
}
