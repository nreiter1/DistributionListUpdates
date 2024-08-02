#Updated 2/7/2024

#Sets default variables to 1

$DistributionGroupList = 1
# $Dis Not sure if this variable is still needed.

#Nulls out PowerShell variables to in effort to avoid adding contacts that don't belong on particular distribution lists.

$DistributionGroup = $null
$DistributionGroupMember = $null

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
                
                try{$DistributionGroupMemberCheck = Get-MailContact -Identity $DistributionGroupMember -ErrorAction Stop} #Check for existing mail contact in AD
                catch
                {
                    try{Get-DistributionGroup - Identity $DistributionGroupMember -ErrorAction Stop}
                    catch
                    {
                        try{Get-ADGroup -Identity $DistributionGroupMember.TrimEnd().Split('@')[0] -ErrorAction Stop} #Check if object is a security group, NOT distribution group
                        catch
                        {
                            $ADUserEmailTrimmed = $DistributionGroupMember.TrimEnd().Split('@')[0] #Removes everything to right of '@' to check if object is AD user
                            #Write-Host $ADUserEmailTrimmed
                            try{Get-ADUser -Identity $ADUserEmailTrimmed -ErrorAction Stop} #Check if object is AD user (first initial + last name)
                            catch
                            {
                                #Write-Output "User not found. Trying something."
                                $ADUserFirstInitialTrimmed = $ADUserEmailTrimmed.TrimStart().Substring(1) #Trims first initial off the beginning
                                $ADUserFirstInitial = $ADUserEmailTrimmed.Substring(0,1) #Gets first initial from $ADUserEmailTrimmed and sets $ADUserFirstInitial
                                $ADUserReversed = $ADUserFirstInitialTrimmed + $ADUserFirstInitial #Last name only + first initial
                                #Write-Output $ADUserReversed #For testing
                                try{Get-ADUser -Identity $ADUserReversed -ErrorAction Stop} #Check if object is AD user(last name + first initial)
                                catch
                                {
                                    if($DistributionGroupMemberCheck -eq $null) #Create mail contact if not in AD and all other try statements failed
                                    {
                                        $DistributionGroupMemberName = $null
                                        Write-Host -ForegroundColor Yellow "Contact NOT found! Going to add. What is the name of the contact? (Leave blank if you DON'T want to create contact): " -NoNewline
                                        $DistributionGroupMemberName = Read-Host
                                        New-MailContact -Name $DistributionGroupMemberName -ExternalEmailAddress $DistributionGroupMember
                                        if($DistributionGroupMemberName -ne $null){Write-Host -ForegroundColor Yellow "***Contact created! Remember to move contact into Exchange Contacts so it syncs correctly!***"}      
                                    }
                                }
                            }
                        } #Works
                    }
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
                
                #Check if contact is a member of any distribution lists in preparation of passing to if statement below.
                
                Write-Host "Checking if contact is still a member of any distribution lists."
                
                try
                {
                    Write-Host $DistributionGroupMember
                    $mct=Get-MailContact -Identity $DistributionGroupMember -ErrorAction Stop
                    $dn=$mct.distinguishedname
                    $Filter="Members -like ""$dn"""
                    $MailContactDLCheck = $null
                    $MailContactDLCheck = Get-DistributionGroup -ResultSize unlimited -Filter $filter | select Displayname, primarysmtpaddress | sort DisplayName -ErrorAction Ignore

                    if($MailContactDLCheck -eq $null) #Delete contact if no longer a member of any distribution lists.
                    {
                        Write-Host -ForegroundColor Yellow "***Contact is no longer a member of any distribution lists. Select Y if in prompt below if you wish to delete contact***"
                        Remove-MailContact -Identity $DistributionGroupMember #This will prompt if you want to delete contact or not.
                    }
                    else{Write-Host -ForegroundColor Yellow "Contact is still a member of at least one distribution list! Leaving.....`n"}
                }
                catch{Write-Host -ForegroundColor Yellow "Object is not a contact. Ignoring.....`n"}
                $DistributionGroupMember = $null #Null out variable to avoid accidentaly readding contact to another DL
            }
            $RemoveAdditionalUsers = Read-Host -Prompt "`nDo we need to remove any additional users/contacts to this distribution group? (Y: Yes, N: No): "
            if($RemoveAdditionalUsers -eq "N")
            {$DistributionGroupUpdate = 0}
        }
    }
    
    Write-Host "Distribution Group Name: $DistributionGroup"
    Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host #Generates updated distribution list

    #Find object in default OU and report to user

    $FindDistrubutionGroupDefaultOU = $null
    $FindDistrubutionGroupDefaultOU = Get-DistributionGroup -Identity $DistributionGroup -OrganizationalUnit "DOMAIN/Users" -ErrorAction Ignore
    if($FindDistrubutionGroupDefaultOU -ne $null)
    {
        Write-Host -ForegroundColor Yellow "***OU is in DEFAULT OU! Please move to sync OU IF no @tenaska.com emails are present!***"
    }

    #Find object in non-sync OU and report to user

    $FindDistrubutionGroupExchOUNonSync = $null
    $FindDistributionGroupExchOUNonSync = Get-DistributionGroup -Identity $DistributionGroup -OrganizationalUnit "DOMAIN/OU" -ErrorAction Ignore
    $FindDistributionGroupExchOUNonSync = Get-DistributionGroup -Identity $DistributionGroup -OrganizationalUnit "DOMAIN/OU" -ErrorAction Ignore
    if($FindDistributionGroupExchOUNonSync -ne $null)
    {
        Write-Host -ForegroundColor Yellow "***OU is in Exchange NON-SYNC OU! You will need to activate Exchange PIM and make your changes in Exchange Admin Center as well!***"
    }

    $Finished = Read-Host "Do we need to make any additional changes? (Y: Yes, N: No): "
    if($Finished -eq "N")
    {$DistributionGroupList = 0}
}
