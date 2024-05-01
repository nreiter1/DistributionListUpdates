$DistributionGroup = Read-Host -Prompt "Distribution Group Name"
Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host

Write-Host -ForegroundColor Yellow "You are about to do a mass removal of contacts to distribution list. Continue? (Y: Yes, N: No): " -NoNewline
$DistributionGroupMemberTrue = Read-Host

if($DistributionGroupMemberTrue -eq "Y")
{
    Import-CSV "LOCATION OF CSV FILE" | foreach {Remove-DistributionGroupMember -Identity $DistributionGroup -Member $_.name -ErrorAction Ignore}
}

#Remove Contacts if no longer in DL

Import-CSV "LOCATION OF CSV FILE" | foreach ({
    $DistributionGroupMember = $_.name #Added to input CSV name
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
})

Write-Host "Distribution Group Name: $DistributionGroup"
Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host #Generates updated distribution list

#Find object in default OU and report to user

$FindDistrubutionGroupDefaultOU = $null
$FindDistrubutionGroupDefaultOU = Get-DistributionGroup -Identity $DistributionGroup -OrganizationalUnit "DOMAIN/OU" -ErrorAction Ignore
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


Clear-Content $CSVEmail | Out-File "LOCATION OF CSV FILE"
Clear-Content $CSVEmail | Out-File "LOCATION OF CSV FILE"
$CSVEmail = "emailaddress"
$CSVEmail | Out-File "LOCATION OF CSV FILE"
$CSVEmail | Out-File "LOCATION OF CSV FILE"

