$DistributionGroup = Read-Host -Prompt "Distribution Group Name"
Get-DistributionGroupMember -Identity $DistributionGroup | Select -ExpandProperty primarysmtpaddress | Write-Host

Write-Host -ForegroundColor Green "You are about to do a mass add of contacts to distribution list. Continue? (Y: Yes, N: No): " -NoNewline
$DistributionGroupMemberTrue = Read-Host

if($DistributionGroupMemberTrue -eq "Y")
{
    Import-CSV "C:\temp\Non-TNSKEmails.csv" | foreach {New-MailContact -ExternalEmailAddress $_.name -Name $_.name -ErrorAction Ignore} #Create each contact if contact doesn't exist.
    Import-CSV "C:\temp\Non-TNSKEmails.csv" | foreach {Add-DistributionGroupMember -Identity $DistributionGroup -Member $_.name} #Add External Emails
    Import-CSV "C:\temp\TNSKEmails.csv" | foreach {Add-DistributionGroupMember -Identity $DistributionGroup -Member $_.name} #Add Internal Emails
}

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
