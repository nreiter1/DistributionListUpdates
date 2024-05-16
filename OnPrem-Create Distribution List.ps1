#Updated 3/28/2024

Write-Host "What is the name of the distribution group that we want to create? (Without @tnsk.com):" -NoNewline
$DistributionGroupName = Read-Host


Write-Host "The distribution group $DistributionGroupName will be created (Y: Yes, N: No):" -NoNewline
$DistributionGroupCreate = Read-Host
if($DistributionGroupCreate -eq "Y")
{
    try
    {
        New-DistributionGroup -Name $DistributionGroupName -DisplayName $DistributionGroupName -OrganizationalUnit "tps.local/Exchange Contacts/Distribution Lists" -ErrorAction Stop
        #Automatically sets contacts/users can only be added by us.

        Write-Host "Created! Do external senders need to be able to send to this group? (Y: Yes, N: No):"
        $AllowExternalSenders = Read-Host

        if($AllowExternalSenders -eq "Y")
        {
            Get-DistributionGroup -identity $DistributionGroupName | Set-DistributionGroup -RequireSenderAuthenticationEnabled $False #Allows External Senders
        }
    }
    catch{Write-Host -ForegroundColor Red "***The distribution group already exists!***"}
}
