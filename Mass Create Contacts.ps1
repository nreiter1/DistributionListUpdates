Write-Host -ForegroundColor Green "You are about to do a mass add of contacts to distribution list. Continue? (Y: Yes, N: No): " -NoNewline
$DistributionGroupMemberTrue = Read-Host

if($DistributionGroupMemberTrue -eq "Y")
{
    Import-CSV "LOCATION OF CSV FILE" | foreach {New-MailContact -ExternalEmailAddress $_.emailaddress -Name $_.fullname -ErrorAction Ignore} #Create each contact if contact doesn't exist.
    Clear-Content "LOCATION OF CSV FILE"
    $CSVEmail = "emailaddress"
    $CSVEmail | Out-File "LOCATION OF CSV FILE"
}

