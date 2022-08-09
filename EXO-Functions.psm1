<#
 The Brains Trust Dev Department presents.. 
#>

#cheeky hint to upgrade to PS7
if ($PSVersionTable.PSVersion -like "6.*" -or $PSVersionTable.PSVersion -like "5.*") {
    Write-Host "You are using PowerShell $($PSVersionTable.PSVersion). Time to upgrade to PowerShell 7"
    Write-Host ""
    Write-Host "*************************************************************"
    Write-Host "*                                                           *"
    Write-Host "* Update here: https://aka.ms/powershell-release?tag=stable *"
    Write-Host "*                                                           *"
    Write-Host "*************************************************************"
}

#Check for presence of EXOV2 module - after login it loads as tmp-[random chars]. We can check with Get-Module 
if (!(Get-Module tmp* | Where-Object Description -Like 'Implicit remoting for*')) {
    Write-Error "Exchange module V2 may not be loaded - please ensure you have connected to Exchange Online"
}

<#
Check if the temp folder exists - may need to change this location after I am finished
Create if doesn't exist
#>

if (!(Test-Path -Path $env:USERPROFILE\Documents\temp)) {
    New-Item -ItemType "directory" -Path "$env:USERPROFILE\Documents\temp" -Force | Out-Null
    Write-Warning "A temp folder is required to store output. One has been created for you here: $($env:USERPROFILE)\Documents\temp"
}

#check that email accounts.txt file exists in temp, create if doesn't exist
if (!(Test-Path $env:USERPROFILE\Documents\temp\accounts.txt)) {
    New-Item -ItemType File -Path $env:USERPROFILE\Documents\temp\accounts.txt | Out-Null
    Write-Warning "An accounts.txt file is required in the location $($env:USERPROFILE)\Documents\temp\accounts.txt"
    Write-Warning "A blank accounts.txt file has been created in the above location for you"
}

#set  accounts variable to accounts.txt - probs shouldn't set global variables but w/e, used script
$Script:accounts = "$env:USERPROFILE\Documents\temp\accounts.txt"
$Script:tempdir = "$env:USERPROFILE\Documents\temp"

#list memebers of a dynamic DL - can't see this in Exchange Online for some reason
function Show-DynamicDistributionGroupMembers {
    foreach ($account in $accounts) {
        $check = Get-DynamicDistributionGroup -Identity $account
        Get-Recipient -RecipientPreviewFilter ($check.RecipientFilter) -ResultSize Unlimited | Export-CSV -Path $($Script:tempdir)\$check.csv -NoTypeInformation
    }
}

Set-Alias -Name Show-DDGM -Value Show-DynamicDistributionGroupMembers #create alias


Export-ModuleMember -Function * -Alias *