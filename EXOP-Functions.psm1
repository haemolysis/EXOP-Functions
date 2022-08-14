<#
 The Brains Trust Dev Department presents.. 
#>
Write-Host @"

`$`$`$`$`$`$`$`$\ `$`$\   `$`$\  `$`$`$`$`$`$\        `$`$\   `$`$\           `$`$\                               
`$`$  _____|`$`$ |  `$`$ |`$`$  __`$`$\       `$`$ |  `$`$ |          `$`$ |                              
`$`$ |      \`$`$\ `$`$  |`$`$ /  `$`$ |      `$`$ |  `$`$ | `$`$`$`$`$`$\  `$`$ | `$`$`$`$`$`$\   `$`$`$`$`$`$\   `$`$`$`$`$`$\  
`$`$`$`$`$\     \`$`$`$`$  / `$`$ |  `$`$ |      `$`$`$`$`$`$`$`$ |`$`$  __`$`$\ `$`$ |`$`$  __`$`$\ `$`$  __`$`$\ `$`$  __`$`$\ 
`$`$  __|    `$`$  `$`$<  `$`$ |  `$`$ |      `$`$  __`$`$ |`$`$`$`$`$`$`$`$ |`$`$ |`$`$ /  `$`$ |`$`$`$`$`$`$`$`$ |`$`$ |  \__|
`$`$ |      `$`$  /\`$`$\ `$`$ |  `$`$ |      `$`$ |  `$`$ |`$`$   ____|`$`$ |`$`$ |  `$`$ |`$`$   ____|`$`$ |      
`$`$`$`$`$`$`$`$\ `$`$ /  `$`$ | `$`$`$`$`$`$  |      `$`$ |  `$`$ |\`$`$`$`$`$`$`$\ `$`$ |`$`$`$`$`$`$`$  |\`$`$`$`$`$`$`$\ `$`$ |      
\________|\__|  \__| \______/       \__|  \__| \_______|\__|`$`$  ____/  \_______|\__|      
                                                            `$`$ |                          
                                                            `$`$ | ! MADE BY MiCHAEL ~                       
                                                            \__|                                                                                                  
                                 
"@

#cheeky hint to upgrade to PS7
if ($PSVersionTable.PSVersion -like "6.*" -or $PSVersionTable.PSVersion -like "5.*") {
    Write-Host "You are using PowerShell $($PSVersionTable.PSVersion). Consider upgrading to PowerShell 7"
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

Write-Host "Lost? Need help? Find commands with " -NoNewline
Write-Host "Get-Command -Module EXOP-Functions " -NoNewline -ForegroundColor Red
Write-Host "or " -NoNewline
Write-Host "Show-EXOHelp" -ForegroundColor Red

Write-Host "Need to find the accounts.txt file? Use " -NoNewLine
Write-Host "Find-AccountsPlease" -ForegroundColor Red

#set  accounts variable to accounts.txt - probs shouldn't set global variables but w/e, used script
$Script:accounts = "$env:USERPROFILE\Documents\temp\accounts.txt"
$Script:tempdir = "$env:USERPROFILE\Documents\temp"

#show commands in this module

function Show-EXOHelp {
    #get-command -Module EXOP-Functions | Where {$_.CommandType -ne "Alias"} | Select Name
    Write-Host @"
List of commands: 
Show-EXOHelp                            : Shows this help
Show-DynamicDistributionGroupMembers    : Show members of a dynamic DL
Run-PreflightChecks                     : Check validity of data set for list of users
"@
    
}

#open the accounts.txt file as well as link it in the terminal window
function Find-AccountsPlease {
    Write-Host "Your accounts.txt file is located at $accounts"
    Start-Process -FilePath $accounts
}
#list memebers of a dynamic DL - can't see this in Exchange Online for some reason
function Show-DynamicDistributionGroupMembers {
    foreach ($account in $accounts) {
        $check = Get-DynamicDistributionGroup -Identity $account
        Get-Recipient -RecipientPreviewFilter ($check.RecipientFilter) -ResultSize Unlimited | Export-CSV -Path $($Script:tempdir)\$check.csv -NoTypeInformation
    }
}

Set-Alias -Name Show-DDGM -Value Show-DynamicDistributionGroupMembers #create alias

#check validity of user data sets, probably a better way of doing this but w/e, looks nice
function Show-PreflightChecks {
    $bad = @()
    $i = 0

    foreach ($user in $accounts) {
        Write-Host "Checking $user"
        try {
            Get-EXOMailbox -Identity $user | Out-Null
        }
        catch {
            Write-Host "User $user does not exist, please check the entry" -ForegroundColor Red
            $bad += $user
        }
        finally {
            $i = $i + 1
            Write-Progress -Activity "checking users" -Status "Progress: " -PercentComplete ($i / $users.count * 100)
        }

        if ($bad) {
            Write-Host -ForegroundColor Red @"
            +++++++++++++++++++++++++
            +                       +                       
            +       BAD  USERS      +
            +        DETECTED       +
            +                       +
            +++++++++++++++++++++++++
"@
            foreach ($u in $bad) {
                Write-Host -ForegroundColor Red "$u"
            }
        }
        else {
            Write-Host -ForegroundColor Green @"
            +++++++++++++++++++++++++
            +                       +                       
            +      NO BAD USERS     +
            +                       +
            +++++++++++++++++++++++++
"@
        }
    }
}

function Add-BulkAccessforUser { 
    param (
        [Parameter(Mandatory, HelpMessage = "Who is the target user?")]
        [string]$Username,
        
        #could also make this non-mandatory and default to "both"
        [Parameter(Mandatory, HelpMessage = "What is the operation? full = add full access only, sendas = add send as only, both = add both permissions")]
        [ValidateSet("full", 'sendas', 'both')]
        [string]$AccessType
    )

    if ($AccessType -eq "full") {
        Write-Host "Adding full access for $Username"
        foreach ($target in $accounts) {
            Add-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false
            Write-Host "$Username added to $target"
        }
    }
    elseif ($AccessType -eq "sendas") {
        Write-Host "Adding send as access for $Username"
        foreach ($target in $accounts) {
            Add-RecipientPermission -Identity $target -Trustee $Username -AccessRights SendAs -Confirm:$false
            Write-Host "$Username added to $target"
        }
    }
    elseif ($AccessType -eq "both") {
        Write-Host "Adding send as and full access for $Username"
        foreach ($target in $accounts) {
            Add-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false
            Add-RecipientPermission -Identity $target -Trustee $Username -AccessRights SendAs -Confirm:$false
            Write-Host "$Username added to $target"
        }
    }
    else {
        Write-Host "error"
    }
    
}

function Remove-BulkAccessforUser { 
    <#
    .SYNOPSIS
        Removes either full access, send as access, or both for a user on a list of mailboxes in accounts.txt
    .DESCRIPTION
        Removes either full access, send as access, or both for a user on a list of mailboxes in accounts.txt
    .EXAMPLE
        Remove-BulkAccessforUser 
        Will prompt for a user's UPN and then prompt for either full, sendas, or both

    .EXAMPLE
        Remove-BulkAccessforUser -Username example@fabrikam.com -AccessType both
        Removes full and send as access for the user example@fabrikam.com to the mailboxes in accounts.txt
    .EXAMPLE
        Remove-BulkAccessforUser -Username example@fabrikam.com -AccessType sendas
        remove only send-as access
    #>
    
    param (
        [Parameter(Mandatory, HelpMessage = "Who is the affected user?")]
        [string]$Username,

        #could also make this non-mandatory and default to "both", might go with this as it's the most common use case
        [Parameter(Mandatory, HelpMessage = "What is the operation? full = remove full access only, sendas = remove send as only, both = remove both permissions")]
        [ValidateSet("full", 'sendas', 'both')]
        [string]$AccessType
    )

    if ($AccessType -eq "full") {
        foreach ($target in $accounts) {
            Write-Host "Removing Full access for $Username to $target"
            Remove-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -Confirm:$false
            Add-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$false
            Remove-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -Confirm:$false
        }
    }
    elseif ($AccessType -eq "sendas") {
        foreach ($target in $accounts) {
            Write-Host "Removing send-as access for $Username to $target"
            Remove-RecipientPermission -Identity $target -Trustee $Username -AccessRights SendAs -Confirm:$false
        }
    }
    elseif ($AccessType -eq "both") {
        foreach ($target in $accounts) {
            Write-Host "Removing full access and send-as permissions for $Username to $target"
            Remove-RecipientPermission -Identity $target -Trustee $Username -AccessRights SendAs -Confirm:$false
            Remove-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -Confirm:$false
            Add-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$false
            Remove-MailboxPermission -Identity $target -User $Username -AccessRights FullAccess -Confirm:$false
        }
    }
    else {
        Write-Error "Not sure what happened but this is not good"
    }
    
}

Export-ModuleMember -Function * -Alias *