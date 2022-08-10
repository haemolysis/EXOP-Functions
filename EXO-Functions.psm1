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
                                                            `$`$ |  MADE BY MiCHAEL                        
                                                            \__|                                                                                                   
                                 
"@

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

Write-Host "Lost? Need help? Find commands with Get-Command -Module EXOP-Functions or Show-EXOHelp"

#set  accounts variable to accounts.txt - probs shouldn't set global variables but w/e, used script
$Script:accounts = "$env:USERPROFILE\Documents\temp\accounts.txt"
$Script:tempdir = "$env:USERPROFILE\Documents\temp"

#show commands in this module

function Show-EXOHelp {
    Write-Host @"
    List of commands: 
    Show-EXOHelp                            : Shows this help
    Show-DynamicDistributionGroupMembers    : Show members of a dynamic DL
    Run-PreflightChecks                     : Check validity of data set for list of users
"@
    
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
            +      BAD  USERS       +
            +       DETECTED        +
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
+     NO BAD USERS      +
+                       +
+++++++++++++++++++++++++
"@
        }
    }
}

Export-ModuleMember -Function * -Alias *