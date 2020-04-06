<# 
.SYNOPSIS
This script automates applying Teams Application Setup Policy on a select groups of users ideally using a Dymantic Azure AD Group.  

.DESCRIPTION
This has been created to run on a schedule in Azure Automation and apply the Teams policy during a rollout phase where users will have a limited view in Teams.    

## Set-TeamsAppSetupPolicy.ps1 [-SourceGroups <Array[ObjectID]>] [-ExcludeGroups <Array[ObjectID]>] [-TeamsAppPolicyName <String[ABP Name]>]

.PARAMETER SourceGroups
The SourceGroup parameter details the ObjectId of the Azure Group which contains all the desired users that need the Address Book Policy.

.PARAMETER ExcludeGroups
The ExcludeGroups switch will force the default global policy for these users and they will be excluded from the default policy. 

.PARAMETER TeamsAppPolicyName
The AddressBookPolicy parameter specifies the name of the Address Book Policy which should be applied. 

.PARAMETER AutomationPSCredential
The AutomationPSCredential parameter defines which Azure Automation Cred you would like to use. 

.EXAMPLE
Set-TeamsAppSetupPolicy.ps1 -SourceGroups '7b7c4926-c6d7-4ca8-9bbf-5965751022c2' -ExcludeGroups '302a6fb0-9f6c-472e-a41d-7771b3285209' -TeamsAppPolicyName 'Executive Teams Policy'

-- SET MEMBERS FOR TEAMS POLICY --

In this example the script will apply the 'Executive Teams Policy' to Teams users who are members of Group '7b7c4926-c6d7-4ca8-9bbf-5965751022c2' but are excluded if they are a member of 302a6fb0-9f6c-472e-a41d-7771b3285209.

.LINK

 Customize Microsoft Teams to highlight the apps that are most important - https://docs.microsoft.com/en-us/powershell/module/skype/grant-csteamsappsetuppolicy?view=skype-ps

.NOTES
This function requires that you have connected the Azure AD & S4B modules and you must created groups have at least 1 member in both groups othwise the Compare-Object will fail. Sorry this is a bit of rush job :(

Please note, when using Azure Automation with more than one user group the array should be set to JSON for example ['ObjectID','ObjectID']

[AUTHOR]
Joshua Bines, Consultant

Find me on:
* Web:     https://theinformationstore.com.au
* LinkedIn:  https://www.linkedin.com/in/joshua-bines-4451534
* Github:    https://github.com/jbines
  
[VERSION HISTORY / UPDATES]
0.0.1 20200406 - JBINES - Created the bare bones

[TO DO LIST / PRIORITY]

#>

Param 
(
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [array]$SourceGroups,
    [Parameter(Mandatory = $False)]
    [ValidateNotNullOrEmpty()]
    [array]$ExcludeGroups,
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string]$TeamsAppPolicyName,
    [Parameter(Mandatory = $False)]
    [ValidateNotNullOrEmpty()]
    [String]$AutomationPSCredential
)

# Set VAR
    $counter = 0

# Success Strings
    $sString0 = "CMDlet:Grant-CsTeamsAppSetupPolicy"

# Info Strings
    $iString0 = "Set Teams App Policy"

# Warn Strings
    $wString0 = "CMDlet:Measure-Object;No Members"

# Error Strings

    $eString1 = "Hey! You hit the -DifferentialScope limit. Let's break out of this loop"
    $eString2 = "Hey! Looks like we are having issues finding your users in the source group. Make sure it's not empty!"

    function Write-Log([string[]]$Message, [string]$LogFile = $Script:LogFile, [switch]$ConsoleOutput, [ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")][string]$LogLevel)
    {
           $Message = $Message + $Input
           If (!$LogLevel) { $LogLevel = "INFO" }
           switch ($LogLevel)
           {
                  SUCCESS { $Color = "Green" }
                  INFO { $Color = "White" }
                  WARN { $Color = "Yellow" }
                  ERROR { $Color = "Red" }
                  DEBUG { $Color = "Gray" }
           }
           if ($Message -ne $null -and $Message.Length -gt 0)
           {
                  $TimeStamp = [System.DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                  if ($LogFile -ne $null -and $LogFile -ne [System.String]::Empty)
                  {
                         Out-File -Append -FilePath $LogFile -InputObject "[$TimeStamp] [$LogLevel] $Message"
                  }
                  if ($ConsoleOutput -eq $true)
                  {
                         Write-Host "[$TimeStamp] [$LogLevel] :: $Message" -ForegroundColor $Color

                    if($AutomationPSCredential)
                    {
                         Write-Output "[$TimeStamp] [$LogLevel] :: $Message"
                    } 
                  }
           }
    }

    #Validate Input Values From Parameter 

    Try{

        if ($AutomationPSCredential) {
            
            $Credential = Get-AutomationPSCredential -Name $AutomationPSCredential

            Connect-AzureAD -Credential $Credential
            
            #$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -Name $ConnectionName 
            #Import-Module (Import-PSSession -Session $ExchangeOnlineSession -AllowClobber -DisableNameChecking) -Global

            $SfBsession = New-CsOnlineSession -Credential $Credential
            Import-PSSession $SfBsession -AllowClobber | Out-Null
            
            }
                            
        $objSourceGroupMembers = @($SourceGroups | ForEach-Object {Get-AzureADGroupMember -All:$true -ObjectId $_})

        #Return Only Unique values remove any duplicates
        $SourceGroupMembers = $objSourceGroupMembers | Select-Object -Unique
        Write-Log -Message "$iString0 - Source Group Member Count: $($SourceGroupMembers.Count)" -LogLevel INFO -ConsoleOutput
        
        if(($ExcludeGroups | Measure-Object).count -gt 0){

            #Compile the exclude user list. 
            $objExcludeGroupMembers = @($ExcludeGroups | ForEach-Object {Get-AzureADGroupMember -All:$true -ObjectId $_})
            
            #Return Only Unique values remove any duplicates
            $ExcludeGroupMembers = $objExcludeGroupMembers | Select-Object -Unique
            
            Write-Log -Message "$iString0 - Exclude Group Member Count: $($ExcludeGroupMembers.Count)" -LogLevel INFO -ConsoleOutput

        }

    }
    
    Catch{
    
        $ErrorMessage = $_.Exception.Message
        Write-Error $ErrorMessage

            If($?){Write-Log -Message $ErrorMessage -LogLevel Error -ConsoleOutput}

        Break

    }
    
    if (($ExcludeGroupMembers | Measure-Object).count -gt 0) {
        ForEach($user in $ExcludeGroupMembers){

            Grant-CsTeamsAppSetupPolicy -Identity $user.UserPrincipalName -PolicyName:$Null
            if($?){Write-Log -Message "$sString0;PolicyName:NULL;UserObjectId:$($user.ObjectId);UserUPN:$($user.UserPrincipalName)" -LogLevel SUCCESS -ConsoleOutput}
        }
        
    }
        
    if ($SourceGroupMembers) {
        
        #Null var
        $user = $null

        If($ExcludeGroupMembers){

            #Strip ExcludeGroupMembers Array Members from the group
            $ScopedGroupMembers = Compare-Object -ReferenceObject $SourceGroupMembers.UserPrincipalName -DifferenceObject $ExcludeGroupMembers.UserPrincipalName

        }
        Else{

            $ScopedGroupMembers = $SourceGroupMembers
        }
        
        #Foreach Source group members
        foreach ($user in $ScopedGroupMembers) {
            If($user.InputObject){

                Grant-CsTeamsAppSetupPolicy -identity $User.InputObject -PolicyName $TeamsAppPolicyName
                if($?){Write-Log -Message "$sString0;PolicyName:$TeamsAppPolicyName;UserUPN:$($user.InputObject)" -LogLevel SUCCESS -ConsoleOutput}
            }
            If($user.UserPrincipalName){
                Grant-CsTeamsAppSetupPolicy -identity $User.UserPrincipalName -PolicyName $TeamsAppPolicyName
                if($?){Write-Log -Message "$sString0;PolicyName:$TeamsAppPolicyName;UserUPN:$($user.UserPrincipalName)" -LogLevel SUCCESS -ConsoleOutput}
            }
        }
    }
    else {
        Write-Log -Message $eString2 -LogLevel Error -ConsoleOutput
    }

if ($AutomationPSCredential) {
    
    #Invoke-Command -Session $ExchangeOnlineSession -ScriptBlock {Remove-PSSession -Session $ExchangeOnlineSession}

    Disconnect-AzureAD

    #Disconnect-PSSession $sfbSession
    Remove-PSSession $sfbSession
}
