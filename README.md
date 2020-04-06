# Set-TeamsAppSetupPolicy
This script automates applying Teams Application Setup Policy on a select groups of users ideally using a Dymantic Azure AD Group.  

#### DESCRIPTION
This has been created to run on a schedule in Azure Automation and apply the Teams policy during a rollout phase where users will have a limited view in Teams.    

```powershell
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

```
