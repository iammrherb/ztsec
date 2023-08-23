#For adding App Registration in Azure AD and Assigning proper API Permissions for Graph
#Resources - https://docs.microsoft.com/en-us/powershell/module/az.resources/new-azadapplication?view=azps-7.4.0 | https://docs.microsoft.com/en-us/powershell/module/az.resources/add-azadapppermission?view=azps-7.4.0 | https://docs.microsoft.com/en-us/graph/permissions-reference

#Install Necessary Modules
Install-Module AZ -Force
Install-Module AzureAD -Force

#Connect to Azure Powershell
Write-Host -f Yellow "Authenticating to your AzAccount and AzureAD target tenant" 
Connect-AzAccount
Connect-AzureAD

#Variables and naming the new AAD Application
$AppName = Read-Host "Name your Votiro M365 Connector App"
$Today = Get-Date
$ExpirationDate = (Get-Date).AddDays(180)

#Create new App Registration and assign ID (ObjectID) to variable
Write-Host -f Yellow "Creating Azure AD App"
New-AzADApplication -DisplayName $AppName -HomePage "https://portal.azure.com" -ReplyUrls "https://portal.azure.com" | Out-Null
$ObjectID = Get-AzADApplication -DisplayName $AppName | Select-Object -ExpandProperty ID
$AppID = Get-AzADApplication -DisplayName $AppName | Select-Object -ExpandProperty AppID
Start-Sleep -s 5

#Add required permissions for Graph API (reference: https://docs.microsoft.com/en-us/graph/permissions-reference)
Write-Host -f Yellow "Setting $AppName API Permissions"
#Directory.Read.All - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId 7ab1d382-f21e-4acd-a863-ba3e13f7da61 -Type Role
#Group.Read.All - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId 5b567255-7703-4780-807c-7be8301ae99b -Type Role
#GroupMember.Read.All - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId 98830695-27a2-44f7-8c18-0c3ebc9698f6 -Type Role
#Mail.ReadWrite - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId e2a3a72e-5f79-4c64-b1b1-878b674786c9 -Type Role
#Mail.Send - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId b633e1c5-b582-4048-a93e-9f11b44c7e96 -Type Role
#User.Read.All - Application
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId df021288-bdef-4463-88db-98f22de89214 -Type Role
#User.Read - Delegated
Add-AzADAppPermission -ObjectId $ObjectID -ApiId 00000003-0000-0000-c000-000000000000 -PermissionId e1fe6dd8-ba31-4d61-89e7-88639da4683d
Start-Sleep -s 8 

#Create Client Secret for secure remote access to GraphAPI
Write-Host -f Yellow "Creating Votiro App Credentials"
$ClientSecret = Get-AzADApplication -ApplicationID $appid | New-AzADAppCredential -StartDate $today -EndDate $ExpirationDate | Select-Object -ExpandProperty SecretText

#Gather output variable details for the App connection:
$TenantID = Get-AzTenant | Select-Object -ExpandProperty ID
$ClientID = (get-azureadapplication -filter "DisplayName eq '$($AppName)'" | foreach { $_.AppID })

#Display details on screen:
$AppConnectionDetails = "Connection details for the newly created $AppName AAD Application:
============================================================
THE BELOW VALUES WERE AUTO-COPIED TO YOUR CLIPBOARD. RECORD THESE VALUES. 
-------------------
Application name:   $AppName
App/Client ID:      $ClientID
Secret Key:         $ClientSecret
Tenant ID:          $TenantID
-------------------
============================================================"

#Obtain admin consent for new AzureAD app registration
Write-Host -f Yellow "Press ENTER to launch Browser and obtain Consent for Microsoft Graph permissions. Use the same client credentials you previously used"
$AppConnectionDetails | clip
Pause
Start "https://login.microsoftonline.com/common/adminconsent?client_id=$ClientID&redirect_uri=https://portal.azure.com"
Write-Host -f Yellow "After obtaining admin consent, press ENTER again to continue"
Pause

Write-Host -f Yellow $AppConnectionDetails
Write-Host -f Yellow "Script is complete - Press ENTER to safely disconnect from Azure and AzureAD"

Pause

Disconnect-AzAccount
Disconnect-AzureAD
