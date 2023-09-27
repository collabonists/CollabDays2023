###
# Tatjana Davis
# web: www.collabonists.io
# mail: td@collabonists.io
# Created: 15.09.2023
# Modified: 15.09.2023
# Description: Gather data for license monitoring and store in SharePoint Online
###

Write-Host "-------------------------------------------------" -ForegroundColor Green
Write-Host "Loading config settings xml..." -ForegroundColor Green
# Path to config File
$configFile = "config\010_config.xml"

# Test if config file  does exist
if((Test-Path $configFile) -eq $false) 
{ 
   Write-host "Config XML not found" 
   #exit 
} 

# Load config file
[XML]$config = Get-Content $configFile

# Load XML values
$log = $config.Config.log
$siteURL = $config.Config.siteCollectionURL
$tenantID = $config.Config.tenantID
$spoListLicenseGroups = $config.Config.SPOListLicenseGroups
$spoListCompanyCodes = $config.Config.SPOListCompanyCodes
$spoListLicenseTypes = $config.Config.SPOListLicenseTypes
$spoListLicensesPerType = $config.Config.SPOListLicensesPerType
$spoListLicensesPerGroup = $config.Config.SPOListLicensesPerGroup
$spoListTenant = $config.Config.SPOListTenant

# Log file
$date = Get-Date -Format FileDateTime
$dateForSPO = Get-Date -Format dd.MM.yyyy
$logFile = $log + "010_GetLicenseData" + $date + ".txt"
Start-Transcript -Path $logFile

# Import graph module and connect - needs to be done before other connections in order to work
Write-Host "Connect to MS Graph..."
#Import-Module Microsoft.Graph
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All","Organization.Read.All" -TenantId $tenantID -NoWelcome

# Import PnP module and connect to PnP Online
Write-Host "Connect to PnP Online..."
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $siteURL -Interactive 

Read-Host "Press Enter to continue..."

Write-Host "-------------------------------------------------" -ForegroundColor Green

# Read information from SPO
Write-Host "Read information from SPO..."
$companyCodes = Get-PnPListItem -List $spoListCompanyCodes 
$licenseGroups = Get-PnPListItem -List $spoListLicenseGroups 
$licenseTypes = Get-PnPListItem -List $spoListLicenseTypes

##########################
# Determine licenses per group
##########################
Write-Host "Determine licenses per group data..."

# Create custom license group object
class LicenseGroupObject
{
   [string]$companyCode
   [int]$companyCodeID
   [string]$licenseGroup
   [int]$licenseGroupID
   [string]$correspondingLicenseType
   [int]$correspondingLicenseTypeID
   [string]$groupName
   [int]$groupMembers;

   LicenseGroupObject([string]$companyCode, [int]$companyCodeID, [string]$licenseGroup, [int]$licenseGroupID, [string]$correspondingLicenseType, [int]$correspondingLicenseTypeID, [string]$groupName, [int]$groupMembers)
   {
      $this.companyCode = $companyCode
      $this.companyCodeID = $companyCodeID
      $this.licenseGroup = $licenseGroup
      $this.licenseGroupID = $licenseGroupID
      $this.correspondingLicenseType = $correspondingLicenseType
      $this.correspondingLicenseTypeID = $correspondingLicenseTypeID
      $this.groupName = $groupName
      $this.groupMembers = $groupMembers
   }
}

# Create array of group objects
[System.Collections.ArrayList]$groups = @()

foreach($companyCode in $companyCodes)
{
   foreach($licenseGroup in $licenseGroups)
   {
      $step = $groups.Add([LicenseGroupObject]:: new($companyCode.FieldValues.CompanyCode, $companyCode.FieldValues.ID, $licenseGroup.FieldValues.Title, $licenseGroup.FieldValues.ID, $licenseGroup.FieldValues.LicenseType.LookupValue, $licenseGroup.FieldValues.LicenseType.LookupId, $companyCode.FieldValues.CompanyCode + "_" + $licenseGroup.FieldValues.Title, 0))
   }
}

# Get group member counts from MS Graph
foreach($group in $groups)
{
   try
   {
      $groupID = Get-MgGroup -Filter "displayName eq '$($group.groupName)'" | Select-Object -ExpandProperty id
      $groupMembers = Get-MgGroupMember -GroupId $groupID
      $group.groupMembers = $groupMembers.Count
   }
   catch
   {
      # Write-Host "Group " $group.groupName " not found" -ForegroundColor Red
   }
}

# Write license group data to SharePoint Online
$countSPOadded = 0
foreach($group in $groups)
{
   try
   {
      $step = Add-PnPListItem -List $spoListLicensesPerGroup -Values @{"Date" = $dateForSPO; "CompanyCode" = $group.companyCodeID; "LicenseGroup" = $group.licenseGroupID; "NumberUsers" = $group.groupMembers} 
      $countSPOadded++
   }
   catch
   {
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

Write-Host $countSPOadded "group data entries added to SPO" -ForegroundColor Green

##########################
# Determine licenses per type
##########################
Write-Host "Determine licenses per type data..."

# Create custom license type object
class LicenseTypeObject
{
   [string]$companyCode
   [int]$companyCodeID
   [string]$licenseType
   [int]$licenseTypeID
   [int]$requiredLicenses
   [string]$licenseGUID;

   LicenseTypeObject([string]$companyCode, [int]$companyCodeID, [string]$licenseType, [int]$licenseTypeID, [int]$requiredLicenses, [string]$licenseGUID)
   {
      $this.companyCode = $companyCode
      $this.companyCodeID = $companyCodeID
      $this.licenseType = $licenseType
      $this.licenseTypeID = $licenseTypeID
      $this.requiredLicenses = $requiredLicenses
      $this.licenseGUID = $licenseGUID
   }
}

# Create array of type objects
[System.Collections.ArrayList]$types = @()

foreach($companyCode in $companyCodes)
{
   foreach($licenseType in $licenseTypes)
   {
      $step = $types.Add([LicenseTypeObject]:: new($companyCode.FieldValues.CompanyCode, $companyCode.FieldValues.ID, $licenseType.FieldValues.Title, $licenseType.FieldValues.ID, 0, $licenseType.FieldValues.LicenseGUID))
   }
}

# Get required license type amounts depending on members in groups
foreach($type in $types)
{
   foreach($group in $groups)
   {
      if($type.companyCode -eq $group.companyCode -and $type.licenseType -eq $group.correspondingLicenseType)
      {
         $type.requiredLicenses += $group.groupMembers
      }
   }
}

# Write license type data to SharePoint Online
$countSPOadded = 0
foreach($type in $types)
{
   try
   {
      $step = Add-PnPListItem -List $spoListLicensesPerType -Values @{"Date" = $dateForSPO; "CompanyCode" = $type.companyCodeID; "LicenseType" = $type.licenseTypeID; "RequiredLicenses" = $type.requiredLicenses}
      $countSPOadded++
   }
   catch
   {
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

Write-Host $countSPOadded "type data entries added to SPO" -ForegroundColor Green

##########################
# Determine licenses per tenant
##########################
Write-Host "Determine licenses per tenant data..."
$tenantLicenses = Get-MgSubscribedSku

$countSPOadded = 0
foreach($tenantLicense in $tenantLicenses)
{
   try
   {
      if($types.licenseGUID -contains $tenantLicense.SkuId)
      {
         $licenseTypeEntry = $types | Where-Object {$_.licenseGUID -eq $tenantLicense.SkuId}
         $totalLicenses = $tenantLicense.PrepaidUnits.Enabled
         $assignedLicenses = $tenantLicense.ConsumedUnits
         $availableLicenses = $totalLicenses - $assignedLicenses
         $expiringSoon = $tenantLicense.PrepaidUnits.Suspended
         $step = Add-PnPListItem -List $spoListTenant -Values @{"Date" = $dateForSPO; "LicenseType" = $licenseTypeEntry.licenseTypeID; "TotalLicenses" = $totalLicenses; "AssignedLicenses" = $assignedLicenses; "AvailableLicenses" = $availableLicenses; "ExpiringSoon" = $expiringSoon}
         $countSPOadded++
      }
   }
   catch
   {
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

Write-Host $countSPOadded "tenant data entries added to SPO" -ForegroundColor Green

Write-Host "-------------------------------------------------" -ForegroundColor Green

Disconnect-PnPOnline
Disconnect-MgGraph
Stop-Transcript

Write-Host "-------------------------------------------------" -ForegroundColor Green

