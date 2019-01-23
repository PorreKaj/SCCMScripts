#Requires -Modules ActiveDirectory

# This Scripts requires that the department AD Attribute has been added in 'Active Directory User Discovery'
# Line 33 creates a list of unique Departmentnames, edit as needed
# Line 37 grabs what in my case is the Department Number which is used for naming the collection, and used in the query
# Line 44 moves the collection to a folder - Folder needs to exist beforehand
#

# Site configuration
$SiteCode = (Read-Host -Prompt "Site Code") # Site code 
$ProviderMachineName = (Read-Host -Prompt "Site Server") # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.

Set-Location "$($SiteCode):\" @initParams
$Departments = Get-ADuser -filter "Department -like '260*'" -Properties Department | select -ExpandProperty Department -Unique | Sort
$Schedule1 = New-CMSchedule -Start "01/01/2019 7:00 AM" -RecurInterval Days -RecurCount 1

Foreach ($Department in $Departments){
    $DepartmentNumber = $Department.Substring(0,7)

    if((Get-CMCollection -Name $DepartmentNumber) -eq $null){
    Write-Host "$DepartmentNumber does not exist"
    
        
        $NewCollection = New-CMCollection -LimitingCollectionId SMS00002 -Name "$DepartmentNumber" -RefreshSchedule $Schedule1 -RefreshType Periodic -CollectionType User 
        Move-CMObject -FolderPath '.\UserCollection\Department Collections' -InputObject $NewCollection
        Add-CMUserCollectionQueryMembershipRule -CollectionName $DepartmentNumber -RuleName "Department$DepartmentNumber" -QueryExpression "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.department like '$DepartmentNumber%'"

    }
}

