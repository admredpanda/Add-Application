Function Add-Application
{
<#
.SYNOPSIS
    The Add-Application function allows the creation of an application in SCCM and the associated structure from an XML file.

.DESCRIPTION
    The Add-Application function allows the creation of the following elements :
    - Creation of the application
    - Modification of the application
    - Creating the deployment type
    - Distribution of the application on the distribution point
    - Move the application to a specific folder
    - Creating an Active Directory group
    - Creating a user collection targeted to the Active Directory group
    - Move the user collection to a specific folder
    - Creating an application deployment on the user collection
    - Creating a peripheral collection
    - Move the peripheral collection to a specific folder
    - Creating an application deployment on the peripheral collection

.PARAMETER XMLfile
    Mandatory parameter. This is the path of the XML file with the application settings.

.EXAMPLE
    PS C:\> Add-Application -XMLFile "C:\ApplicationConf.xml"

.NOTES
    Author  : Valentin LÉPINE
    Email   : vlepineadm@outlook.com
    Date    : 08/11/2020
    Twitter : @vlepineadm
    Github  : https://github.com/vlepineadm/Add-Application
    Website : https://blog.labvl.net/
    Version : 1.0

    Microsoft Documentation :
    - https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c?redirectedfrom=MSDN
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/New-CMApplication
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Set-CMApplication
    - https://docs.microsoft.com/en-us/powershell/module/configurationmanager/add-cmscriptdeploymenttype
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Start-CMContentDistribution
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Get-CMApplication
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Move-CMObject
    - https://docs.microsoft.com/en-us/powershell/module/addsadministration/Get-ADGroup
    - https://docs.microsoft.com/en-us/powershell/module/addsadministration/new-adgroup
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Get-CMUserCollection
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/New-CMUserCollection
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Add-CMUserCollectionQueryMembershipRule
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Get-CMApplicationDeployment
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/New-CMApplicationDeployment
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/Get-CMDeviceCollection
    - https://docs.microsoft.com/en-us/powershell/module/ConfigurationManager/New-CMDeviceCollection
#>

    [CmdletBinding()]
    [OutputType( [System.Object] )]
    PARAM
    (
        [Parameter(Mandatory=$True)][System.String]$XMLfile
    )

    # -------------------------------------- #
    # --- Definition of global variables --- #
    # -------------------------------------- #

    # SCCM Variables #
    $ExeLocation = Get-Location
    $SiteServer = "CORPWSCM1" 
    $SiteCode = "COR"
    $SiteDrive = $SiteCode + ":"

    # Domain variables #
    $DomainName = "CORP"

    # Application variables #
    $DefaultLanguageId = "1033" # en-US
    $SuffixAppName = "APP-"
    $SuffixAppDeployType = "APP-DT-"

    # Group variables #
    $SuffixAppGrp = "GRP_SCCM_APP_"
    $PathAppGrp = "OU=SCCM,OU=Groups,DC=corp,DC=priv"

    # User collection variables #
    $SuffixColUser = "COL-USR-APP-"
    $FolderColUser = "" # "\PROD-Application"

    # Device collection variables #
    $SuffixColDevice = "COL-DEV-APP-"
    $FolderColDevice = "" # "\PRO-Application"


    # ------------------------- #
    # --- Import of modules --- #
    # ------------------------- #

    # Import of the SCCM module #
    try {
        Write-Host -ForegroundColor Green "Import of the SCCM module."
        Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386","\bin\configurationmanager.psd1")
    } catch {
        Write-Host -ForegroundColor Red "Module import error : $($_.exception.message)"
    }

    # Import of the Active Directory module #
    try {
        Write-Host -ForegroundColor Green "Import of the Active Directory module."
        Import-Module ActiveDirectory
    } catch {
        Write-Host -ForegroundColor Red "Module import error : $($_.exception.message)"
    }


    # Move to the SCCM site #
    if ((Get-PSDrive $SiteCode -ErrorAction SilentlyContinue | Measure-Object).Count -ne 1) {
        New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer
    }
    Set-Location $SiteDrive


    If (Get-content -Path $XMLfile) 
    {
        Write-Host -ForegroundColor Green "Retrieving the file : $XMLfile."

        # Import of the XML file into the $ App variable #
        $App = Select-Xml -Path $XMLfile -XPath "/AppDetails" | Select-Object -ExpandProperty Node

        # ------------------------------- #
        # --- Processing of variables --- #
        # ------------------------------- #

        # Application variables #
        $SCCMAppName = $SuffixAppName+$($App.Name).Replace(" ", "-")+"-v"+$App.Version
        $SCCMAppDeployType = $SuffixAppDeployType+$($App.Name).Replace(" ", "-")+"-v"+$App.Version
        $FolderApp = "\"+$App.Folder

        # Group variables #
        $AppDesc = $App.Name
        $AppName = $($App.Name).Replace(" ", "-")
        $GrpName = $SuffixAppGrp+$AppName
        $GrpDisName = $GrpName
        $GrpDes = "SCCM group for the application $AppDesc."

        # Collection variables #
        $ColUserName = "$SuffixColUser"+"$AppName" # Collection name user
        $ColUserComment = "User collection for the application $AppDesc linked to the group $GrpName." # User collection comment
        $ColDeviceName = "$SuffixColDevice"+"$AppName" # Device collection name
        $ColDeviceComment = "Device collection for the application $AppDesc." # Device collection comment


        # ----------------------------------------------------------- #
        # --------------- Creation of the application --------------- #
        # ----------------------------------------------------------- #

        # Checking the presence of the application #
        if (!(Get-CMApplication -Name $SCCMAppName))
        {
            # Creation of the application #
            try {
                Write-Host -ForegroundColor Green "Creation of the application $SCCMAppName."
                New-CMApplication -Name "$SCCMAppName" `
                    -Publisher "$($App.Publisher)" `
                    -SoftwareVersion "$($App.Version)" `
                    -AutoInstall $True `
                    -DefaultLanguageId "$($DefaultLanguageId)" `
                    -LocalizedName "$($App.Name)" `
                    -IconLocationFile "$($App.IconLocationFile)" `
                    -IsFeatured $false `
                    -ErrorAction Stop | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Error creating the application : $($_.exception.message)"
            }

            # Modification of the application #
            try {
                Write-Host -ForegroundColor Green "Modification of the application $SCCMAppName."
                Set-CMApplication -Name "$SCCMAppName" `
                    -DistributionPriority "$($App.DistributionPriority)" `
                    -ErrorAction Stop | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Application modification error : $($_.exception.message)"
            }  

            # Creating the deployment type #
            try {
                Write-Host -ForegroundColor Green "Creating the deployment type $SCCMAppDeployType."
                Add-CMScriptDeploymentType -ApplicationName "$SCCMAppName" `
                    -DeploymentTypeName $SCCMAppDeployType `
                    -ContentLocation "$($App.ContentLocation)" `
                    -InstallCommand "$($App.InstallCommand)" `
                    -UninstallCommand "$($App.UninstallCommand)" `
                    -InstallationBehaviorType "$($App.InstallationBehaviorType)" `
                    -ScriptLanguage PowerShell `
                    -ScriptText "$($App.DetectScript)" `
                    -InstallationProgramVisibility "$($App.InstallationProgramVisibility)" `
                    -LogonRequirementType "$($App.LogonRequirementType)" `
                    -SlowNetworkDeploymentMode Download `
                    -ErrorAction Stop | Out-Null
            } catch {
                 Write-Host -ForegroundColor Red "Application deployment type error : $($_.exception.message)"
            }  

            # Distribution of the application on the DP #
            try {
                Write-Host -ForegroundColor Green "Distribution of the application on the DP."
                Start-CMContentDistribution -ApplicationName "$SCCMAppName" `
                    -DistributionPointName "$($App.DistributionPointName)" `
                    -ErrorAction Stop | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Application distribution error : $($_.exception.message)"
            }  

            # Moving the application #
            try {
                Write-Host -ForegroundColor Green "Moving the application $SCCMAppName."
                $AppObj = Get-CMApplication -Name "$SCCMAppName"
                Move-CMObject -InputObject $AppObj `
                    -FolderPath $($SiteCode+':\Application'+$FolderApp) `
                    -ErrorAction Stop 
            } catch {
                Write-Host -ForegroundColor Red "Error moving application : $($_.exception.message)"
            }

        } else {
            Write-Host -ForegroundColor Red "The $SCCMAppName application already exists."
        }


        # ------------------------------------------------------------------ #
        # --------------- Creation of the group & collections -------------- #
        # ------------------------------------------------------------------ #

        # -------------------------------- #
        # --- Creation of the AD group --- #
        # -------------------------------- #

        # If the group exists #
        $ADGroupExist = $(try {Get-ADGroup $GrpName} catch {$null})
        If (!$ADGroupExist) {
            Write-Host -ForegroundColor Green "Group creation $GrpName."
            New-ADGroup -Name $GrpName `
                -DisplayName $GrpDisName `
                -Description $GrpDes `
                -Path $PathAppGrp `
                -GroupCategory Security `
                -GroupScope DomainLocal             
        } else {   
            Write-Host -ForegroundColor Red "The $GrpName already exists."
        }


        # --------------------------------------- #
        # --- Creation of the user collection --- #
        # --------------------------------------- #
 
        $ColUserExist = $(try {Get-CMUserCollection -Name "$ColUserName"} catch {$null})
        If (!$ColUserExist) {
            # Creation of the user Collection #
            try {
                Write-Host -ForegroundColor Green "Creation of the collection $ColUserName."
                $ColTab = @{Name = $ColUserName; Query = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.UserGroupName = '$DomainName\\$GrpName'"}
                New-CMUserCollection -Name $ColUserName `
                    -LimitingCollectionName "All Users and User Groups" `
                    -Comment $ColUserComment `
                    -RefreshType Continuous | Out-Null
                Add-CMUserCollectionQueryMembershipRule -CollectionName $ColTab.Name `
                    -QueryExpression $ColTab.Query `
                    -RuleName $ColTab.Name
            } catch {
                Write-Host -ForegroundColor Red "Error creating collection : $($_.exception.message)"
            }

            # Move the user collection #
            try {
                Write-Host -ForegroundColor Green "Moving the collection $ColUserName."
                $ColObj = Get-CMUserCollection -Name $ColUserName
                Move-CMObject -InputObject $ColObj `
                    -FolderPath $($SiteCode+':\UserCollection'+$FolderColUser)
            } catch {
                Write-Host -ForegroundColor Red "Error moving collection : $($_.exception.message)"
            }
        } else {   
            Write-Host -ForegroundColor Red "The $ColUserName collection already exists."
        }            


        # --------------------------------------------------------------------- #
        # --- Creation of the application deployment on the user collection --- #
        # --------------------------------------------------------------------- #

        $ColUserDeploymentExist = $(try {Get-CMApplicationDeployment -Name "$SCCMAppName" -CollectionName "$ColUserName"} catch {$null})
        If (!$ColUserDeploymentExist) {
            # Creation of the user deployment #
            try {
                Write-Host -ForegroundColor Green "Creation of the user deployment of the  $SCCMAppName application on the collection $ColUserName."
                New-CMApplicationDeployment -Name "$SCCMAppName" `
                    -CollectionName "$ColUserName" `
                    -DeployAction Install `
                    -DeployPurpose Available `
                    -UserNotification DisplaySoftwareCenterOnly `
                    -ErrorAction Stop | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Error creating deployment : $($_.exception.message)"
            }  
        } else {   
            Write-Host -ForegroundColor Red "Deployment of application $SCCMAppName on collection $ColUserName already exists."
        }


        # ----------------------------------------- #
        # --- Creation of the device collection --- #
        # ----------------------------------------- #

        $ColDeviceExist = $(try {Get-CMDeviceCollection -Name "$ColDeviceName"} catch {$null})
        If (!$ColDeviceExist) {     
            # Creation of the device collection #
            try {
                Write-Host -ForegroundColor Green "Creation of the collection $ColDeviceName."
                New-CMDeviceCollection -Name $ColDeviceName `
                    -LimitingCollectionName "All Systems" `
                    -Comment $ColDeviceComment `
                    -RefreshType Continuous | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Error creating collection : $($_.exception.message)"
            }

            # Move device collection #
            try {
                Write-Host -ForegroundColor Green "Move device collection $ColDeviceName."
                $ColObj = Get-CMDeviceCollection -Name $ColDeviceName
                Move-CMObject -InputObject $ColObj `
                    -FolderPath $($SiteCode+':\DeviceCollection'+$FolderColDevice)
            } catch {
                Write-Host -ForegroundColor Red "Error moving collection : $($_.exception.message)"
            }
        } else {   
            Write-Host -ForegroundColor Red "The $ColDeviceName collection already exists."
        }

        # -------------------------------------------------------------------- #
        # --- Creating the application deployment on the device collection --- #
        # -------------------------------------------------------------------- #

        $ColDeviceDeploymentExist = $(try {Get-CMApplicationDeployment -Name "$SCCMAppName" -CollectionName "$ColDeviceName"} catch {$null})
        If (!$ColDeviceDeploymentExist) {
            # Creating the device deployment #
            try {
                Write-Host -ForegroundColor Green "Creation of the computer deployment of the $SCCMAppName application on the collection $ColDeviceName."
                New-CMApplicationDeployment -Name "$SCCMAppName" `
                    -CollectionName "$ColDeviceName" `
                    -DeployAction Install `
                    -DeployPurpose Required `
                    -UserNotification DisplaySoftwareCenterOnly `
                    -ErrorAction Stop | Out-Null
            } catch {
                Write-Host -ForegroundColor Red "Error creating deployment : $($_.exception.message)"
            }
        } else {   
            Write-Host -ForegroundColor Red "Deployment of application $SCCMAppName on collection $ColDeviceName already exists."
        }


    } else {
        Write-Host -ForegroundColor Red "The $XMLfile file could not be found."
    }

    # Replacement in the original directory #
    Set-Location $ExeLocation

} ## End of the Add-Application function ##