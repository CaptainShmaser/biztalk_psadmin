# Get local BizTalk DBName and DB Server from WMI
$btsSettings = get-wmiobject MSBTS_GroupSetting -namespace 'root\MicrosoftBizTalkServer'
$dbInstance = $btsSettings.MgmtDbServerName
$dbName = $btsSettings.MgmtDbName

# Load BizTalk ExplorerOM
[void] [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ExplorerOM")
$BizTalkOM = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer
$BizTalkOM.ConnectionString = "SERVER=$dbInstance;DATABASE=$dbName;Integrated Security=SSPI"

#This is useful on servers where installing assemblies to the GAC is required, but GACUtil is not present,
#e.g. a production server.
function Install-GacAssembly {
    param
    (
        $AssemblyName,
        $AssemblyLocation
    )
    [System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    $settingsdll = Get-ChildItem -Path $AssemblyLocation -Filter $AssemblyName -Recurse
    $publish = New-Object System.EnterpriseServices.Internal.Publish
    if ($settingsdll -is [System.Array]) {
        $publish.GacInstall($($settingsdll[0].FullName))
    }
    else {
        $publish.GacInstall($($settingsdll.FullName))
    }
}

#Returns all applications installed within a specific instance.
function Get-BizTalkInstanceApplications {
    Param
    (
        [alias("AppName")]
        [string]$AppPattern,
        [switch]$Force
    )
    if ($AppPattern -eq $null -or $AppPattern -eq "*") {
        if (!$Force) {
            return $BizTalkOM.Applications | Select-Object Name, Description, Status
        }
        else {
            return $BizTalkOM.Applications
        }
    }
    else {
        if (!$Force) {
            return $BizTalkOM.Applications | Where-Object {$_.Name -match "$AppPattern"} | Select-Object Name, Description, Status
        }
        else {
            return $BizTalkOM.Applications | Where-Object {$_.Name -match "$AppPattern"}
        }
    }
}

function Get-BizTalkApplicationReferences {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("AppName")]
        [string]$Name,
        [switch]$Force
    )

    $app = $null
    $app = $BizTalkOM.Applications | Where-Object {$_.Name -eq $Name}
    if (!$Force) {
        return $app.References
    }
    else {
        return $app.References | Select-Object Name, Description, Status
    }
}

function Get-BizTalkApplicationBackReferences {
    [cmdletbinding()]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("AppName")]
        [string]$Name,
        [switch]$Force
    )

    $app = $null
    $app = $BizTalkOM.Applications | Where-Object {$_.Name -eq $Name}
    if (!$Force) {
        $app.BackReferences
    }
    else {
        $app.BackReferences | Select-Object Name, Description, Status
    }
}

function Set-BizTalkApplicationState {
    # declare -stop -start switch parameters
    param
    (
        [switch] $start,
        [switch] $stop,
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("AppName")]
        [System.Object[]]$Name,
        [switch]$Force
    )

        
    if ($null -eq $Name -or $Name -eq "*" -or [System.String]::IsNullOrEmpty($Name) -eq $true) {
        if ($stop -and $Force) {
            $BizTalkOM.Applications | where-object {$_.status -eq "started"}  | ForEach-Object { $_.stop("StopAll")}
            $BizTalkOM.SaveChanges()
        }
        if ($stop -and $null -eq $Force -or $stop -and $false -eq $Force) {
            $BizTalkOM.Applications | ForEach-Object { $_.stop("DisableAllReceiveLocations")}
            $BizTalkOM.Applications | ForEach-Object { $_.stop("UndeployAllPolicies")}
            $BizTalkOM.Applications | ForEach-Object { $_.stop("UnenlistAllOrchestrations")}
            $BizTalkOM.Applications | ForEach-Object { $_.stop("UnenlistAllSendPortGroups")}
            $BizTalkOM.Applications | ForEach-Object { $_.stop("UnenlistAllSendPorts")}
            $BizTalkOM.SaveChanges()
        }
        if ($start -and $Force) {
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("StartAll")}
            $BizTalkOM.SaveChanges()
        }
        if ($start -and !$Force) {
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("DeployAllPolicies")}
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("EnableAllReceiveLocations")}
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("StartAllSendPortGroups")}
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("StartAllSendPorts")}
            $BizTalkOM.Applications | where-object {$_.status -eq "stopped"}  | ForEach-Object { $_.start("StartAllOrchestrations")}
            $BizTalkOM.SaveChanges()
        }
    }
    else {
        foreach ($var in $Name) {
            if ($stop -and $Force) {
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("StopAll")}
                $BizTalkOM.SaveChanges()
            }
            if ($stop -and $null -eq $Force -or $stop -and $false -eq $Force) {
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("DisableAllReceiveLocations")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("UndeployAllPolicies")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("UnenlistAllOrchestrations")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("UnenlistAllSendPortGroups")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.stop("UnenlistAllSendPorts")}
                $BizTalkOM.SaveChanges()
            }
            if ($start -and $null -ne $Name -and $Force) {
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("StartAll")}
                $BizTalkOM.SaveChanges()
            }
            if ($start -and $null -ne $Name -and !$Force) {
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("DeployAllPolicies")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("EnableAllReceiveLocations")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("StartAllSendPorts")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("StartAllSendPortGroups")}
                $BizTalkOM.Applications | where-object {$_.Name -eq "$var"}  | ForEach-Object { $_.start("StartAllOrchestrations")}
                $BizTalkOM.SaveChanges()
            }
        }
    }
}

function Get-BizTalkApplicationState {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("AppName")]
        [string]$Name
    )

    if ($null -eq $appName -or $appName -eq "*") {
        $BizTalkOM.Applications | Select-Object Name, Status
    }
    else {
        $BizTalkOM.Applications | Where-Object {$_.Name -eq "$Name"} | Select-Object Name, Status
    }

}

function Start-BizTalkOrchestration {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("OrchestrationName")]
        [string]$FullName
    )

    $TargetOrchestration = $null
    $Applications = $BizTalkOM.Applications
    $Orchestrations = $null

    foreach ($Application in $Applications) {
        $Orchestrations += $Application.Orchestrations
    }
    $TargetOrchestration = $Orchestrations | Where-Object {$_.FullName -eq $FullName}

    $TargetOrchestration.Status = "Started"
    $BizTalkOM.SaveChanges()
}

function Stop-BizTalkOrchestration {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("OrchestrationName")]
        [string]$FullName,
        [switch]$Force

    )

    $TargetOrchestration = $null
    $Applications = $BizTalkOM.Applications
    $Orchestrations = $null
   
    foreach ($Application in $Applications) {
        $Orchestrations += $Application.Orchestrations
    }
    $TargetOrchestration = $Orchestrations | Where-Object {$_.FullName -eq $FullName}

    if ($Force) {
        $TargetOrchestration.Status = "Unenlisted"
    }
    else {
        $TargetOrchestration.Status = "Enlisted"
    }
    $BizTalkOM.SaveChanges()
}

function Stop-BizTalkAppOrchestrations {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("AppName")]
        [string]$Name,
        [switch]$Force

    )

    $TargetOrchestration = $null
    $Application = Get-BizTalkInstanceApplications -AppPattern $Name -Force
    $Orchestrations = $Application.Orchestrations
  
    foreach($orch in $Orchestrations){
        if ($Force) {
            $orch.Status = "Unenlisted"
        }
        else {
            $orch.Status = "Enlisted"
        }
    }
    $BizTalkOM.SaveChanges()
}

function Get-BizTalkOrchestration {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("OrchestrationName", "Name")]
        [string]$FullName
    )

    $Applications = $BizTalkOM.Applications
    $Orchestrations = $null
    foreach ($Application in $Applications) {
        $Orchestrations += $Application.Orchestrations
    }
    $Orchestrations | Where-Object {$_.FullName -match "$FullName"}
}

function Start-BizTalkSendPort {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("SendPortName")]
        [string]$Name
    )


    $Port = $BizTalkOM.SendPorts | Where-Object {$_.Name -eq $Name}
    if ($null -ne $Port) {
        $Port.Status = "Started"
    }
    $BizTalkOM.SaveChanges()
}

function Stop-BizTalkSendPort {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("SendPortName")]
        [string]$Name,
        [switch]$Force
    )
    $Port = $null
    $Port = $BizTalkOM.SendPorts | Where-Object {$_.Name -eq $Name}

    if ($null -ne $Port) {
        if ($Force) {
            $Port.Status = "Bound"
        }
        else {
            $Port.Status = "Stopped"
        }
    }
    $BizTalkOM.SaveChanges()
}

function Start-BizTalkReceivePort {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ReceivePortName")]
        [string]$Name
    )

    $Port = $null
    $Port = $BizTalkOM.ReceivePorts | Where-Object {$_.Name -eq $Name}

    if ($null -ne $Port) {
        $ReceivePortLocations = $Port.ReceiveLocations
        foreach ($Location in $ReceivePortLocations) {
            $Location.Enable = $true
            $BizTalkOM.SaveChanges()
        }
    }
}

function Stop-BizTalkReceivePort {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ReceivePortName")]
        [string]$Name
    )

    $Port = $null
    $Port = $BizTalkOM.ReceivePorts | Where-Object {$_.Name -eq $Name}

    if ($null -ne $Port) {
        $ReceivePortLocations = $Port.ReceiveLocations
        foreach ($Location in $ReceivePortLocations) {
            $Location.Enable = $false
            $BizTalkOM.SaveChanges()
        }
    }
}

function Start-BizTalkSendPortGroup {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("SendPortGroupName")]
        [string]$Name
    )

    $SendPortGroup = $null
    $SendPortGroup = $BizTalkOM.SendPortGroups | Where-Object {$_.Name -eq $Name}
    $SendPortGroup.Status = "Started"
    $BizTalkOM.SaveChanges()
}

function Stop-BizTalkSendPortGroup {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("SendPortGroupName")]
        [string]$Name
    )

    $SendPortGroup = $null
    $SendPortGroup = $BizTalkOM.SendPortGroups | Where-Object {$_.Name -eq $Name}

    $SendPortGroup.Status = "Bound"
    $BizTalkOM.SaveChanges()
}

function Get-BizTalkDefaultApplication {
    $BizTalkOM.Applications | Where-Object {$_.IsDefaultApplication -eq "True"}
}

function Set-BizTalkDefaultApplication {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name
    )
    $BizTalkApp = $BizTalkOM.Applications[0].BtsCatalogExplorer.Applications | Where-Object {$_.Name -eq $Name}
    $BizTalkOM.Applications[0].BtsCatalogExplorer.DefaultApplication = $BizTalkApp
    $BizTalkOM.Applications[0].BtsCatalogExplorer.SaveChanges()
}

function Remove-BizTalkApplication {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name,
        [switch]$Force
    )

    if ($env:Processor_Architecture -ne "x86") {
        # Get the command parameters
        $ArgumentList = "-Name $Name"
        if ($force) {
            $ArgumentList += " -force"
        }
        $ModulePath = (Get-Module -Name "BizTalkAdministration_2013R2").Path
        write-warning "Re-launching in x86 PowerShell with $($ArgumentList -join ' ')"
        &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile -executionpolicy bypass -NoExit -Command "Import-Module $ModulePath; Remove-BizTalkApplication $ArgumentList ;exit"
        exit
    }

    $App = $null
    $App = Get-BizTalkInstanceApplications -AppPattern $Name -Force
    $appSendPorts = $App.SendPorts
    $appSendGroupPorts = $App.SendPortGroups
    $appReceivePorts = $App.ReceivePorts
    $appBackReferences = $App.BackReferences

    if ($appBackReferences.Count -gt 0 -and $Force -ne $true) {
        $Message = "$($App.Name) has back references which will block uninstallation. Please re-run this command with the -Force switch set to $true.
        ` WARNING this will remove ALL referencing applications prior to uninstalling the desired application!"
        Write-Warning $Message
        exit
    }
    if ($appBackReferences.Count -gt 0 -and $Force -eq $true) {
        foreach ($backApp in $appBackReferences) {
            Remove-BizTalkApplication -Name $backApp.Name -Force
        }
    }
    #Terminate All Messages
    $appSendPorts = $App.SendPorts
    foreach($sendPort in $appSendPorts)
    {
        Remove-BizTalkMessageInstances -ServiceStatus "*" -PortName $SendPort.Name
    }

    #Stop all Orchestrations
    $Orchestrations = $App.Orchestrations
    foreach ($Orchestration in $Orchestrations) {
        Stop-BizTalkOrchestration $Orchestration.FullName -Force
        Remove-OrchestrationBindings -Orchestration $Orchestration.FullName

    }

    #Remove all Send Port Groups
    foreach ($SendPortGroup in $appSendGroupPorts) {
        Stop-BizTalkSendPortGroup $SendPortGroup.name
        $BizTalkOM.RemoveSendPortGroup($SendPortGroup)
        $BizTalkOM.SaveChanges()
    }

    #Remove all Receive Ports
    foreach ($ReceivePort in $appReceivePorts) {
        Stop-BizTalkReceivePort $ReceivePort
        $BizTalkOM.RemoveReceivePort($ReceivePort)
        $BizTalkOM.SaveChanges()
    }

    #Remove all Send Ports
    foreach ($SendPort in $appSendPorts) {
        Stop-BizTalkSendPort -Name $SendPort.Name -Force
        $BizTalkOM.RemoveSendPort($SendPort)
        $BizTalkOM.SaveChanges()
    }
	
}

function Remove-BizTalkMessageInstances {
    param
    (
        [string]$ServiceStatus,
        [string]$PortName
    )

    # ServiceStatus = 1 Ready To Run
    # ServiceStatus = 2 Active
    # ServiceStatus = 4 Suspended (Resumable)
    # ServiceStatus = 8 Dehydrated
    # ServiceStatus = 16 Completed With Discarded MessagesÃ¢â‚¬â„¢ in BizTalk Server 2004
    # ServiceStatus = 32 Suspended (Not Resumable)
    # ServiceStatus = 64 In Breakpoint

    [array]$wmiQuery = $null
    if ($ServiceStatus -eq "*" -or [System.String]::IsNullOrEmpty($ServiceStatus)) {
        $wmiQuery = '1,2,4,8,32,64'
        $wmiQuery = $wmiQuery.Split(",")
    }
    else {
        $wmiQuery = $ServiceStatus.Split(",")
    }

    [array]$Messages = $null
    foreach ($statusCode in $wmiQuery) {
        $Messages += Get-WmiObject -Class MSBTS_ServiceInstance  -namespace "root\MicrosoftBizTalkServer" -filter "ServiceStatus =  $($statusCode)"
    }

    if ([System.String]::IsNullOrEmpty($PortName) -ne $true) {
        $Messages = $Messages | Where-Object {$_.ServiceName -eq $PortName}
    }
    foreach ($message in $Messages) {
        if ([string]::IsNullOrEmpty($PortName) -ne $true) {
            if ($message.ServiceName -eq $PortName) {
                $old_ErrorActionPreference = $ErrorActionPreference
                $ErrorActionPreference = "Ignore"
                $message.Terminate()
                $ErrorActionPreference = $old_ErrorActionPreference
            }
        }
        else {
                $old_ErrorActionPreference = $ErrorActionPreference
                $ErrorActionPreference = "Ignore"
                $message.Terminate()
                $ErrorActionPreference = $old_ErrorActionPreference
        }
    }
}

function Get-BizTalkMessageInstances {
    param
    (
        [string]$ServiceStatus,
        [string]$PortName
    )

   
    [array]$wmiQuery = $null
    if ($ServiceStatus -eq "*" -or [System.String]::IsNullOrEmpty($ServiceStatus)) {
        $wmiQuery = '1,2,4,8,32,64'
        $wmiQuery = $wmiQuery.Split(",")
    }
    else {
        $wmiQuery = $ServiceStatus.Split(",")
    }

    [array]$Messages = $null
    foreach ($statusCode in $wmiQuery) {
        $Messages += Get-WmiObject -Class MSBTS_ServiceInstance  -namespace "root\MicrosoftBizTalkServer" -filter "ServiceStatus =  $($statusCode)"
    }

    if ([System.String]::IsNullOrEmpty($PortName) -ne $true) {
        $Messages = $Messages | Where-Object {$_.ServiceName -eq $PortName}
    }
    else {
        $Messages
    }
}

function Remove-OrchestrationBindings {
    param
    (

        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("OrchestrationName")]
        [string]$FullName
    )

    $Orch = $null
    $Orch = Get-BizTalkOrchestration $FullName
    [array]$Ports = $Orch.Ports
    foreach ($Port in $Ports) {
        $Port.SendPort = $null
        $Port.ReceivePort = $null
        $BizTalkOM.SaveChanges()
    }
}

function Remove-BizTalkApplicationResource {
    param
    (
        [string]$ResourceLuid,
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name
    )
    if ($env:Processor_Architecture -ne "x86") {
        # Get the command parameters
        $ArgumentList = "-Name $Name"
        $ArgumentList += " -ResourceLuid $ResourceLuid"

        $ModulePath = (Get-Module -Name "BizTalkAdministration_2013R2").Path
        write-warning "Re-launching in x86 PowerShell with $($ArgumentList -join ' ')"
        &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile -executionpolicy bypass -NoExit -Command "Import-Module $ModulePath; Remove-BizTalkApplicationResource $ArgumentList ;exit"
        exit
    }

    $app = $null
    $app = Get-BizTalkInstanceApplications -AppPattern $Name -Force
        [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.Admin")
        [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ApplicationDeployment.Engine")
        [System.reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")
        $lDepGroup = New-Object Microsoft.BizTalk.ApplicationDeployment.Group
        $lDepGroup.DBName = $dbName
        Write-Host "Connecting to DB:$dbName"
        $lDepGroup.DBServer = $dbInstance
        Write-Host "Connecting to Server:$dbInstance"
        $appName = $Name
        Write-Host "Connecting to application:$($appName)"
    
        Write-Host "lDepGroup Object configured with the following"
        Write-Host "DbServer:$($lDepGroup.DBServer)"
        Write-Host "DbName:$($lDepGroup.DBName)"
    
        Write-Host "Connection Status"
        if ($null -eq $lDepGroup.SqlConnection.State) {
            Write-Host "Failed to open SQL connection to $($lDepGroup.DbServer),$($lDepGroup.DBName)"
        }
    
        [Microsoft.BizTalk.ApplicationDeployment.Application]$lApp = $lDepGroup.Applications["$appName"]
        if ($null -eq $lApp) {
            Write-Host "Specified app:$appName was not found in deployment group"
        }
        else {
            
        
        $resourceToRemove = $null
        $resources = $lApp.ResourceCollection

        foreach ($resource in $resources) {
            $installedLuid = $resource.Luid
            if ($installedLuid -eq $ResourceLuid) {
                $resourceToRemove = $resource
            }
        }
        $ErrorActionPreference = "Stop"
        try {
            Write-Verbose -Message "Attempting to remove $($resourceToRemove.Luid)"
            $lApp.RemoveResource($resourceToRemove.ResourceType, $resourceToRemove.Luid)
            Write-Output "Removal succeeded, committing transaction"
            $lDepGroup.Commit()
            Write-Output $true
        }
        catch {
            Write-Verbose -Message "Removal failed, rolling back transaction"
            $lDepGroup.Abort()
            Write-Output $false

        }
        finally {
            Write-Verbose -Message "Disposing connections."
            $lApp.Dispose()
            $lDepGroup.Dispose()
            $ErrorActionPreference = "Continue"
        }
    }
    $ErrorActionPreference = "SilentlyContinue"
    $lApp.Dispose()
    $lDepGroup.Dispose()
    $ErrorActionPreference = "Continue"
}
function Remove-AllBizTalkApplicationResources {
    param
    (
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name
    )
    $app = $null
    $app = Get-BizTalkInstanceApplications -AppPattern $Name -Force


    [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.Admin")
    [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ApplicationDeployment.Engine")
    [System.reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")
    $lDepGroup = New-Object Microsoft.BizTalk.ApplicationDeployment.Group
        
    $lDepGroup.DBName = $dbName
    Write-Host "Connecting to DB:$dbName"
    $lDepGroup.DBServer = $dbInstance
    Write-Host "Connecting to Server:$dbInstance"
    $appName = $Name
    Write-Host "Connecting to application:$($appName)"

    Write-Host "lDepGroup Object configured with the following"
    Write-Host "DbServer:$($lDepGroup.DBServer)"
    Write-Host "DbName:$($lDepGroup.DBName)"

    Write-Host "Connection Status"
    if ($null -eq $lDepGroup.SqlConnection.State) {
        Write-Host "Failed to open SQL connection to $($lDepGroup.DbServer),$($lDepGroup.DBName)"
    }

    [Microsoft.BizTalk.ApplicationDeployment.Application]$lApp = $lDepGroup.Applications["$appName"]
    if ($null -eq $lApp) {
        Write-Host "Specified app:$appName was not found in deployment group"
    }
    else {
        Write-Output "Application $appName contains the following resources."
        $resources = $lApp.ResourceCollection
        foreach ($resource in $resources) {
            Write-Output $resource.Luid
        }

        $ErrorActionPreference = "Stop"
        try {
            Write-Output "Attempting to remove all resources for $appName"
            $lApp.RemoveResources($resources)
            Write-Output "Removal succeeded, committing transaction"
            $lDepGroup.Commit()
            Write-Output $true
        }
        catch {
            Write-Verbose -Message "Removal failed, rolling back transaction"
            $lDepGroup.Abort()
            Write-Output $false

        }
        finally {
            Write-Verbose -Message "Disposing connections."
            $lApp.Dispose()
            $lDepGroup.Dispose()
            $ErrorActionPreference = "Continue"
        }
    }

    $ErrorActionPreference = "SilentlyContinue"
    $lApp.Dispose()
    $lDepGroup.Dispose()
    $ErrorActionPreference = "Continue"
}

function Update-BizTalkOMContext {
    $BizTalkOM.Refresh()
}

function Get-BizTalkOMContext {
    return $BizTalkOM
}

function Set-BizTalkOMContect{
    param(
    $OMContext
    )
    $BizTalkOM = $OMContext
}

function Get-BTSTaskLocation{
    $UtilName = "BTSTask.exe"
    $BaseLocation = Get-Item -Path "C:\Program Files (x86)\Microsoft BizTalk Server*"

    $BTSTaskLocation  = (Get-ChildItem -Path $BaseLocation -Filter $UtilName -Recurse).FullName
    return $BTSTaskLocation
}

function Stop-BizTalkApplication{
    param(
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name,
        [switch]$Force
    )
    
    Remove-BizTalkMessageInstances
    $app = Get-BizTalkInstanceApplications -AppPattern $Name -Force
    $Orchestrations = $app.Orchestrations
    Stop-BizTalkAppOrchestrations -Name $Name -Force:$Force

    $ReceivePorts = $app.ReceivePorts
    foreach($rPort in $ReceivePorts)
    {
        Stop-BizTalkReceivePort -Name $rPort.Name
    }

    $SendPortGroups = $app.SendPortGroups
    foreach($sgPort in $SendPortGroups)
    {
        Stop-BizTalkSendPortGroup -Name $sgPort.Name
    }

    $SendPorts = $app.SendPorts
    foreach($sPort in $SendPorts)
    {
        Stop-BizTalkSendPort -Name $sPort.Name -Force:$Force
        Remove-BizTalkMessageInstances -PortName $sPort.Name
    }

    ##TODO POLICY MANAGEMENT
    #$app.BtsCatalogExplorer.RuleDeploymentDriver.GetPublishedUndeployedRuleSets()....etc
}

function Start-BizTalkApplication{
    param(
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [alias("ApplicationName")]
        [string]$Name
    )
    
    $app = Get-BizTalkInstanceApplications -AppPattern $Name -Force
    $Orchestrations = $app.Orchestrations

    $ReceivePorts = $app.ReceivePorts
    foreach($rPort in $ReceivePorts)
    {
        Start-BizTalkReceivePort -Name $rPort.Name
    }

    $SendPorts = $app.SendPorts
    foreach($sPort in $SendPorts)
    {
        Start-BizTalkSendPort -Name $sPort.Name
    }

    foreach($Orch in $Orchestrations)
    {
        Start-BizTalkOrchestration -FullName $Orch.FullName
    }

    ##TODO POLICY MANAGEMENT
    #$app.BtsCatalogExplorer.RuleDeploymentDriver.GetPublishedUndeployedRuleSets()....etc
}

function Set-SendPortPropertyValue
{
     param(
        [string]$PortName,
        [string]$AppName,
        [string]$PropertyName,
        [String]$Token,
        [string]$ReplacementData
    )

    $Port = (Get-BizTalkInstanceApplications -AppPattern $AppName -Force).SendPorts | Where-Object {$_.Name -eq $PortName}

    if([System.String]::IsNullOrEmpty($Token) -ne $true)
    {
        [string]$PortData = $Port.PrimaryTransport."$PropertyName"
        $PortData = $PortData.Replace("$Token",$ReplacementData)
        $Port.PrimaryTransport."$PropertyName" = $PortData
    }
    else
    {
        $Port.PrimaryTransport."$PropertyName" = $ReplacementData    
    }
    
    $BizTalkOM.SaveChanges()
}

function Set-ReceiveLocationPropertyValue
{
     param(
        [string]$PortName,
        [string]$ReceiveLocationName,
        [string]$AppName,
        [string]$PropertyName,
        [String]$Token,
        [string]$ReplacementData
    )

    $Port = (Get-BizTalkInstanceApplications -AppPattern $AppName -Force).ReceivePorts | Where-Object {$_.Name -eq $PortName}
    $ReceiveLocation = $Port.ReceiveLocations | Where-Object {$_.Name -eq $ReceiveLocationName }
    if([System.String]::IsNullOrEmpty($Token) -ne $true)
    {
        [string]$PortData = $ReceiveLocation."$PropertyName"
        $PortData = $PortData.Replace("$Token",$ReplacementData)
        $ReceiveLocation."$PropertyName" = $PortData
    }
    else
    {
        $ReceiveLocation."$PropertyName" = $ReplacementData    
    }
    $BizTalkOM.SaveChanges()
}