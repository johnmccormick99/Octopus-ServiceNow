
$StepTemplateName = "VCS - ServiceNow Change Request Notes Update"
$StepTemplateDescription = "Source Controlled Step Template - ServiceNow Change Request Notes Update"
$StepTemplateParameters = @()

$_serviceNowChangeRequestNumber = $OctopusParameters["ServiceNowChangeRequestNumber"]
$_serviceNowApiUsername         = $OctopusParameters["ServiceNowApiUserName"];
$_serviceNowApiPassword         = $OctopusParameters["ServiceNowApiPassword"];
$_octopusUsername               = $OctopusParameters["Octopus.Deployment.CreatedBy.DisplayName"]
$_octopusProjectName            = $OctopusParameters["Octopus.Project.Name"]
$_octopusReleaseNumber          = $OctopusParameters["Octopus.Release.Number"]
$_octopusEnvironmentName        = $OctopusParameters["Octopus.Environment.Name"]

if (!($_serviceNowChangeRequestNumber))     {throw "Octopus Variable: ServiceNowChangeRequestNumber not defined"}
if (!($_serviceNowApiUsername))             {throw "Octopus Variable: ServiceNowApiUserName not defined"}
if (!($_serviceNowApiPassword))             {throw "Octopus Variable: ServiceNowApiPassword not defined"}

write-host ("Octopus Variable: ServiceNow Change Request = '$_serviceNowChangeRequestNumber'")
write-host ("Octopus Variable: ServiceNow Api UserName   = '$_serviceNowApiUsername'")
write-host ("Octopus Variable: Octopus Username          = '$_octopusUsername'")
write-host ("Octopus Variable: Octopus Project Name      = '$_octopusProjectName'")
write-host ("Octopus Variable: Octopus Release Number    = '$_octopusReleaseNumber'")
write-host ("Octopus Variable: Octopus Environment Name  = '$_octopusEnvironmentName'")

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

$secureServiceNowApiPassword = ConvertTo-SecureString -String $_serviceNowApiPassword -AsPlainText -Force 

$serviceNowCred = New-Object System.Management.Automation.PSCredential ($_serviceNowApiUsername, $secureServiceNowApiPassword)

$serviceNowPowershellModulePath = "E:\Tools\PowershellLibrary\Modules\ServiceNow\ServiceNow.psd1"

try
{
    Import-Module $serviceNowPowershellModulePath
}
catch 
{
    throw "Unable to find $serviceNowPowershellModulePath-- please ensure that the VCS - ServiceNow Change Request Notes Update step is executing on the Octopus Server and the Powershell module https://github.com/Sam-Martin/servicenow-powershell is present"
}

try
{
    $serviceNowChangeRequest = Get-ServiceNowChangeRequest -ServiceNowCredential $serviceNowCred -ServiceNowURL xxx.service-now.com -MatchContains @{number=$_serviceNowChangeRequestNumber}
}
catch 
{
    throw "The ServiceNow Api returned an exception - please check that the ServiceNow Change Request fields have been filled out correctly"
}

try 
{
    $serviceNowUpdateChangeRequest = Update-ServiceNowChangeRequest -ServiceNowCredential $serviceNowCred -ServiceNowURL investec.service-now.com -SysId $serviceNowChangeRequest.sys_id -Values @{work_notes = " [$_octopusUsername] deployed Octopus Deploy project [$_octopusProjectName] version [$_octopusReleaseNumber] to environment [$_octopusEnvironmentName] at [$(Get-Date)] "}
}
catch 
{
    throw "The ServiceNow Api returned an exception - please check that the ServiceNow Change Request fields have been filled out correctly"
}

write-host "ServiceNow $_serviceNowChangeRequestNumber work notes have been updated"