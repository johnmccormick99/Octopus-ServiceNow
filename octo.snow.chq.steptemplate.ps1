
$StepTemplateName = "VCS - ServiceNow Change Request Approval"
$StepTemplateDescription = "Source Controlled Step Template - ServiceNow Change Request Approval"
$StepTemplateParameters = @(
    @{
        "Name"            = "ServiceNowChangeRequestNumber_";
        "Label"           = "ServiceNow Change Request Number";
        "HelpText"        = "e.g. CHG0053177";
        "DefaultValue"    = "#{ServiceNowChangeRequestNumber}";
        "DisplaySettings" = @{
            "Octopus.ControlType" = "SingleLineText";
        }
    },
    @{
        "Name"            = "ServiceNowApiUserName_";
        "Label"           = "ServiceNow Api Username";
        "HelpText"        = "";
        "DefaultValue"    = "#{ServiceNowApiUserName}";
        "DisplaySettings" = @{
            "Octopus.ControlType" = "SingleLineText";
        }
    },
    @{
        "Name"            = "ServiceNowApiPassword_";
        "Label"           = "ServiceNow Api Password";
        "HelpText"        = "";
        "DefaultValue"    = "#{ServiceNowApiPassword}";
        "DisplaySettings" = @{
            "Octopus.ControlType" = "SingleLineText";
        }
    }
)

$_serviceNowChangeRequestNumber = $OctopusParameters["ServiceNowChangeRequestNumber"];
$_serviceNowApiUsername         = $OctopusParameters["ServiceNowApiUserName"];
$_serviceNowApiPassword         = $OctopusParameters["ServiceNowApiPassword"];
$_octopusDisplayName              = $OctopusParameters["Octopus.Deployment.CreatedBy.DisplayName"];

if (!($_serviceNowChangeRequestNumber))     {throw "Octopus Variable: ServiceNowChangeRequestNumber not defined"}
if (!($_serviceNowApiUsername))             {throw "Octopus Variable: ServiceNowApiUserName not defined"}
if (!($_serviceNowApiPassword))             {throw "Octopus Variable: ServiceNowApiPassword not defined"}

# converting Octopus Display Name to ServiceNow Display Name
$_adUser          = Get-ADUser -f { Name -like $_octopusDisplayName } -ErrorAction Stop
$_adUserGivenname = $_adUser | Select-Object -ExpandProperty Givenname
$_adUserSurname   = $_adUser | Select-Object -ExpandProperty Surname
$_octopusUsername = $_adUserGivenname + " " + $_adUserSurname


write-host ("Octopus Variable: ServiceNow Change Request = '$_serviceNowChangeRequestNumber'");
write-host ("Octopus Variable: ServiceNow Api UserName   = '$_serviceNowApiUsername'");
write-host ("Octopus Variable: Octopus Username          = '$_octopusUsername'");

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
    throw "Unable to find $serviceNowPowershellModulePath-- please ensure that the VCS - ServiceNow Change Request Approval step is executing on the Octopus Server and the Powershell module https://github.com/Sam-Martin/servicenow-powershell is present"
}

try 
{
    $serviceNowChangeRequest = Get-ServiceNowChangeRequest -ServiceNowCredential $serviceNowCred -ServiceNowURL xxx.service-now.com -MatchContains @{number=$_serviceNowChangeRequestNumber}
}
catch 
{
    throw "The ServiceNow Api returned an exception - please check that the ServiceNow Change Request fields have been filled out correctly"
}    

$serviceNowChangeRequestState       = $serviceNowChangeRequest.state
$serviceNowChangeRequestStartDate   = $serviceNowChangeRequest.start_date
$serviceNowChangeRequestEndDate     = $serviceNowChangeRequest.end_date
$serviceNowChangeRequestImplementer = $serviceNowChangeRequest.assigned_to.display_value

if (!($serviceNowChangeRequestState))       {throw "ServiceNow Change Request $_serviceNowChangeRequestNumber State not defined"}
if (!($serviceNowChangeRequestStartDate))   {throw "ServiceNow Change Request $_serviceNowChangeRequestNumber Planned Start Date not defined"}
if (!($serviceNowChangeRequestEndDate))     {throw "ServiceNow Change Request $_serviceNowChangeRequestNumber Planned End Date not defined"}
if (!($serviceNowChangeRequestImplementer)) {throw "ServiceNow Change Request $_serviceNowChangeRequestNumber Primary Implementer not defined"}

write-host ("ServiceNow Change Request State               = '$serviceNowChangeRequestState'");
write-host ("ServiceNow Change Request Planned State Date  = '$serviceNowChangeRequestStartDate'");
write-host ("ServiceNow Change Request Planned End Date    = '$serviceNowChangeRequestEndDate'");
write-host ("ServiceNow Change Request Primary Implementer = '$serviceNowChangeRequestImplementer'");

if ($serviceNowChangeRequestState -ne "Implement")
{
    throw "ServiceNow $_serviceNowChangeRequestNumber is not ready to implement - current state: $serviceNowChangeRequestState"
}

Function IsBetweenDates ( [Datetime]$start,[Datetime]$end )
{
	$d = get-date
	if (($d -ge $start) -and ($d -le $end))
	{
		return $true
	}
	else
	{
		return $false
	}
}

[DateTime] $parsedStartDate = Get-Date ; [DateTime] $parsedEndDate = Get-Date

[Void] [DateTime]::TryParse($serviceNowChangeRequestStartDate,[ref]$parsedStartDate) ; [Void] [DateTime]::TryParse($serviceNowChangeRequestEndDate,[ref]$parsedEndDate)

$serviceNowChangeRequestIsBetweenDates = IsBetweenDates -start $parsedStartDate -end $parsedEndDate

if (!($serviceNowChangeRequestIsBetweenDates))
{
    throw "The deployment does not fall inside of the ServiceNow $_serviceNowChangeRequestNumber planned start and end date"
}

if ($_octopusUsername -ne $serviceNowChangeRequestImplementer)
{
    throw "The user who initiated the deployment does not match the ServiceNow $_serviceNowChangeRequestNumber primary implementer"
}

write-host "ServiceNow $_serviceNowChangeRequestNumber is ready to implement"