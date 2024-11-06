
<#PSScriptInfo

.VERSION 1.0

.GUID 1bc77904-e154-403e-923b-d8e86371981a

.AUTHOR timmcmic

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 This script downloads the Office 365 SKUs csv file for use in code. 

#> 
Param(
    [Parameter(Mandatory = $true)]
    [string]$logFolderPath=$NULL
)

$ErrorActionPreference = 'Stop'

#-------------------------------------------------------------------------------

Function new-LogFile
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$logFileName,
        [Parameter(Mandatory = $true)]
        [string]$logFolderPath
    )

    [string]$logFileSuffix=".log"
    [string]$fileName=$logFileName+$logFileSuffix

    # Get our log file path

    $logFolderPath = $logFolderPath+"\"+$logFileName+"\"
    
    #Since $logFile is defined in the calling function - this sets the log file name for the entire script
    
    $global:LogFile = Join-path $logFolderPath $fileName

    #Test the path to see if this exists if not create.

    [boolean]$pathExists = Test-Path -Path $logFolderPath

    if ($pathExists -eq $false)
    {
        try 
        {
            #Path did not exist - Creating

            New-Item -Path $logFolderPath -Type Directory
        }
        catch 
        {
            throw $_
        } 
    }
}

#-------------------------------------------------------------------------------
Function Out-LogFile
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        $String,
        [Parameter(Mandatory = $false)]
        [boolean]$isError=$FALSE
    )

    # Get the current date

    [string]$date = Get-Date -Format G

    # Build output string
    #In this case since I abuse the function to write data to screen and record it in log file
    #If the input is not a string type do not time it just throw it to the log.

    if ($string.gettype().name -eq "String")
    {
        [string]$logstring = ( "[" + $date + "] - " + $string)
    }
    else 
    {
        $logString = $String
    }

    # Write everything to our log file and the screen

    $logstring | Out-File -FilePath $global:LogFile -Append

    #Write to the screen the information passed to the log.

    if ($string.gettype().name -eq "String")
    {
        Write-Host $logString
    }
    else 
    {
        write-host $logString | select-object -expandProperty *
    }

    #If the output to the log is terminating exception - throw the same string.

    if ($isError -eq $TRUE)
    {
        #Ok - so here's the deal.
        #By default error action is continue.  IN all my function calls I use STOP for the most part.
        #In this case if we hit this error code - one of two things happen.
        #If the call is from another function that is not in a do while - the error is logged and we continue with exiting.
        #If the call is from a function in a do while - write-error rethrows the exception.  The exception is caught by the caller where a retry occurs.
        #This is how we end up logging an error then looping back around.

        write-error $logString

        #Now if we're not in a do while we end up here -> go ahead and create the status file this was not a retryable operation and is a hard failure.

        exit
    }
}

#-------------------------------------------------------------------------------
Function Test-PowershellVersion
    {
    [cmdletbinding()]

    $functionPowerShellVersion = $NULL

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN TEST-POWERSHELLVERSION"
    Out-LogFile -string "********************************************************************************"

    #Write function parameter information and variables to a log file.

    $functionPowerShellVersion = $PSVersionTable.PSVersion

    out-logfile -string "Determining powershell version."
    out-logfile -string ("Major: "+$functionPowerShellVersion.major)
    out-logfile -string ("Minor: "+$functionPowerShellVersion.minor)
    out-logfile -string ("Patch: "+$functionPowerShellVersion.patch)
    out-logfile -string $functionPowerShellVersion

    if ($functionPowerShellVersion.Major -ge 7)
    {
        out-logfile -string "Powershell 7 and higher is currently not supported due to module compatibility issues."
        out-logfile -string "Please run module from Powershell 5.x"
        out-logfile -string "" -isError:$true
    }
    else
    {
        out-logfile -string "Powershell version is not powershell 7.1 proceed."
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END TEST-POWERSHELLVERSION"
    Out-LogFile -string "********************************************************************************"

}

#-------------------------------------------------------------------------------

function get-Office365SKUHTMLData
{
    param(
        [Parameter(Mandatory = $true)]
        $office365CloudLocation
    )

    $functionHTMLData = $null

    out-logfile -string "Starting get-Office365SKUHTMLData"

    try {
        out-logfile -string "Invoking web request to obtain html data."
        $functionHTMLData = invoke-webRequest -Uri $office365CloudLocation -errorAction Stop
        out-logfile -string "Web data successfully retrieved."
    }
    catch {
        out-logfile -string "Unable to obtain azure html data."
        out-logfile -string $_ -isError:$true
    }

    out-logFile -string "Ending get-Office365SKUHTMLData"

    return $functionHTMLData
}

#-------------------------------------------------------------------------------

function get-Office365SKUDownloadLink
{
    param(
        [Parameter(Mandatory = $true)]
        $office365CloudLocation
    )

    $functionDownloadLink = $NULL
    $functionLinkString = "Here"

    out-logfile -string "Starting get-Office365SKUDownloadLink"

    $functionDownloadLink = $office365CloudLocation.links | where-object {$_.InnerText -eq $functionLinkString}

    out-logfile -string $functionDownloadLink

    $functionDownLoadLink = $functionDownLoadLink.href

    out-logfile -string $functionDownLoadLink

    out-logfile -string "Ending get-Office365SKUDownloadLink"

    return $functionDownloadLink
}

#-------------------------------------------------------------------------------
function get-Office365SKUCSVData
{
    param(
        [Parameter(Mandatory = $true)]
        $office365CloudLocation
    )

    $functionOffice365CSVData = $NULL

    out-logfile -string "Starting get-Office365SKUCSVData"

    try
    {
        out-logfile -string "Invoking web request to obtain json data..."

        $functionOffice365CSVData = invoke-webRequest -uri $office365CloudLocation -errorAction STOP

        out-logfile -string "Web request to obtain json data successful."
    }
    catch
    {
        out-logfile -string "Unable to invoke web request to obtain json data."
        out-logfile -string $_ -isError:$TRUE
    }

    out-logfile -string "Converting downloaded data to powershell json format."

    try
    {
        $functionOffice365CSVData = ConvertFrom-CSV $functionOffice365CSVData -errorAction STOP
    }
    catch
    {
        out-logfile -string "Unable to convert data to json format."
        out-logfile -string $_ -isError:$TRUE
    }


    out-logfile -string "Ending get-Office365SKUCSVData"

    return $functionOffice365CSVData
}

#-------------------------------------------------------------------------------
function export-Office365SKUCSVData
{
    param(
        [Parameter(Mandatory = $true)]
        $exportLocation,
        [Parameter(Mandatory = $true)]
        $csvData
    )

    out-logfile -string "Starting export-Office365SKUCSVData"

    try
    {
        $csvData | export-csv -Path $exportLocation
    }
    catch
    {
        out-logfile -string "Unable to export the json data."
        out-logfile -string $_ -isError:$TRUE
    }

    out-logfile -string "Ending export-Office365SKUCSVData"
}

#-------------------------------------------------------------------------------
function fix-Office365SKUCSVData
{
    param(
        [Parameter(Mandatory = $true)]
        $csvData
    )

    $functionOutputArray = @()

    out-logfile -string "Starting fix-Office365SKUCSVData"

    foreach ($member in $csvData)
    {
        out-logfile -string ("Fixing up: "+$member.service_plan_id)
        $functionObject = new-object PSObject -property @{
            Product_Display_Name = $member.'???Product_Display_Name'
            String_ID = $member.String_ID
            GUID = $member.GUID
            Service_Plan_Name = $member.Service_Plan_Name
            Service_Plan_ID = $member.service_plan_id
            Service_Plans_Included_Friendly_Names = $member.Service_Plans_Included_Friendly_Names
        }

        $functionOutputArray += $functionObject
    }

    out-logfile -string "Ending fix-Office365SKUCSVData"

    return $functionOutputArray
}

#-------------------------------------------------------------------------------

#*******************************************************************************
#Begin main script function
#*******************************************************************************

#Define function specific variables.

[string]$logFileName = ""
[string]$staticLogFileName = "Office365SKUData"
[string]$office365SKUInformation = "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"
$office365SKUHTMLData = $null
$office365SKUDataDownloadLink = $null
$office365CSVDataData = $NULL
#Define the log file name

$logFileName = $staticLogFileName

new-logfile -logFileName $logFileName -logFolderPath $logFolderPath

$office365CSVExport = $global:logFile.replace(".log",".csv")

out-logfile -string "*************************************************************"
out-logfile -string "Starting Office365SKUs"
out-logfile -string "*************************************************************"

out-logfile -string "Testing Powershell Version - 5.x required..."

Test-PowershellVersion

out-logfile -string "Obtaining the azure html data."

$office365SKUHTMLData = get-Office365SKUHTMLData -office365CloudLocation $office365SKUInformation

out-logfile -string "Obtaining data download link..."

$office365SKUDataDownloadLink = get-Office365SKUDownloadLink -office365CloudLocation $office365SKUHTMLData

out-logfile -string "Obtaining csv data..."

$office365CSVDataData = get-Office365SKUCSVData -office365CloudLocation $office365SKUDataDownloadLink

out-logfile -string "Fixup the downloaded data to fix column headers."

$office365CSVDataData = fix-Office365SKUCSVData -csvData $office365CSVDataData

out-logfile -string "Export the CSV data to the logging directory."

export-Office365SKUCSVData -exportLocation $office365CSVExport -csvData $office365CSVDataData

out-logfile -string "The Office 365 SKU data is now available for us.."