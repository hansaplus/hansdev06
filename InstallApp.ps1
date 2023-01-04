##################################################################################################
#
# Author : Encodian Solutions Ltd
# Date   : 2019-07-08
# Last	 : 2019-07-08
# Notes  : Script to:
#				Install an app to sites listed in a csv file
#				Entries must have a Status set to New to be processed
#
# Usage	:		Fill out the parameters then run the script
#
# History :
#			
################################################################################################## 

################################
# Parameters to specify
#
# Specify .csv location, for example:
$filePath = "C:\Users\admsasip\Desktop\"
$fileName = "deploy1.csv"

# Specify app name for example:
$appname = "bss-o365-PressApp"
# The SP Online tenant name, for example:
$orgName="bdfgrp"
#
# End of parameters
################################

$appId=""
$env:dp0 = [System.IO.Path]::GetDirectoryName($filePath)
$csvPath = "$env:dp0\$fileName"

Connect-PnPOnline -Url https://$orgName-admin.sharepoint.com -UseWebLogin

#
# Check if app already in app catalog - if so, get its app id
#
Get-PnPApp -Scope Tenant | ForEach-Object {
 $app = $_
 if ($app.Title -eq $appname)
 {
    Write-Host -ForegroundColor Green 'App already in the app catalog:' $app.Title $app.Id
    $appId = $app.Id;
 }
}

##################
# Functions
##################

# Installs the App to the site specified
function InstallApp([string] $appid, [string] $siteurl)
{
    if($appid -eq '' -or $siteurl -eq '')
    {
	    Write-Host -ForegroundColor Red 'All parameters are required, please try again.';
	    exit;
    }

    Try {

            $localConn = Connect-PnPOnline -Url $siteurl -UseWebLogin -ReturnConnection
            if (-not (Get-PnPContext)) {
                Write-Host "Error connecting to SharePoint Online, unable to establish context" -foregroundcolor black -backgroundcolor Red
                return
            } 
		    Write-Host -ForegroundColor Yellow 'Installing app on site' $siteurl

            Install-PnPApp -Identity $appid # -Connection $localConn

		    Write-Host -ForegroundColor Green 'App is installed on site' $siteurl
    } 
    catch
    {
        Write-Host -ForegroundColor Red 'Error encountered when trying to install app' $siteurl, ':' $Error[0].ToString();
    }
}

# Load csv file and iterate through each row to process
function ProcessDeploymentCSV([string] $csvPath, [string] $appid)
{
    Import-Csv $csvPath  | ForEach-Object {
        $csvRow = $_;
        if($csvRow.Status -eq "New")
        {
            Write-Host 'Site Name:' $csvRow.Name ' URL:' $csvRow.Url;
            InstallApp -appid $appid -siteurl $csvRow.Url
        }
    }
}

#####################
# Main Function Call
#####################

#
# If not in app catalog, cannot run script
#
if ($appId -eq "")
{
    Write-Host -ForegroundColor Yellow 'Please add the app to the app catalog before running this script'
}
else
{
    ProcessDeploymentCSV -csvPath $csvPath -appid $appId
}



