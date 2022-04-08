Add-PSSnapin Microsoft.SharePoint.PowerShell


$cs = @"
using System;
using System.Collections.Generic;

namespace SPDiagnostics
{
    public enum Severity
    {
        Default         = 0,
        Informational   = 1,
        Warning         = 2,
        Critical        = 4
    }

    public enum Format
    {
        Table           = 1,
        List            = 2
    }

    [Flags]
    public enum Category
    {
        Farm            = 1,
        Performance     = 2,
        Authentication  = 4,
        Search          = 8,
        Workflow        = 16
    }

    public class Finding
    {
        public Severity Severity { get; set; }
        public Category Category { get; set; }
        public string Name;
        public string Description;
        public string WarningMessage;
        public Uri ReferenceLink;
        public object InputObject;
        public bool Expand;
        public Format Format;
        public Finding[] ChildFindings;
    }

    public class FindingCollection : List<Finding>
    {
//        internal FindingCollection();
    }
}
"@

Add-Type -TypeDefinition $cs -Language CSharp





function Write-DiagnosticFindingFragment
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [SPDiagnostics.Finding]
        $Finding
    )

    switch ($Finding.Severity)
    {
        Critical {$class = "error"}
        Warning {$class = "warning"}
        default {$class = [string]::Empty}
    }

    $expandStr = [string]::Empty
    if($Finding.Expand)
    {
        $expandStr = " open"
    }

    $preContent = "<details{0}><summary class=`"heading {1}`">{2}</summary><div class=`"finding`">" -f $expandStr, $class, $Finding.Name
    if(![string]::IsNullOrEmpty($Finding.WarningMessage))
    {
        $preContent+="<div class=`"warning-message`">!!! {0} !!!</div>" -f $Finding.WarningMessage
    }
    $preContent+="<div class=`"description`">{0}</div>" -f $Finding.Description
    if(![string]::IsNullOrEmpty($Finding.ReferenceLink))
    {
        $preContent+="<div><a href=`"{0}`">{0}</a></div>" -f $Finding.ReferenceLink
    }

    $postContent = "</details></div>"
    
    $htmlFragment = $Finding.InputObject | ConvertTo-Html -PreContent $preContent -As $Finding.Format -Fragment

    foreach($child in $Finding.ChildFindings)
    {
        $childContent = Write-DiagnosticFindingFragment -Finding $child
        $htmlFragment+=$childContent
    }

    $htmlFragment+=$postContent

    return $htmlFragment

}



function Write-DiagnosticReport
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [SPDiagnostics.Finding[]]
        $Findings
    )

    $globalCss = @"
    <style>
        table {
            font-family: sans-serif;
            border: 2px solid;
            border-radius: 5px;
            border-style: solid;
            border-color: black;
        }

        th {
            padding-top: 6px;
            padding-bottom: 6px;
            text-align: center;
            background-color: rgb(120, 182, 177);
            color: black;
            border-radius: 5px;
        }

        body {
            font-family: sans-serif;
        }

        .error {
            color: Red;
        }

        .warning {
            color: Orange;
        }

        .warning-message {
            color: Red;
            font-weight: bold;
        }

        tbody tr:nth-child(even) {
            background: #bdbdbd;
        }

        .finding {
            padding-left: 30px;
        }

        .heading {
            font-size: larger;
            font-weight: bold;
            padding-top: 10px;
            padding-bottom: 10px;
            border-radius: 5px;
        }
    </style>
"@

    $html = "<!DOCTYPE html><head><Title=`"SPFarmReport`" /></head><body>"
    $html+=$globalCss
    
    foreach($finding in $Findings)
    {
        $fragment = Write-DiagnosticFindingFragment -Finding $finding
        $html+=$fragment
    }

    $html+="</body></html>"

    return $html
}

Function GetFarmBuild()
{
    $farm = [Microsoft.SharePoint.Administration.SPFarm]::Local
    $configDb = Get-SPDatabase | ?{$_.TypeName -match "Configuration Database"}
    ##Write-Host "Getting SP Farm Build"
    ##Write-Host ""
    $farmBuildVersion = $farm.BuildVersion
    $configDbName = $configDb.Name
    $configDbId = $configDb.Id
    $configDbSql = $configDb.Server.Address

$farmBuildTxt = @"
    [ SharePoint Farm Build: $farmBuildVersion ]
    
    ConfigDB:        $configDbName
    ConfigDbID:      $configDbId
    SQL Server:      $configDbSql
"@
    ##Write-Host $farmBuildTxt -ForegroundColor Cyan

    $retObj = [PSCustomObject]@{
        FarmBuildVersion = $farm.BuildVersion.ToString()
        ConfigDbName = $configDb.Name
        ConfigDbId = $configDb.Id
        ConfigDbSql = $configDb.ServiceInstance.Server.Address
        ConfigDbInstance = $configDb.ServiceInstance.Instance
    }
    

    return $retObj

}

function GetServersInFarm()
{
    ##Write-Host ""
    ##Write-Host "Getting Servers in the Farm"
    ##Write-Host ""
    ##Write-Host ""
    ##Write-Host "#########################################################################################"
    ##Write-Host "   Servers in the Farm "
    ##Write-Host "#########################################################################################"
    ##Write-Host ""
    $serverColl = @()
    foreach($svr in (Get-SPServer | Sort Role, Name))
    {
        $spProduct = Get-SPProduct
        
        if($svr.Role -ne "Invalid")
        {
            $productStatus = $null
            $productStatus = $spProduct.GetStatus($svr.DisplayName) | select -Unique

            $timeZone = $(Get-WMIObject -Class Win32_TimeZone -Computer $svr.DisplayName -ErrorAction SilentlyContinue).Description
            ##Write-Host $svr.DisplayName + " || " + $svr.Id + " || " + $svr.Role + " || " + $svr.Status + " || " + $productStatus + " || " + $timeZone
            if($productStatus -eq "UpgradeBlocked" -or $productStatus -eq "InstallRequired" -or $productStatus -eq "UpgradeInProgress")
            {
                $message = "'" + $productStatus.ToString() + "'" + " has been detected on server: " + $svr.DisplayName + ". This puts the farm\server in an 'UNSUPPORTED' and unstable state and patching\psconfig needs to be completed before any further troubleshooting. Support cannot provided until this is resolved"
                Write-Warning -Message $message
                ##Write-Host ""
                $productStatusBool = $true
            }

            $serverColl+=  [PSCustomObject]@{
                Name = $svr.DisplayName
                Role = $svr.Role
                Id = $svr.Id
                Status = $svr.Status
                ProductStatus = $productStatus
                TimeZone = $timeZone
            }

        }
        else
        {
            $timeZone = $(Get-WMIObject -Class Win32_TimeZone -Computer $svr.address -ErrorAction SilentlyContinue).Description
            ##Write-Host $svr.DisplayName + " || " + $svr.Id + " || " + $svr.Role + " || " + $svr.Status + " || " + $timeZone
            $serverColl+= [PSCustomObject]@{
                Name = $svr.DisplayName
                Role = $svr.Role
                Id = $svr.Id
                Status = $svr.Status
                ProductStatus = $null
                TimeZone = $timeZone
            }
            
        }
        
        
    }
    if($productStatusBool)
    {
        ##Write-Host ""
        ##Write-Host ""
        ##Write-Host " ** WARNING: We have detected that some servers are in an 'UpgradeBlocked\InstallRequired\UpgradeInProgress' state. This puts the farm\server in an 'UNSUPPORTED' and unstable state and patching\psconfig needs to be completed before any further troubleshooting. Support cannot provided until this is resolved! ** "
    }
    ##Write-Host ""


    return $serverColl
}



$inputObject = GetServersInFarm

$finding = New-Object SPDiagnostics.Finding
$finding.Category = [SPDiagnostics.Category]::Farm
$finding.Severity = [SPDiagnostics.Severity]::Critical
$finding.Name = "Servers in farm"
$finding.Expand = $false
$finding.Format = [SPDiagnostics.Format]::Table
$finding.WarningMessage = "Inconsistent patch state has been detected, servers must all be fully patched and upgraded"
$finding.InputObject = $inputObject


#$finding.InputObject

#Write-DiagnosticFindingFragment $finding | clip


$inputObject2 = GetFarmBuild

$finding2 = New-Object SPDiagnostics.Finding
$finding2.Category = [SPDiagnostics.Category]::Farm
$finding2.Severity = [SPDiagnostics.Severity]::Default
$finding2.Name = "Configuration database"
$finding2.Expand = $true
$finding2.Format = [SPDiagnostics.Format]::List
$finding2.InputObject = $inputObject2

$finding2.ChildFindings+=$finding
#$htmlfindings = Write-DiagnosticFindingFragment -Finding $finding2



$findingCollection = New-Object SPDiagnostics.FindingCollection
$findingCollection+=$finding
$findingCollection+=$finding2



$htmlContent = Write-DiagnosticReport -Findings $findingCollection


$fileName = "c:\temp\SPFarmReport_{0}" -f [datetime]::Now.ToString("yyyy_MM_dd_hh_mm_ss") + ".html"
Set-Content -Value $htmlContent -LiteralPath $fileName

Invoke-Item $fileName

