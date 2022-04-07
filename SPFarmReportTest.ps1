## Globals
$Global:DiagSeverity = ("Critical", "Warning", "Informational")


function Write-DiagnosticFinding
{
    Param
    (
        [validateset("Critical", "Warning", "Informational")]
        [string]
        $Severity,

        [validateset("Performance", "Authentication", "Farm", "Search", "Workflow")]
        [string]
        $Category,

        [string]
        $Name,

        [string]
        $Description,

        [string]
        $WarningMessage,

        [object]
        $BodyTable
    )

    Write-Host $Severity -ForegroundColor Green
}



function GetServersInFarm()
{
    Write-Host ""
    Write-Host "Getting Servers in the Farm"
    Write-Host ""
    Write-Host ""
    Write-Host "#########################################################################################"
    Write-Host "   Servers in the Farm "
    Write-Host "#########################################################################################"
    Write-Host ""
    $serverColl = @()
    foreach($svr in (Get-SPServer | Sort Name, Role))
    {
        $spProduct = Get-SPProduct
        
        if($svr.Role -ne "Invalid")
        {
            $productStatus = $null
            $productStatus = $spProduct.GetStatus($svr.DisplayName) | select -Unique

            $timeZone = $(Get-WMIObject -Class Win32_TimeZone -Computer $svr.DisplayName -ErrorAction SilentlyContinue).Description
            Write-Host $svr.DisplayName + " || " + $svr.Id + " || " + $svr.Role + " || " + $svr.Status + " || " + $productStatus + " || " + $timeZone
            if($productStatus -eq "UpgradeBlocked" -or $productStatus -eq "InstallRequired" -or $productStatus -eq "UpgradeInProgress")
            {
                $message = "'" + $productStatus.ToString() + "'" + " has been detected on server: " + $svr.DisplayName + ". This puts the farm\server in an 'UNSUPPORTED' and unstable state and patching\psconfig needs to be completed before any further troubleshooting. Support cannot provided until this is resolved"
                Write-Warning -Message $message
                Write-Host ""
                $productStatusBool = $true
            }

            $serverColl+=  [PSCustomObject]@{
                Name = $svr.DisplayName
                Id = $svr.Id
                Status = $svr.Status
                ProductStatus = $productStatus
                TimeZone = $timeZone
            }

        }
        else
        {
            $timeZone = $(Get-WMIObject -Class Win32_TimeZone -Computer $svr.address -ErrorAction SilentlyContinue).Description
            Write-Host $svr.DisplayName + " || " + $svr.Id + " || " + $svr.Role + " || " + $svr.Status + " || " + $timeZone
            $serverColl+= [PSCustomObject]@{
                Name = $svr.DisplayName
                Id = $svr.Id
                Status = $svr.Status
                ProductStatus = $null
                TimeZone = $timeZone
            }
            
        }
        
        
    }
    if($productStatusBool)
    {
        Write-Host ""
        Write-Host ""
        Write-Host " ** WARNING: We have detected that some servers are in an 'UpgradeBlocked\InstallRequired\UpgradeInProgress' state. This puts the farm\server in an 'UNSUPPORTED' and unstable state and patching\psconfig needs to be completed before any further troubleshooting. Support cannot provided until this is resolved! ** "
    }
    Write-Host ""
    return $serverColl
}


$serversInFarm = GetServersInFarm



Write-DiagnosticFinding -Severity Informational -Category Farm -Name "Servers in farm" -Description "All servers in farm" -WarningMessage "Server found with upgrade or install required status" -body $serversInFarm






$farm = Get-SPFarm
$configDb = Get-SPDatabase | ?{$_.TypeName -match "Configuration Database"}

Function GetFarmBuild()
{
    Write-Host "Getting SP Farm Build"
    Write-Host ""
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
    Write-Host $farmBuildTxt -ForegroundColor Cyan

    $retObj = [PSCustomObject]@{
        FarmBuildVersion = $farm.BuildVersion.ToString()
        ConfigDbName = $configDb.Name
        ConfigDbId = $configDb.Id
        ConfigDbSql = $configDb.ServiceInstance.Server.Address
        ConfigDbInstance = $configDb.ServiceInstance.Instance
    }
    
    return $retObj

}




$farmBuildInfoHtml = GetFarmBuild | ConvertTo-Html -Fragment -As List -PreContent "<h2>Farm Information:</h2>"
$serversInFarmHtml = $serversInFarm | ConvertTo-Html  -Property "Name", "Id", "Status", "ProductStatus", "TimeZone" -PreContent "<h2>Servers In Farm</h2>" -Fragment




$globalStyle = @"
    <style>
        table {
            font-family: sans-serif;
            border: 2px solid;
            border-radius: 5px;
            border-style: solid;
            border-color: black;
            /*width: 100;*/

        }

        th {
            padding-top: 6px;
            padding-bottom: 6px;
            text-align: center;
            background-color: rgb(120, 182, 177);
            color: black;
        }

        body {
            font-family: sans-serif;
        }
    </style>
"@



$htmlContent = "<html><body>"
$htmlContent+=$globalStyle
$htmlContent+=$farmBuildInfoHtml
$htmlContent+=$serversInFarmHtml
$htmlContent+="</body></html>"

$htmlContent | clip


