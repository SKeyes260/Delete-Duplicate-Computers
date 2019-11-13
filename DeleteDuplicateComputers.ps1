# This script will identify all computers whose name appears in SCCM more than one time
# The record with the newest heartbeat will be kept and the rest will be deleted
# The Action parameter accepts two values "LoggingOnly"(default) OR  "Delete"

param([string]$SiteServer, [string]$SiteCode, [string]$Action, [string]$ExtDBName, [string]$ConfirmYesNo, [string]$InstanceName  )

Function Log-Append () {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $strLogFileName,
    [Parameter(Position=2)] $strLogText )
    
    Write-Host $strLogText
    $strLogText = ($(get-date).tostring()+" ; "+$strLogText.ToString()) 
    Out-File -InputObject $strLogText -FilePath $strLogFileName -Append -NoClobber
}



Function Get-DuplicateComputerNames() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SQLServer,
    [Parameter(Position=2)] $SCCMDBName )

    $objSCCMComputers = @()
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SCCMDBName;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = "
    SELECT Netbios_Name0, COUNT(Netbios_Name0) AS NameCount
    FROM v_R_System
    GROUP BY Netbios_Name0
    HAVING (COUNT(Netbios_Name0) > 1) AND  NetBios_Name0 <> 'Unknown'"

    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $objSCCMComputers = New-Object System.Data.Datatable
    $NumRows = $SqlAdapter.Fill($objSCCMComputers)
    $SqlConnection.Close()
    return $objSCCMComputers
}

Function Get-RecordsForName() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SQLServer,
    [Parameter(Position=2)] $SCCMDBName,
    [Parameter(Position=3)] $ExtDBName,
    [Parameter(Position=4)] $ComputerName )

    $objSCCMComputers = @()
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SCCMDBName;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = "
    SELECT dbo.v_R_System.Netbios_Name0, dbo.v_R_System.ResourceID, SCCM_EXT.dbo.v_LastHeartbeat.Agent_Time AS LastHeartBeat, dbo.v_R_System.Creation_Date0
    FROM SCCM_EXT.dbo.v_LastHeartbeat 
    RIGHT OUTER JOIN dbo.v_R_System ON SCCM_EXT.dbo.v_LastHeartbeat.ResourceID = dbo.v_R_System.ResourceID
    WHERE (dbo.v_R_System.Netbios_Name0 = N'$ComputerName')
    ORDER BY LastHeartBeat DESC, dbo.v_R_System.Creation_Date0 DESC"
    #Write-Host $SQLCMD.CommandText
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $objSCCMComputers = New-Object System.Data.Datatable
    $NumRows = $SqlAdapter.Fill($objSCCMComputers)
    $SqlConnection.Close()
    return $objSCCMComputers

}


###############################  
#      MAIN
###############################

# Set parameter defaults
$PassedParams =  (" -SiteServer "+$SiteServer+" -SiteCode "+$SiteCode+" -SQLServer "+$SQLServer+" -InstanceName "+$InstanceName)
If (!$SiteServer)    { $SiteServer    = "XSNW10W142C.pharma.aventis.com"}
If (!$SiteCode)      { $SiteCode      = "PP0"                           }
If (!$Action)        { $Action        = "LoggingOnly"                   }  # LoggingOnly   or  Delete
If (!$ExtDBName)     { $ExtDBName     = "SCCM_EXT"                      }
If (!$ConfirmYesNo)  { $ConfirmYesNo  = "Yes"                           }
If (!$InstanceName)  { $InstanceName  = "DeleteDuplicateComputers-$SiteCode"      }

#Define environment specific  variables
$TodaysDate         = Get-Date
$ScriptPath         = $PSScriptRoot
$LoggingFolder      = ($ScriptPath+"\Logs\")
$ToKeep             = 0
$ToDelete           = 0
$ActuallyDeleted    = 0

# Lookup the SCCM site definition to populate global variables
$objSiteDefinition = Get-WmiObject -ComputerName $SiteServer -Namespace ("root\sms\Site_"+$SiteCode) -Query ("SELECT * FROM sms_sci_sitedefinition WHERE SiteCode = '"+$SiteCode+"'")
$SCCMDBName = $objSiteDefinition.SQLDatabaseName
$SQLServer = $objSiteDefinition.SQLServerName

# Get the logging prepared
If (!(Test-Path $LoggingFolder))  { $Result = New-Item $LoggingFolder -type directory }
$LogFileName = ($LoggingFolder+$InstanceName+"-"+$SiteCode+"-"+$TodaysDate.Year+$TodaysDate.Month.ToString().PadLeft(2,"0")+$TodaysDate.Day.ToString().PadLeft(2,"0")+".log")

# connct to SCCM
$initParams = @{}
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer @initParams
}
Set-Location "$($SiteCode):\" @initParams

#get a list of computer names that have a duplicate record
$DuplicateNames = @()
$DuplicateNames = Get-DuplicateComputerNames -SQLServer $SQLServer -SCCMDBName $SCCMDBName

# for each computer name get all matching names sorted DESC by Last Heartbeat
# delete all but the first one
ForEach ( $DuplicateName in $DuplicateNames ) {
    $NameGroup = Get-RecordsForName -SQLServer $SQLServer -SCCMDBName $SCCMDBName -ExtDBName $ExtDBName -ComputerName $DuplicateName.Netbios_Name0
    $I = 0
    Log-Append -strLogFileName $LogFileName -strLogText ""
    ForEach ( $ComputerName in $NameGroup )  {
        If ( $I -ne 0 ) {
            Log-Append -strLogFileName $LogFileName -strLogText ("Will delete computer named "+$ComputerName.Netbios_Name0.PadRight(16," ")+" with resourceID "+$ComputerName.ResourceID+" whose last Heartbeat was on " + $ComputerName.LastHeartbeat+" and was created on "+$ComputerName.Creation_Date0)
            If ($Action -eq "Delete" ) {
                If ( $ConfirmYesNo -ne "No" ) { 
                    $Result = Remove-CMResource -ResourceId $Computername.ResourceID -Force -Confirm 
                    If (Get-CMResource -ResourceId $ComputerName.ResourceID -Fast) { Write-Host ("Failed to delete resource " + $Computername.NetBios_name0  +" with a ResourceID of "  +$ComputerName.ResourceID) }
                    ELSE { 
                        Log-Append -strLogFileName $LogFileName -strLogText  ("Succesfully deleted resource "+$Computername.NetBios_name0+" with a ResourceID of "+$ComputerName.ResourceID )
                        $ActuallyDeleted += 1
                    } 
                }
                If ( $ConfirmYesNo -eq "No" ) {
                    $Result = Remove-CMResource -ResourceId $Computername.ResourceID -Force
                    If (Get-CMResource -ResourceId $ComputerName.ResourceID -Fast) { Write-Host ("Failed to delete resource " + $Computername.NetBios_name0  +" with a ResourceID of "  +$ComputerName.ResourceID) }
                    ELSE { 
                        Log-Append -strLogFileName $LogFileName -strLogText  ("Succesfully deleted resource "+$Computername.NetBios_name0+" with a ResourceID of "+$ComputerName.ResourceID )
                        $ActuallyDeleted += 1
                    } 
                }
            }
            $ToDelete += 1
        }
        ELSE {
            Log-Append -strLogFileName $LogFileName -strLogText  ("Will keep computer named   "+$ComputerName.Netbios_Name0.PadRight(16," ")+" with resourceID "+$ComputerName.ResourceID+" whose last Heartbeat was on " + $ComputerName.LastHeartbeat+" and was created on "+$ComputerName.Creation_Date0)
            $ToKeep += 1
        }
        $I += 1
    }
}
Write-Host
Log-Append -strLogFileName $LogFileName -strLogText  "Script complete... Recommended keeping $ToKeep and deleting $ToDelete records... actually deleted $ActuallyDeleted records"