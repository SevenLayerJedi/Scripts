function Out-Bootstrap
{
  <#
  .SYNOPSIS
  Backs up CPS via SSH
  .DESCRIPTION
  Outputs an html report that is pulled from a mssql db. Displays it with bootstrap css
  .LINK
  http://www.KeithSmithOnline.com
  https://twitter.com/@SevenLayerJedi
  .PARAMETER NA
  NA
  .EXAMPLE
  Out-Bootstrap
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

function Project-YellowNikon
{
 [cmdletbinding()]
  
  param (
	[parameter(Mandatory=$true)]
  [String]
  $Title="",
  
  [parameter(Mandatory=$false)]
  [String]
  $SubTitle="",
  
  [parameter(Mandatory=$false)]
  [String]
  $Author="",
  
  [parameter(Mandatory=$true)]
  [Array]
  $Object
  )

  ####################################################
  # --Variables--                                    #
  ####################################################  

$global:HTML = ""

$html_start = @"
<!DOCTYPE html>
<html lang="en">`r
"@

$head_start = @"
<head>
"@

$head_content = @"
  <title>STRING</title>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>`r
"@


$head_stop = @"
</head>`r
"@

$body_start = @"
<body>`r
"@

$body_stop = @"
</body>`r
"@

$html_stop = @"
</html>`r
"@

$tag_h2 = @"
<h2>STRING</h2>`r
"@

$tag_h4 = @"
<h4>STRING</h4>`r
"@

$tag_h5 = @"
<h5>STRING</h5>`r
"@


$div_container_start = @"
<div class="container">`r
"@
 
$div_container_stop = @"
</div>`r
"@

 
$div_panel_primary_start = @"
    <div class="panel panel-primary">`r
"@
 
$div_panel_primary_stop = @"
    </div>`r
"@
 
 
$div_panel_heading = @"
<div class="panel-heading">JSJSJSJSJSJSJSJS</div>`r
"@
 
$table_striped_start = @"
      <table class="table table-striped table-hover">`r
"@
 
$table_striped_stop = @"
      </table>`r
"@
 
 
$table_head_start = @"
        <thead>
          <tr>`r
"@

$table_heading = @"
            <th>STRING</th>`r
"@

$table_head_stop = @"
          </tr>
        </thead>`r
"@

$table_row_start = @"
        <tr>`r
"@

$table_row_data = @"
          <td>STRING</td>`r
"@

$table_row_stop = @"
        </tr>`r
"@

$progressbar_start = @"
<div class="row">
<div class="col-md-2"></div>
<div class="col-md-8">
<div class="progress">`r
"@

$progressbar_stop = @"
</div>           
</div>
<div class="col-md-2"></div>
</div>`r
"@
 
$progressbar_percent_danger = @"
<div class="progress-bar progress-bar-danger" role="progressbar" style="width:PARAM1%">STRING</div>`r
"@
 
$progressbar_percent_success = @"
<div class="progress-bar progress-bar-success" role="progressbar" style="width:PARAM1%">STRING</div>`r
"@


  ####################################################
  # --Functions--                                    #
  ####################################################

Function Create-HTML
{
  $global:HTML += $html_start
  Create-HEAD
  Create-BODY
  $global:HTML += $html_stop
  RETURN $global:HTML
}

Function Create-HEAD
{

  $global:HTML += $head_start
  if ($Title.count -gt 0){$global:HTML += $head_content.replace('STRING',"$($Title)")}ELSE{$HTML += $head_content.replace('STRING',"Report")}
  $global:HTML += $head_stop
}

function Create-BODY
{
  $global:HTML += $body_start
  Create-CONTENT
  $global:HTML += $body_stop
}

function Create-CONTENT
{
  Create-CONTAINER
  
}

function Create-CONTAINER
{
  $global:HTML += $div_container_start
  if ($Title.count -gt 0){$global:HTML += $tag_h2.replace('STRING',"$($Title)")}ELSE{}
  if ($SubTitle.count -gt 0){$global:HTML += $tag_h4.replace('STRING',"$($SubTitle)")}ELSE{}
  if ($Author.count -gt 0){$global:HTML += $tag_h5.replace('STRING',"$($Author)")}ELSE{}
  
  #foreach ($obj in $Object)
  #{
    Create-PANEL
  #}
  $global:HTML += $div_container_stop
}

function Create-PANEL
{
  $global:HTML += $div_panel_primary_start
  
  $StartTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addhours(-24).ticks)"
  $StopTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addseconds(-1).ticks)"

  $StartTimeFormated = "$(Get-date $([long]$StartTicks) -Format G)" + " - "
  $StopTimeFormated = "$(Get-date $([long]$StopTicks) -Format `"HH:mm:ss`")" + " PM"
  $FullTimeFormated = "$StartTimeFormated" + "$StopTimeFormated"

  $global:HTML += $div_panel_heading.replace('JSJSJSJSJSJSJSJS',"$($FullTimeFormated)") # not sure what to call the heading on each panel ???
  Create-TABLE
  $global:HTML += $div_panel_primary_stop
}

function Create-TABLE
{
  $global:HTML += $table_striped_start
  $global:HTML += $table_head_start
  
  foreach ($thing in $Object[0].psobject.Properties)
  {
    $global:HTML += $table_heading.replace('STRING',"$($thing.name)")
  }

  $global:HTML += $table_head_stop
  
  
  foreach ($row in $Object)
  {
    $global:HTML += $table_row_start
    foreach ($prop in $row.PSObject.Properties)
    {
      $global:HTML += $table_row_data.replace('STRING',"$($prop.value)")
    }
    $global:HTML += $table_row_stop
  }

  
  $global:HTML += $table_striped_stop
}



Create-HTML


}


  function SilentlyContinueErrors
  {
    #Set Error Action to Silently Continue
    $ErrorActionPreference = 'SilentlyContinue'
    
    Write-host "[Disabling Errors]" -foregroundcolor green
  }
  #SilentlyContinueErrors

  function LoadExchangePSModule
  {
    Write-host "[Adding Exchange PSSnapin]" -foregroundcolor green
    
    # Load Exchange Powershell module
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
  }
  LoadExchangePSModule

  #add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

  #C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell" 

  #& 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
  #Connect-ExchangeServer -auto -ClientApplication:ManagementShell

  function OpenSQLConnection
  {
    Write-host "[Opening SQL Connection]" -foregroundcolor green
    
    $SQLServer = "SERVER1" #use Server\Instance for named SQL instances!
    $SQLDBName = "DB01"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SQLDBName;Integrated Security=true"

    ###################
    # Open SQL Socket
    ###################

    "[OPENCONNECTION] // Connection Opened - $(Get-Date)" | out-file -append c:\reports\Import-ExchangeMailboxAuditToSQL.log
    
    $SqlConnection.Open()
    $SqlConnection
  }
  $SqlConnection = OpenSQLConnection


  function CreateSQLCommandObject($SqlConnection)
  {
    Write-host "[Creating SQL CMD Object]" -foregroundcolor green
    ###################
    # Create SQL Command Object
    ###################

    # Create NULL command object
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand

    # Add the Connection Settings
    $SqlCmd.Connection = $SqlConnection

    $SqlCmd
  }
  $SqlCmd = CreateSQLCommandObject $SqlConnection


  function Send-SQLQuery($SqlCmd, $SqlQuery)
  {
    Write-host "[Sending SQL CMD]" -nonewline -foregroundcolor yellow;Write-host "$SQLQuery" -foregroundcolor green
    ##############
    # Send SQL Query *
    ##############

    # Command to Send
    $SqlCmd.CommandText = $SqlQuery

    #Write-host "Gonna Run"
    #Write-host "$($SqlCmd.CommandText)"
    
    # Execute the SQL Command
    $SqlCmd.ExecuteNonQuery()
    
    $Data = $NULL
    $Data = $SqlCmd.ExecuteReader()
    $DataTable = new-object “System.Data.DataTable”
    $DataTable.Load($Data)
    $DataTable | format-table
    
    # Clear out SqlCmd
    $SqlCmd.CommandText = $NULL
  }

  function CloseSQLConnection
  {
    ###################
    # Close SQL Socket
    ###################

    Write-host "[Closing SQL Connection]" -foregroundcolor green
    
    $SqlConnection.Close()
  }


  ##############################################################################################
  ##############################################################################################
  ##############################################################################################
  ##############################################################################################
  ##############################################################################################


  function Get-YesterdayAuditReport($SqlCmd)
  {
    $StartTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addhours(-24).ticks)"
    $StopTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addseconds(-1).ticks)"

    # $SqlCmd.CommandText = "SELECT LogonUserDisplayName,Operation,MailboxOwnerUPN,FolderPathName,ClientInfoString,LastAccessed FROM tbl_ExcMailboxAudit WHERE LastAccessed BETWEEN $($StartTicks) AND $($StopTicks)"
    $SqlCmd.CommandText = "SELECT LogonUserDisplayName,MailboxOwnerUPN,FolderPathName,LastAccessed FROM tbl_ExcMailboxAudit WHERE LastAccessed BETWEEN $($StartTicks) AND $($StopTicks)"

    $Data = $SqlCmd.ExecuteReader()
    $DataTable = new-object “System.Data.DataTable”
    $DataTable.Load($Data)
    $SqlCmd.CommandText = $NULL
    
    Return $DataTable
  }

  $DataTable = Get-YesterdayAuditReport($SqlCmd) | sort lastaccessed

  $SomeData = @()

  foreach ($thing in $DataTable)
  {
    $obj = New-Object PSObject 
    Add-Member -inputObject $obj -memberType NoteProperty -name "LogonUserDisplayName" -value $thing.LogonUserDisplayName 
    # Add-Member -inputObject $obj -memberType NoteProperty -name "Operation" -value $thing.Operation 
    Add-Member -inputObject $obj -memberType NoteProperty -name "MailboxOwnerUPN" -value $thing.MailboxOwnerUPN
    Add-Member -inputObject $obj -memberType NoteProperty -name "FolderPathName" -value $thing.FolderPathName
    # Add-Member -inputObject $obj -memberType NoteProperty -name "ClientInfoString" -value $thing.ClientInfoString
    Add-Member -inputObject $obj -memberType NoteProperty -name "LastAccessed" -value "$(Get-Date $([long]$($thing.LastAccessed)) -format G)"

    $SomeData += $obj
  }

#$SomeData = $SomeData | sort lastaccessed

  $Results = Project-YellowNikon -title "Microsoft Exchange" -subtitle "Mailbox Audit Log" -object $SomeData


#$StartDate = "$(get-date (Get-date).Addhours(-1) -Format G)"
#$AuditDetails = Search-MailboxAuditLog -LogonTypes Admin,Delegate -StartDate "$($StartDate)" -ResultSize 250000 -ShowDetails | ?{$_.logonuserdisplayname -notlike "*jrnl*"} | select LogonUserDisplayName,Operation,MailboxOwnerUPN,FolderPathName,ClientProcessName,LastAccessed

#Project-YellowNikon -title "Microsoft Exchange"
# Project-YellowNikon -title "Microsoft Exchange" -object $SomeData | out-file index-test.html;invoke-item index-test.html

### Need to do some logic on if no data is returned....


$YearMonthDayHourMinuteSec = get-date -f "yyyyddMM-HHMMss"
$File = "C:\Reports\Reports - MailboxAudit\MailboxAudit-$($YearMonthDayHourMinuteSec).html"

$Results | out-file -FilePath $($File)

$emailSmtpServer = "smtp.cbvm.com"
$emailSmtpServerPort = "25"

$StartTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addhours(-24).ticks)"
$StopTicks = "$(((Get-date).Date + (new-object system.timespan([math]::Round($dt.TimeofDay.TotalHours)),0,0)).addseconds(-1).ticks)"

$StartTimeFormated = "$(Get-date $([long]$StartTicks) -Format G)" + " - "
$StopTimeFormated = "$(Get-date $([long]$StopTicks) -Format `"HH:mm:ss`")" + " PM"
$FullTimeFormated = "$StartTimeFormated" + "$StopTimeFormated"

$emailMessage = New-Object System.Net.Mail.MailMessage
$emailMessage.From = "Postmaster <postmaster@cbvm.space>"
$emailMessage.To.Add( "bmarley22@gmail.com" )
$emailMessage.To.Add( "bmarley32@gmail.com" )
$emailMessage.To.Add( "bmarley43@gmail.com" )
$emailMessage.Subject = "Audit Log - $($FullTimeFormated)"
$emailMessage.Attachments.Add($File)
$emailMessage.IsBodyHtml = $true
$emailMessage.Body = @"
<h4> Please open attachment in IE or EDGE.</h4>
"@ 
$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
$SMTPClient.EnableSsl = $false
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
 
$SMTPClient.Send( $emailMessage )

CloseSQLConnection
}