function Get-ServerThatisLockingUserOut
{
  <#
  .SYNOPSIS
  Update
  .DESCRIPTION
  Update
  .LINK
  http://www.KeithSmithOnline.com
  https://twitter.com/@SevenLayerJedi
  .PARAMETER Identity
  Username of user
  .EXAMPLE
  Get computers that are locking out bmarley
  Get-ServerThatisLockingUserOut -Identity bmarley
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

  param (
	[parameter(Mandatory=$true)]
  [String]
  $Identity
  )

  $Pdce = (Get-AdDomain).PDCEmulator
  $GweParams = @{
       ‘Computername’ = $Pdce
       ‘LogName’ = ‘Security’
       ‘FilterXPath’ = "*[System[EventID=4740] and EventData[Data[@Name='TargetUserName']='$($Identity)']]"
  }
  $Events = Get-WinEvent @GweParams
  $Results = $Events | fl | findstr Call
  $Results = $($Results.split(":"))[1].trim()
  RETURN $Results
}
