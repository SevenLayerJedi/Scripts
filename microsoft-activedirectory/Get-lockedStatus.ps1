function Get-LockedStatus
{
  <#
  .SYNOPSIS
  Update
  .DESCRIPTION
  Update
  .LINK
  http://www.KeithSmithOnline.com
  https://twitter.com/@SevenLayerJedi
  .PARAMETER Null
  Update
  .EXAMPLE
  Get locked status for user bmarley
  Get-LockedStatus -Identity bmarley
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

  param (
	[parameter(Mandatory=$true)]
  [String]
  $Identity
  )

  $Status = get-aduser -Identity $($Identity) -Properties lockedout
  RETURN $STATUS
}
