function Get-SPListItems
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
  Update
  Get-SPListItems -Site "http://share.contoso.local/it/team" -List "Shared Documents"
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

  param (
	[parameter(Mandatory=$true)]
  [String]
  $Site,
  
  param (
	[parameter(Mandatory=$true)]
  [String]
  $List
  )

  [string]$source = "$($Site)"
  $spSourceWeb = Get-SPWeb $source
  $list = $spSourceWeb.Lists["$($List)"]

  $AllItems = $list.GetItems()
  RETURN $AllItems
}
