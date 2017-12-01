
function Get-SPListItemByID
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
  Get-SPListItems -Site "http://share.contoso.local/it/team" -List "Shared Documents" -ID "27"
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
  $List,
  
  param (
	[parameter(Mandatory=$true)]
  [String]
  $ID
  )

  [string]$source = "$($Site)"
  $spSourceWeb = Get-SPWeb $source
  $list = $spSourceWeb.Lists["$($List)"]

  $AllItems = $list.GetItems()
  $Item = $AllItems.GetItemById($($ID))
  RETURN $Item
}
