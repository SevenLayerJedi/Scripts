function Get-ParamCount
{
  <#
  .SYNOPSIS
  Gets the amount of parameters and then does something based on that information
  .DESCRIPTION
  Get the parameter count. If the count is greater than 1, write verbose the information
  to the screen and stop the script. If count is equal to 1 then say something positive.
  .LINK
  http://www.KeithSmithOnline.com
  https://twitter.com/@SevenLayerJedi
  .PARAMETER Null
  Stuff
  .EXAMPLE
  Pass 4 Parameters to the function
  Get-ParamCount -Stuff "Bob", "Billy", "Bo", "Bon Jovi"
  .EXAMPLE
  Pass 1 Parameters to the function
  Get-ParamCount -Stuff "Bob"
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>
  
  param
  ( 
    [Parameter(Mandatory=$False)] 
    [array]$Stuff   
  )
  
  #Old Setting
  $oldverbose = $VerbosePreference
  
  # New Setting
  $VerbosePreference = "continue"

  
  if ($($Stuff.count) -gt 1)
  {
    Write-Verbose "You can only pass 1 param dooood!"
    $VerbosePreference = $oldverbose
    BREAK
  }
  else
  {
    Write-host "Woo!, you only passed 1 param! Winner!"
  }
  
  # Change back to old setting
  $VerbosePreference = $oldverbose
}
Get-Help Get-ParamCount
