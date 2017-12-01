function Get-SSHBackup
{
  <#
  .SYNOPSIS
  Backs up CPS via SSH
  .DESCRIPTION
  The Get-SSHBackup function uses the Putty Suite (PLINK and PSCP) to run a backup for CPS.
  It first runs PLINK to send a tar command to backup everything up into a *tar.gz file. It
  then uses PSCP to SSH the backup to the specified Output directory specified.
  .LINK
  http://www.KeithSmithOnline.com
  https://twitter.com/@SevenLayerJedi
  .PARAMETER UserName
  A single user name used to send the commands.
  .PARAMETER Hostame
  A single Hostname that you want to connect to.
  .PARAMETER Domain
  The Domain the Host is apart of
  .PARAMETER OutputDir
  The Directory you would like to save the backup to
  .PARAMETER PrivateKey
  The location of the private key. The private key must be in the correct putty format. used
  PuttyGen to create the private key.
  .EXAMPLE
  Backup the files on a host to c:\temp
  Get-SSHBackup -UserName appadmin -Hostame "SERVER1" -Domain "cbvm.space" -OutputDir "C:\Temp" -PrivateKey "c:\sshkeys\ssh.ppk"
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

  param (
	[parameter(Mandatory=$true)]
  [String]
  $UserName,
  
  [parameter(Mandatory=$true)]
  [String]
  $Hostame,
  
  [parameter(Mandatory=$true)]
  [String]
  $Domain,
  
  [parameter(Mandatory=$true)]
  [String]
  $OutputDir,
  
  [parameter(Mandatory=$true)]
  [String]
  $PrivateKey 
  )

  # Set Date Time in Ticks
  $Ticks = (get-date).ticks

  # Set the FQDN
  $varFQDN = $($Hostame) + "." + $($Domain)
  
  # User @ FQDN
  # bmarley@contoso.local
  $varUserAtFQDN = $($UserName) + "@" + $($varFQDN)

  # Setup TAR Command to run with plink
  $CMDToRun = "tar cvzf /tmp/`"$($Ticks)_$($Hostame).tar.gz`" --exclude=`'/apps/tomcat7/logs/`*`' --exclude=`'/apps/tomcat7/temp/`*`' /apps/tomcat7 /apps/certs /apps/keystore"
  $CMDToRun = $CMDToRun.replace("`"","`'") #`'
  
  # Send Plink command to backup everything 
  $CMDplink = "plink -ssh -i `"$($PrivateKey)`" `"$($varUserAtFQDN)`" `"$($CMDToRun)`""
  invoke-expression $CMDplink

  # Download backup via ssh
  $CMDpscp = "pscp -i `"$($PrivateKey)`" `"$($varUserAtFQDN):/tmp/$($Ticks)_$($Hostame).tar.gz`" `"$($OutputDir)`""
  invoke-expression $CMDpscp
}
