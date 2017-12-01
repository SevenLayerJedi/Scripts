function Send-MailMessage
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
  Send email through smtp.contoso.local on port 25. Send the email from Postmaster@contoso.local to bmarley@contoso.local with a subject and body of "Test Email"
  Send-MailMessage -SmtpServer "smtp.contoso.local" -Port "25" -From "Postmaster@contoso.local" -To "bmarley@contoso.local" -Subject "Test Email" -Body "Test Email"
  .NOTES
  Life is really simple, but we insist on making it complicated.
  #>

  param (
	[parameter(Mandatory=$true)]
  [String]
  $SmtpServer,
  
  [parameter(Mandatory=$true)]
  [String]
  $Port,

  [parameter(Mandatory=$true)]
  [String]
  $From,
  
  [parameter(Mandatory=$true)]
  [String]
  $To,
  
  [parameter(Mandatory=$true)]
  [String]
  $Subject,
  
  [parameter(Mandatory=$true)]
  [String]
  $Body
  )

  $emailSmtpServer = "$($SmtpServer)"
  $emailSmtpServerPort = "$($Port)"
   
  $emailMessage = New-Object System.Net.Mail.MailMessage
  $emailMessage.From = "$($From)"
  $emailMessage.To.Add( "$($To)" )
  $emailMessage.Subject = "$($Subject)"
  $emailMessage.IsBodyHtml = $true
  $emailMessage.Body = "$($Body)"

  $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
  $SMTPClient.EnableSsl = $true

  $SMTPClient.Send( $emailMessage )
}
