
function LoadExchangePSModule
{
  # Load Exchange Powershell module
  add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
}
LoadExchangePSModule

## Define UPN of the Account that has impersonation rights
$AccountWithImpersonationRights = "bmarley@cbvm.space"

##Define the SMTP Address of the mailbox to impersonate
$MailboxToImpersonate = "keith.smith@cbvm.space"

## Load Exchange web services DLL
## Download here if not present: http://go.microsoft.com/fwlink/?LinkId=255472
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

Import-Module $dllpath

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

## Create Exchange ExchangeService Object
$ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 

#Get valid Credentials using UPN for the ID that is used to impersonate mailbox
#$psCred = Get-Credential 
$user = 'svc_bmarley'
$pass = '123456789' | ConvertTo-SecureString -AsPlainText -Force

#$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString()) 
$creds = New-Object System.Net.NetworkCredential($User,$pass) 
$ExchangeService.Credentials = $creds 

## Set the URL of the CAS (Client Access Server)
$ExchangeService.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})

##Login to Mailbox with Impersonation

Write-Host 'Using ' $AccountWithImpersonationRights ' to Impersonate ' $MailboxToImpersonate

$ExchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate ); 

#Connect to the Inbox and display basic statistics

$InboxFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName) 
$SentItemsFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$ImpersonatedMailboxName) 
$DeletedItemsFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems,$ImpersonatedMailboxName) 
$DraftsFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts,$ImpersonatedMailboxName) 
$OutboxFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Outbox,$ImpersonatedMailboxName) 
$RecoverableItemsRootFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot,$ImpersonatedMailboxName) 
$RecoverableItemsDeletionsFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsDeletions,$ImpersonatedMailboxName) 
$RecoverableItemsPurgesFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsPurges,$ImpersonatedMailboxName) 

$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$InboxFolder)
$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$SentItemsFolder)
$DeletedItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$DeletedItemsFolder)
$Drafts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$DraftsFolder)
$Outbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$RecoverableItemsRootFolder)
$RecoverableItemsRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$RecoverableItemsDeletionsFolder)
$RecoverableItemsDeletions = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$RecoverableItemsDeletionsFolder)
$RecoverableItemsPurges = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$RecoverableItemsPurgesFolder)

# ItemView Class - Represents the view settings in a folder search operation
# Initializes a new instance of the ItemView class by using the supplied page size of 1
$ItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(2)  

# Returns a list of items by searching the contents of the Inbox folder. 
$FolderFindResults = $Inbox.finditems($ItemView)

$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

$ExchangeService.LoadPropertiesForItems($FolderFindResults,$PropertySet) 


foreach($Item in $FolderFindResults.Items){  
  "RecivedDate : " + $Item.DateTimeReceived   
  "Subject     : " + $Item.Subject   
  "Size        : " + $Item.Size  
  "Recipients  : " + $Item.ToRecipients
  "Body        : " + $Item.Body
  "Headers     : " + $Item.InternetMessageHeaders
}  

# Load the Content of the attachments
$FolderFindResults.Items[0].Attachments.load()

$AttachmentContents = $FolderFindResults.Items[0].Attachments[0].Content
[System.Text.Encoding]::ASCII.GetString($AttachmentContents)


$xl_Workbook_xml = New-Object System.Xml.XmlDocument 



$file = [io.file]::ReadAllBytes('c:\scripts\test.jpg')

[io.file]::WriteAllBytes('c:\temp\temp.xlsx',$AttachmentContents)


$FolderFindResults.Items[0].InternetMessageHeaders


Get-ChildItem .\temp\ -Recurse | Get-Content | findstr /i blu3


$zipfilename = "C:\Toolshed\Software\Strings\temp.zip"
$destination = "C:\Toolshed\Software\Strings\temp2"

$shellApplication = new-object -com shell.application 
$zipPackage = $shellApplication.NameSpace($zipfilename) 
$destinationFolder = $shellApplication.NameSpace($destination) 
$destinationFolder.CopyHere($zipPackage.Items()) 

$content = [IO.File]::ReadAllText("$zipPackage")




[Void][Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')  


$ZipContents = foreach ($thing in "C:\Toolshed\Software\Strings\temp.zip"){[IO.Compression.ZipFile]::OpenRead($thing).Entries}
$SpecificFile = $ZipContents[4]


function Extract-Zip 
{ 
    param([string]$zipfilename, [string] $destination) 
    if(test-path($zipfilename)) 
    { 
        $shellApplication = new-object -com shell.application 
        $zipPackage = $shellApplication.NameSpace($zipfilename) 
        $destinationFolder = $shellApplication.NameSpace($destination) 
        $destinationFolder.CopyHere($zipPackage.Items()) 
    }
    else 
    {   
        Write-Host $zipfilename "not found"
    }
}

Get-ChildItem .\temp\ -Recurse | Get-Content | findstr /i blu3


















$view =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(10)  
$findResults = $ExchangeService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$view)





foreach($Item in $findResults.Items){  
    "RecivedDate : " + $Item.DateTimeReceived   
    "Subject     : " + $Item.Subject   
    "Size        : " + $Item.Size  
    "Recipients  : " + $Item.ToRecipients  
}  






$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet
([Microsoft.Exchange.WebServices.Data.BasePropertySet]::BasePropertySet.FirstClassProperties)






$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(4)
$view.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::BasePropertySet.FirstClassProperties)
$view.PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::BasePropertySet.FirstClassProperties)
$PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$findResults = $ExchangeService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$view)


$ExchangeService.LoadPropertiesForItems($findResults, $view);




$message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($event.MessageData,$itmId)

$PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text



$message.Load($PropertySet)
$bodyText= $message.Body.toString()







#$view.SearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$findResults = $service.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$view)



$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)


PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
itempropertyset.RequestedBodyType = BodyType.Text;
ItemView itemview = new ItemView(1000);
itemview.PropertySet = itempropertyset;

FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, "subject:TODO", itemview);
Item item = findResults.FirstOrDefault();
item.Load(itempropertyset);
Console.WriteLine(item.Body);





#Number of Items to Get
$pageSize =50
$Offset = 0
$MoreItems =$true
$ItemCount=5
$ItemSize=0




$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView($pageSize,$Offset,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
#$itemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow

$oSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, ([System.Datetime]::now).Adddays(-7))


$service.finditems()


var items = serviceInstance.FindItems( 
       //Find Mails from Inbox of the given Mailbox 
       new FolderId(WellKnownFolderName.Inbox, new Mailbox(ConfigurationManager.AppSettings["MailBox"].ToString())), 
       //Filter criterion 
       new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter[] {         new SearchFilter.ContainsSubstring(ItemSchema.Subject, ConfigurationManager.AppSettings["ValidEmailIdentifier"].ToString()), 
       new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false) 
       }), 
       //View Size as 15 
       new ItemView(15)); 
 





Write-Host 'Total Item count for Inbox:' $Inbox.TotalCount
Write-Host 'Total Items Unread:' $Inbox.UnreadCount 

$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$SentItemsFolder)
Write-Host 'Total Item count for SentItems:' $SentItems.TotalCount
Write-Host 'Total Items Unread:' $SentItems.UnreadCount 








$CalendarFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$ImpersonatedMailboxName) 

$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$CalendarFolder)

$TotalMinutes = 0

For ($i=31; $i -ne 0; $i--)
{
  $t = 0
  $cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView([DateTime]::Today.AddDays(-$i).AddHours(24),[DateTime]::Today.AddDays(-$($i - 2)).AddSeconds(-1),2000)
  $cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

  $CalendarEvents = $Calendar.FindAppointments($cvCalendarview)

  Write-host "[$(Get-date ([DateTime]::Today.AddDays(-$i).AddHours(24)) -Format D)]" -foregroundcolor green
  
  #$CalendarEvents.Length
  #$CalendarEvents.count
  
  if ($CalendarEvents)
  {
    $t = 0
    foreach($Event in $CalendarEvents)
    {
      Write-host "`t[$(Get-date ([long]$Event.Start.Ticks) -Format t) - $(Get-date ([long]$Event.End.Ticks) -Format t)] $($Event.Location) - $($Event.Subject)"
      $t += ($(Get-date ([long]$Event.End.Ticks)) - $(Get-date ([long]$Event.Start.Ticks))).totalminutes
      $TotalMinutes += $t
    }
    
    Write-Host "`t`t[Total Minutes] $($t)" -foregroundcolor cyan
  }
  else
  {
    # Nothing
  }
}
Write-host " [Total Hours in Meetings] $((New-TimeSpan -Minutes ($TotalMinutes.ToString())).totalhours)" -foregroundcolor magenta

