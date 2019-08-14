$O365User = "andrew.brooke@suffolk.gov.uk" 
$SiteURL = "https://suffolknet-my.sharepoint.com/personal/andrew_brooke_suffolk_gov_uk";

$Folder = "C:\TEMP\SharePoint-Monitoring\SharePoint"  # Local location of fie to be transferred.
#DocDocLibName is document libary name 
$DocLibName = "Documents"
$foldername = "Test-Performance"
$DocLibName = "Documents"
$TodayDate = Get-Date -Format "yyyy-MM-dd"
$WorkPath = 'C:\Local\VisualStudio\office365-bau-checks'
$Folder = $WorkPath + "\SharePoint2"
$OutputFile = $WorkPath + "\Output\SharePoint_" + $TodayDate + ".csv"

#Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
#Add-Type -Path "C:\Temp\SharePoint-Monitoring\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8119.1200\lib\net45\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Temp\SharePoint-Monitoring\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8119.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "E:\Source\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8119.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Add-Type -Path "E:\Source\SPO_CSOM\Microsoft.SharePointOnline.CSOM.16.1.8119.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"

#Get Credentials from Hash file
#. 'C:\TEMP\SharePoint-Monitoring\eUser-LowerCredentials.ps1'
. 'C:\Local\PowerShell\Office365\eUser-LowerCredentials.ps1'


#Setup Proxy Settings
[system.net.webrequest]::DefaultWebProxy = new-object system.net.webproxy('proxy.eadidom.com:8080')
[system.net.webrequest]::Defaultwebproxy.BypassProxyOnLocal = $true

$webclient=New-Object System.Net.WebClient
$webclient.Proxy.Credentials = New-Object System.Net.NetworkCredential("euser\brooaj1", $O365Password)



#Import SharePoint module and bind to site collection
# > Import-Module -Name Microsoft.Online.SharePoint.PowerShell
[System.Reflection.Assembly]::LoadWithPartialName("System.IO.MemoryStream")
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($O365User,$O365Password)
$Context.Credentials = $Creds


#Retrieve list
$List = $Context.Web.Lists.GetByTitle("$DocLibName")
$Context.Load($List)
$Context.Load($List.RootFolder)
$Context.ExecuteQuery()
$ServerRelativeUrlOfRootFolder = $List.RootFolder.ServerRelativeUrl
$uploadFolderUrl=  $ServerRelativeUrlOfRootFolder+"/"+$foldername




#Upload file
Foreach ($File in (dir $Folder -File))
{
$FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $File
 if($foldername -eq $null)
  {
  $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
  }
  Else
  {
   $targetFolder = $Context.Web.GetFolderByServerRelativeUrl($uploadFolderUrl)
   $Upload = $targetFolder.Files.Add($FileCreationInfo);
  }
#$Upload = $List.RootFolder.Files.Add($FileCreationInfo)
$Context.Load($Upload)
$CommandTime = Measure-Command {$Context.ExecuteQuery()}
}
#Export Timing
$TodayTime = Get-Date -Format "HH:mm:ss"
$LineOut = $TodayDate + "," + $TodayTime + "," + $CommandTime.TotalSeconds
$LineOut | Out-File $OutputFile -Encoding ascii -Append


#################################
# Store results in  SharePoint List.
$siteListUrl = "https://suffolknet.sharepoint.com/sites/ITProblemManagement"
$listName = "SharePoint"

#$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteListUrl)
#$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($O365User,$O365Password)
$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteListUrl)
$context.Credentials = $creds
[Microsoft.SharePoint.Client.Web]$web = $context.Web
[Microsoft.SharePoint.Client.List]$list = $web.Lists.GetByTitle($listName)
 
$ListItems = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$Context.Load($ListItems)
$Context.ExecuteQuery()      
 
#write-host "Total Number of List Items found to Delete:"$ListItems.Count
# 
#    if ($ListItems.Count -gt 0)
#    {
#        #Loop through each item and delete
#        For ($i = $ListItems.Count-1; $i -ge 0; $i--)
#        {
#            $ListItems[$i].DeleteObject()
#        }
#        $Context.ExecuteQuery()
#         
#        Write-Host "All Existing List Items deleted Successfully!"
#    }

    $today = Get-Date -Format g

    [Microsoft.SharePoint.Client.ListItemCreationInformation]$itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation;
    [Microsoft.SharePoint.Client.ListItem]$item = $list.AddItem($itemCreateInfo);
    $item["Title"] = $today
    #$item["Date_x002d_Time"] = $today
    $item["Uploadtime"] = $CommandTime.TotalSeconds 
   
    ############################
    # More columns here...
    ############################
    $item.Update();
    $context.ExecuteQuery(); 