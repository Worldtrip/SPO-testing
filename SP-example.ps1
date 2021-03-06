#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

Function Get-SP-Credentials () # Looks up credentails stored in the KeyStore location and retuns a PSCredential Object
    {
	param (
		[Parameter(Mandatory=$true)] [string] $User,
		[Parameter(Mandatory=$true)] [string] $KeyStore
        )
	
	
	# Set AES Key to encrypt password, this could be a generated key and stored somewhere more restrictive.
	
	$KeyCrypt = Get-Content "AES.key"

	$O365User = (Get-ADUser $User -Properties UserPrincipalName | Select-Object UserPrincipalName).UserPrincipalName 

	$O365PassFile = "$KeyStore$O365User.txt"
 

	# Prompt if passfile not present
	if (!(Test-Path -Path $O365PassFile)) {Read-Host -Prompt "Please enter password for $O365User :" -AsSecureString | ConvertFrom-SecureString -Key $KeyCrypt | Out-File $O365PassFile}
	$O365Password = Get-Content $O365PassFile | ConvertTo-SecureString -ErrorAction Stop -Key $KeyCrypt
	return New-Object System.Management.Automation.PSCredential ($O365User, $O365Password)
    
}

Function UploadSPO () # Uploads a file to SharePoint Online
    {
    param (
        [Parameter(Mandatory=$true)] [SecureString] $O365Credentials,
		[Parameter(Mandatory=$true)] [string] $Site,
        [Parameter(Mandatory=$true)] [string] $Library,
        [Parameter(Mandatory=$true)] [string] $File
        )
    #
    $user = Get-ADUser -Filter {EmailAddress -eq "andrew.brooke@suffolk.gov.uk"} -Properties CanonicalName | Select-Object CanonicalName


    #Setup to use Proxy
    [system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('proxy.eadidom.com:8080')
    [system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true
    
    #[system.net.webrequest]::defaultwebproxy.credentials = $O365Credentials
   
    $webclient=New-Object System.Net.WebClient
    $webclient.Proxy.Credentials = New-Object System.Net.NetworkCredential($user,$O365Credentials.Password)
   
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Site)
    $Context.Credentials = $O365Credentials

    #Retrieve library details
    $List = $Context.Web.Lists.GetByTitle($Library)
    $Context.Load($List)
    $Context.ExecuteQuery()

    # Setup ready for file transfer
    $FileStream = New-Object IO.FileStream($File,[System.IO.FileMode]::Open)
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $File
    $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
    $Context.Load($Upload)
    $Context.ExecuteQuery()
    

    }
