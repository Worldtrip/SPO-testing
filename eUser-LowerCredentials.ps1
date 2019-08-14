# Preset Variables. Adjust Passfile as required
$KeyCrypt = (120,80,177,104,51,125,207,56,9,193,73,130,194,179,251,82,35,70,169,109,92,180,55,125,88,209,58,166,75,92,211,116)
$O365User = 'andrew.brooke@suffolk.gov.uk'
$O365PassFile = "C:\TEMP\SharePoint-Monitoring\Keys\$O365User.txt"

# Prompt if passfile not present
if (!(Test-Path -Path $O365PassFile)) {Read-Host -Prompt "Please enter password for $O365User :" -AsSecureString | ConvertFrom-SecureString -Key $KeyCrypt | Out-File $O365PassFile}
$O365Password = Get-Content $O365PassFile | ConvertTo-SecureString -ErrorAction Stop -Key $KeyCrypt
$O365Credentials = New-Object System.Management.Automation.PSCredential ($O365User, $O365Password)
