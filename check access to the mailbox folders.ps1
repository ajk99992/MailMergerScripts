# Import the MSAL.PS module
Import-Module MSAL.PS

# Load the EWS Managed API DLL
Import-Module "C:\mailmerger\PowerShell-EWS-Scripts\Legacy\Microsoft.Exchange.WebServices.dll"

# Set Exchange service
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$service.Url = New-Object System.Uri("https://outlook.office365.com/EWS/Exchange.asmx")

# Set credentials for OAuth
$clientId = "f53893b2-fb1d-41f5-ac6f-d68e90235012"
$clientSecretPlainText = "Acm8Q~5sDMGxPgz-QpsKozVSy4QQpz3mRjLPHcrI" # Ensure this is kept secure
$clientSecret = ConvertTo-SecureString $clientSecretPlainText -AsPlainText -Force
$tenantId = "ffec9c21-f95c-4502-aee3-591c93943a6f"

# Acquire an access token using MSAL.PS
$tokenResult = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret -Scopes "https://outlook.office365.com/.default"

# Set the credentials with the access token for the Exchange service
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($tokenResult.AccessToken)

# Define impersonation for the specific mailbox you want to access
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    "AdeleV@7c40pt.onmicrosoft.com"
)

# Bind to the mailbox's root folder
$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, "AdeleV@7c40pt.onmicrosoft.com")
try {
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $rootFolderId)
} catch {
    Write-Host "Error encountered during Bind operation: $_"
    exit
}

# List the mailbox's folders
function ListFolders($folder) {
    Write-Host ("Folder: " + $folder.DisplayName)
    # Get the child folders
    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    try {
        $folders = $folder.FindFolders($folderView)
        foreach ($childFolder in $folders.Folders) {
            ListFolders $childFolder
        }
    } catch {
        Write-Host "An error occurred: $_"
    }
}

# Start listing from the root folder
ListFolders $rootFolder
