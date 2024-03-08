# Import the MSAL.PS module
Import-Module MSAL.PS

# Load the EWS Managed API DLL
Import-Module "C:\Mailmigration\Microsoft.Exchange.WebServices.dll"

# Set Exchange service
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
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

# Define email addresses for source and target mailboxes
$sourceMailboxEmail = "AdeleV@7c40pt.onmicrosoft.com"
$targetMailboxEmail = "MS365admin@7c40pt.onmicrosoft.com"

# Define impersonation for the specific mailbox you want to access
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    $sourceMailboxEmail
)

# Bind to the source Inbox and Sent Items folders
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$sentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems)

# Define a function to copy items from one folder to another
function CopyFolderItems($sourceFolder, $targetFolder) {
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    # Exclude the "Hashtags" property from the view
    $itemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $itemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Id)
    $itemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $itemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
    try {
        $findResults = $sourceFolder.FindItems($itemView)
        foreach ($item in $findResults.Items) {
            # Copy the item to the target folder
            $item.Copy($targetFolder.Id)
        }
    } catch {
        Write-Host "An error occurred while copying items: $_"
    }
}

# Impersonate the target mailbox to find the target folders' IDs
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    $targetMailboxEmail
)

# Bind to the target Inbox and Sent Items folders to get their IDs
$targetInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$targetSentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems)

# Copy items from source to target folders
CopyFolderItems $inbox $targetInbox
CopyFolderItems $sentItems $targetSentItems

Write-Host "Folders copied successfully."
