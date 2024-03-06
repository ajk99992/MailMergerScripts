# Required Modules
#Install-Module -Name MSAL.PS -Scope CurrentUser -Force
Import-Module MSAL.PS

# Application (Client) ID, Tenant ID, and Client Secret
$clientId = "f53893b2-fb1d-41f5-ac6f-d68e90235012"
$tenantId = "ffec9c21-f95c-4502-aee3-591c93943a6f"
$clientSecret = "Acm8Q~5sDMGxPgz-QpsKozVSy4QQpz3mRjLPHcrI" # Securely store and access this

# Acquire Token
$scopes = @("https://graph.microsoft.com/.default")
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientSecret (ConvertTo-SecureString $clientSecret -AsPlainText -Force) -Scopes $scopes

#return $token

# Set Base URI for Microsoft Graph
$baseUri = "https://graph.microsoft.com/v1.0"

# Headers with Bearer Token
$headers = @{
    Authorization = "Bearer $($token.AccessToken)"
}

# Source and Target User Mailbox Email Addresses
$sourceUser = "AdeleV@7c40pt.onmicrosoft.com"
$targetUser = "MS365admin@7c40pt.onmicrosoft.com"

# Get Source User Mail Folders (to identify Folder ID for "Inbox" or any specific folder)
$sourceMailFoldersUri = "$baseUri/users/$sourceUser/mailFolders"

$sourceFoldersResponse = Invoke-RestMethod -Uri $sourceMailFoldersUri -Headers $headers -Method Get
return $sourceFoldersResponse
# Assuming you're copying from "Inbox". Find the Inbox folder ID
$inboxFolderId = ($sourceFoldersResponse.value | Where-Object { $_.displayName -eq "Inbox" }).id

# List Messages in Source Inbox
$sourceMessagesUri = "$baseUri/users/$sourceUser/mailFolders/$inboxFolderId/messages"
$sourceMessagesResponse = Invoke-RestMethod -Uri $sourceMessagesUri -Headers $headers -Method Get

# Loop through each message in the source inbox
foreach ($msg in $sourceMessagesResponse.value) {
    # For simplicity, this example only copies subject. You can extend this to include other message properties.
    $messageCopyBody = @{
        message = @{
            subject = $msg.subject
            # Add more properties as needed
        }
        # Specify the target folder in the target mailbox if necessary
    } | ConvertTo-Json

    # Create a new message in the target user's mailbox (default to "Drafts" folder)
    $createMessageUri = "$baseUri/users/$targetUser/messages"
    $null = Invoke-RestMethod -Uri $createMessageUri -Headers $headers -Method Post -Body $messageCopyBody -ContentType "application/json"
}

Write-Host "Messages copied successfully."
