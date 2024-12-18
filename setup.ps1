# login
Write-Host "Sign in to Microsoft 365..."
npx -p @pnp/cli-microsoft365 -- m365 login --authType browser --appId 14d82eec-204b-4c2f-b7e8-296a70dab67e

Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e&response_type=code&scope=https%3A%2F%2Fgraph.microsoft.com%2FApplication.ReadWrite.All%20https%3A%2F%2Fgraph.microsoft.com%2FAppRoleAssignment.ReadWrite.All"

Read-Host -Prompt "Press Enter to continue when ready"

Write-Output "Waiting 10s for permissions to propagate..."
Start-Sleep -Seconds 10

# create AAD app
Write-Host "Creating AAD app..."
$appInfo=$(npx -p @pnp/cli-microsoft365 -- m365 aad app add --name "Microsoft Graph documentation" --withSecret --apisApplication "https://graph.microsoft.com/ExternalConnection.ReadWrite.OwnedBy, https://graph.microsoft.com/ExternalItem.ReadWrite.OwnedBy" --grantAdminConsent --output json)

# write app to env.ts
Write-Host "Writing app to src/env.ts..."
New-Item -ItemType File -Name "src/env.ts" -Value "export const appInfo = $($appInfo | Out-String)" -Force

Write-Host "DONE"
