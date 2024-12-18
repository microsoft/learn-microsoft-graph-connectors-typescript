#!/usr/bin/env bash

# login
echo "Sign in to Microsoft 365..."
npx -p @pnp/cli-microsoft365 -- m365 login --authType browser --appId 14d82eec-204b-4c2f-b7e8-296a70dab67e

open "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e&response_type=code&scope=https%3A%2F%2Fgraph.microsoft.com%2FApplication.ReadWrite.All%20https%3A%2F%2Fgraph.microsoft.com%2FAppRoleAssignment.ReadWrite.All"

read -p "Press enter to continue when ready"

echo "Waiting 10s for permissions to propagate..."
sleep 10

# create Entra app
echo "Creating Entra app..."
appInfo=$(npx -p @pnp/cli-microsoft365 -- m365 entra app add --name "Microsoft Graph documentation" --withSecret --apisApplication "https://graph.microsoft.com/ExternalConnection.ReadWrite.OwnedBy, https://graph.microsoft.com/ExternalItem.ReadWrite.OwnedBy" --grantAdminConsent --output json)

# write app to env.ts
echo "Writing app to src/env.ts..."
echo "export const appInfo = $appInfo;" > src/env.ts

echo "DONE"