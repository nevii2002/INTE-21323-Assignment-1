# EPMS Azure Setup Script
# Run this once in a NEW PowerShell terminal after Azure CLI is installed
# Usage: .\azure-setup.ps1

$RESOURCE_GROUP = "epms-rg"
$LOCATION = "southeastasia"
$BACKEND_APP_NAME = "epms-backend"
$STORAGE_ACCOUNT = "epmsstorage$(Get-Random -Maximum 9999)"
$FILE_SHARE = "epmsdata"
$APP_SERVICE_PLAN = "epms-plan"
$STATIC_APP_NAME = "epms-frontend"
$GITHUB_REPO = "https://github.com/nevii2002/INTE-21323-Assignment-1"

Write-Host "=== EPMS Azure Setup ===" -ForegroundColor Cyan

# Login
Write-Host "`n[1/8] Logging in to Azure..." -ForegroundColor Yellow
az login

# Create Resource Group
Write-Host "`n[2/8] Creating resource group '$RESOURCE_GROUP'..." -ForegroundColor Yellow
az group create --name $RESOURCE_GROUP --location $LOCATION

# Create App Service Plan (Free tier)
Write-Host "`n[3/8] Creating App Service Plan..." -ForegroundColor Yellow
az appservice plan create --name $APP_SERVICE_PLAN --resource-group $RESOURCE_GROUP --sku B1 --is-linux

# Create Backend Web App (Node 18)
Write-Host "`n[4/8] Creating backend Web App '$BACKEND_APP_NAME'..." -ForegroundColor Yellow
az webapp create --name $BACKEND_APP_NAME --resource-group $RESOURCE_GROUP --plan $APP_SERVICE_PLAN --runtime "NODE:18-lts"

# Set backend environment variables
Write-Host "`n[5/8] Configuring backend environment variables..." -ForegroundColor Yellow
$JWT_SECRET = [System.Web.Security.Membership]::GeneratePassword(32, 4)
az webapp config appsettings set --name $BACKEND_APP_NAME --resource-group $RESOURCE_GROUP --settings `
  JWT_SECRET="$JWT_SECRET" `
  NODE_ENV="production" `
  SCM_DO_BUILD_DURING_DEPLOYMENT="false" `
  WEBSITE_RUN_FROM_PACKAGE="1"

# Create Storage Account for SQLite persistence
Write-Host "`n[6/8] Creating storage for SQLite persistence..." -ForegroundColor Yellow
az storage account create --name $STORAGE_ACCOUNT --resource-group $RESOURCE_GROUP --location $LOCATION --sku Standard_LRS
$STORAGE_KEY = az storage account keys list --account-name $STORAGE_ACCOUNT --resource-group $RESOURCE_GROUP --query "[0].value" -o tsv
az storage share create --name $FILE_SHARE --account-name $STORAGE_ACCOUNT --account-key $STORAGE_KEY

# Mount file share to Web App
az webapp config storage-account add `
  --name $BACKEND_APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --custom-id epms-data `
  --storage-type AzureFiles `
  --account-name $STORAGE_ACCOUNT `
  --share-name $FILE_SHARE `
  --access-key $STORAGE_KEY `
  --mount-path /home/data

az webapp config appsettings set --name $BACKEND_APP_NAME --resource-group $RESOURCE_GROUP --settings SQLITE_PATH="/home/data/database.sqlite"

# Create Static Web App for frontend
Write-Host "`n[7/8] Creating Static Web App '$STATIC_APP_NAME'..." -ForegroundColor Yellow
az staticwebapp create --name $STATIC_APP_NAME --resource-group $RESOURCE_GROUP --location "eastasia" --source $GITHUB_REPO --branch master --app-location "/client" --output-location "dist" --login-with-github

# Get deployment credentials
Write-Host "`n[8/8] Getting GitHub secrets..." -ForegroundColor Yellow
$PUBLISH_PROFILE = az webapp deployment list-publishing-profiles --name $BACKEND_APP_NAME --resource-group $RESOURCE_GROUP --xml
$SWA_TOKEN = az staticwebapp secrets list --name $STATIC_APP_NAME --resource-group $RESOURCE_GROUP --query "properties.apiKey" -o tsv

Write-Host "`n=== DONE! Add these as GitHub repository secrets ===" -ForegroundColor Green
Write-Host "`nGo to: https://github.com/nevii2002/INTE-21323-Assignment-1/settings/secrets/actions" -ForegroundColor Cyan
Write-Host "`nSecret name: AZURE_WEBAPP_PUBLISH_PROFILE" -ForegroundColor White
Write-Host "Value:" -ForegroundColor White
Write-Host $PUBLISH_PROFILE -ForegroundColor Gray
Write-Host "`nSecret name: AZURE_STATIC_WEB_APPS_API_TOKEN" -ForegroundColor White
Write-Host "Value: $SWA_TOKEN" -ForegroundColor Gray

Write-Host "`n=== Backend URL ===" -ForegroundColor Cyan
Write-Host "https://$BACKEND_APP_NAME.azurewebsites.net" -ForegroundColor Green
Write-Host "`n=== Frontend URL ===" -ForegroundColor Cyan
az staticwebapp show --name $STATIC_APP_NAME --resource-group $RESOURCE_GROUP --query "defaultHostname" -o tsv
