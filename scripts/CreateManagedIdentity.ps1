$env = "dev"

# Connect to Azure
az login

# Get the access token for Dynamics
$accessToken = (az account get-access-token --resource "https://$env.crm6.dynamics.com" --query accessToken -o tsv)


$requestBody = @{
    "applicationid" = ""
    "managedidentityid" = ""
    "credentialsource" = 2
    "subjectscope" = 1
    "tenantid" = ""
} | ConvertTo-Json

# Call the web API to associate the user with the field level security profile
$response = Invoke-RestMethod -Uri "https://$env.crm6.dynamics.com/api/data/v9.0/managedidentities" `
    -Method POST `
    -Body $requestBody `
    -Headers @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}