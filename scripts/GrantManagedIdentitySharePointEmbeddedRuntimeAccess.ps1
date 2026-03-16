param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    [Parameter(Mandatory = $true)]
    [string]$OwningAppId,
    [Parameter(Mandatory = $true)]
    [string]$ManagedIdentityAppId,
    [Parameter(Mandatory = $true)]
    [string]$ManagedIdentityServicePrincipalId,
    [Parameter(Mandatory = $true)]
    [string]$ContainerTypeId,
    [Parameter(Mandatory = $true)]
    [string]$OwningAppClientSecret,
    [ValidateSet('read', 'full')]
    [string[]]$ApplicationPermissions = @('full')
)

$ErrorActionPreference = 'Stop'

function Get-GraphAccessToken {
    az account get-access-token --resource-type ms-graph --query accessToken --output tsv --only-show-errors
}

function Get-GraphServicePrincipal {
    param([string]$AppId)

    az ad sp show --id $AppId --output json --only-show-errors | ConvertFrom-Json
}

function Ensure-AppRoleAssignment {
    param(
        [string]$PrincipalId,
        [string]$ResourceId,
        [string]$AppRoleId,
        [string]$DisplayName
    )

    $accessToken = Get-GraphAccessToken
    $headers = @{ Authorization = "Bearer $accessToken" }
    $existing = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PrincipalId/appRoleAssignments" -Headers $headers
    $matchedAssignment = $existing.value | Where-Object { $_.appRoleId -eq $AppRoleId -and $_.resourceId -eq $ResourceId }
    if ($matchedAssignment) {
        Write-Host "App role already assigned: $DisplayName"
        return
    }

    $assignment = @{
        principalId = $PrincipalId
        resourceId = $ResourceId
        appRoleId = $AppRoleId
    } | ConvertTo-Json -Compress

    Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PrincipalId/appRoleAssignments" -Headers $headers -ContentType 'application/json' -Body $assignment | Out-Null
    Write-Host "Assigned app role: $DisplayName"
}

function Get-ClientCredentialToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = 'client_id=' + [uri]::EscapeDataString($ClientId) +
        '&client_secret=' + [uri]::EscapeDataString($ClientSecret) +
        '&scope=' + [uri]::EscapeDataString('https://graph.microsoft.com/.default') +
        '&grant_type=client_credentials'

    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -ContentType 'application/x-www-form-urlencoded' -Body $body
    $tokenResponse.access_token
}

function Invoke-GraphJson {
    param(
        [string]$Method,
        [string]$Uri,
        [string]$AccessToken,
        [object]$Body
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    if ($null -eq $Body) {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }

    Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType 'application/json' -Body ($Body | ConvertTo-Json -Depth 10 -Compress)
}

$graphServicePrincipal = Get-GraphServicePrincipal -AppId '00000003-0000-0000-c000-000000000000'
$graphRoleMap = @{
    'FileStorageContainer.Selected' = '40dc41bc-0f7e-42ff-89bd-d9516947e474'
    'Files.ReadWrite.All' = '75359482-378d-4052-8f01-80520e7db3cd'
}

Ensure-AppRoleAssignment -PrincipalId $ManagedIdentityServicePrincipalId -ResourceId $graphServicePrincipal.id -AppRoleId $graphRoleMap['FileStorageContainer.Selected'] -DisplayName 'FileStorageContainer.Selected'
Ensure-AppRoleAssignment -PrincipalId $ManagedIdentityServicePrincipalId -ResourceId $graphServicePrincipal.id -AppRoleId $graphRoleMap['Files.ReadWrite.All'] -DisplayName 'Files.ReadWrite.All'

$owningAppToken = Get-ClientCredentialToken -TenantId $TenantId -ClientId $OwningAppId -ClientSecret $OwningAppClientSecret

$registrationUri = "https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypeRegistrations/$ContainerTypeId"
Invoke-GraphJson -Method Put -Uri $registrationUri -AccessToken $owningAppToken -Body @{} | Out-Null
Write-Host "Registered container type in tenant: $ContainerTypeId"

$grantUri = "https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypeRegistrations/$ContainerTypeId/applicationPermissionGrants/$ManagedIdentityAppId"
$grantBody = @{
    applicationPermissions = $ApplicationPermissions
}
Invoke-GraphJson -Method Put -Uri $grantUri -AccessToken $owningAppToken -Body $grantBody | Out-Null
Write-Host "Granted managed identity app access to container type: $($ApplicationPermissions -join ', ')"

[pscustomobject]@{
    TenantId = $TenantId
    ManagedIdentityAppId = $ManagedIdentityAppId
    ManagedIdentityServicePrincipalId = $ManagedIdentityServicePrincipalId
    GraphApplicationPermissions = @('FileStorageContainer.Selected', 'Files.ReadWrite.All')
    ContainerTypeId = $ContainerTypeId
    ContainerTypeApplicationPermissions = $ApplicationPermissions
} | Format-List