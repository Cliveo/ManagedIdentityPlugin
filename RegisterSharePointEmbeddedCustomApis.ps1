param(
    [Parameter(Mandatory = $true)]
    [string]$OrgUrl,
    [switch]$ForceRecreate
)

$ErrorActionPreference = 'Stop'

function Get-DataverseToken {
    return az account get-access-token --resource $OrgUrl --query accessToken -o tsv
}

function Get-Headers {
    param([string]$Token)

    return @{
        Authorization = "Bearer $Token"
        Accept = 'application/json'
        'OData-Version' = '4.0'
        'OData-MaxVersion' = '4.0'
    }
}

function Invoke-DataverseGet {
    param(
        [string]$Path,
        [string]$Token
    )

    $uri = "$OrgUrl/api/data/v9.2/$Path"
    Invoke-RestMethod -Method Get -Uri $uri -Headers (Get-Headers -Token $Token)
}

function Invoke-DataversePatch {
    param(
        [string]$Path,
        [string]$Token,
        [hashtable]$Body
    )

    Invoke-RestMethod -Method Patch -Uri "$OrgUrl/api/data/v9.2/$Path" -Headers (Get-Headers -Token $Token) -ContentType 'application/json' -Body ($Body | ConvertTo-Json -Depth 10)
}

function Invoke-DataversePost {
    param(
        [string]$EntitySet,
        [string]$Token,
        [hashtable]$Body
    )

    $headers = Get-Headers -Token $Token
    $headers['Prefer'] = 'return=representation'
    Invoke-RestMethod -Method Post -Uri "$OrgUrl/api/data/v9.2/$EntitySet" -Headers $headers -ContentType 'application/json' -Body ($Body | ConvertTo-Json -Depth 10)
}

function Invoke-DataverseDelete {
    param(
        [string]$Path,
        [string]$Token
    )

    Invoke-RestMethod -Method Delete -Uri "$OrgUrl/api/data/v9.2/$Path" -Headers (Get-Headers -Token $Token)
}

function Get-CustomApiChildren {
    param(
        [string]$EntitySet,
        [string]$PrimaryId,
        [string]$CustomApiId,
        [string]$Token
    )

    $result = Invoke-DataverseGet -Path "${EntitySet}?`$select=$PrimaryId,uniquename&`$filter=_customapiid_value eq $CustomApiId" -Token $Token
    if ($result.value) {
        return @($result.value)
    }

    return @()
}

function Get-PluginTypeId {
    param(
        [string]$TypeName,
        [string]$Token
    )

    $result = Invoke-DataverseGet -Path "plugintypes?`$select=plugintypeid,typename&`$filter=typename eq '$TypeName'" -Token $Token
    if (-not $result.value -or $result.value.Count -eq 0) {
        throw "Plugin type '$TypeName' was not found. Deploy the plugin assembly first."
    }

    return $result.value[0].plugintypeid
}

function Get-RecordByUniqueName {
    param(
        [string]$EntitySet,
        [string]$UniqueName,
        [string]$PrimaryId,
        [string]$Token,
        [string]$CustomApiId
    )

    $filter = "uniquename eq '$UniqueName'"
    if (-not [string]::IsNullOrWhiteSpace($CustomApiId)) {
        $filter = "$filter and _customapiid_value eq $CustomApiId"
    }

    $result = Invoke-DataverseGet -Path "${EntitySet}?`$select=$PrimaryId,uniquename&`$filter=$filter" -Token $Token
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0]
    }

    return $null
}

function Upsert-CustomApi {
    param(
        [hashtable]$Definition,
        [string]$Token
    )

    $existing = Get-RecordByUniqueName -EntitySet 'customapis' -UniqueName $Definition.uniquename -PrimaryId 'customapiid' -Token $Token
    if ($existing) {
        if ($ForceRecreate) {
            Invoke-DataverseDelete -Path "customapis($($existing.customapiid))" -Token $Token
            $created = Invoke-DataversePost -EntitySet 'customapis' -Token $Token -Body $Definition
            return $created.customapiid
        }

        Invoke-DataversePatch -Path "customapis($($existing.customapiid))" -Token $Token -Body $Definition | Out-Null
        return $existing.customapiid
    }

    $created = Invoke-DataversePost -EntitySet 'customapis' -Token $Token -Body $Definition
    return $created.customapiid
}

function Upsert-RequestParameter {
    param(
        [hashtable]$Definition,
        [string]$CustomApiId,
        [string]$Token
    )

    if ($ForceRecreate) {
        Invoke-DataversePost -EntitySet 'customapirequestparameters' -Token $Token -Body $Definition | Out-Null
        return
    }

    $existing = Get-RecordByUniqueName -EntitySet 'customapirequestparameters' -UniqueName $Definition.uniquename -PrimaryId 'customapirequestparameterid' -Token $Token -CustomApiId $CustomApiId
    if ($existing) {
        Invoke-DataversePatch -Path "customapirequestparameters($($existing.customapirequestparameterid))" -Token $Token -Body $Definition | Out-Null
        return
    }

    Invoke-DataversePost -EntitySet 'customapirequestparameters' -Token $Token -Body $Definition | Out-Null
}

function Upsert-ResponseProperty {
    param(
        [hashtable]$Definition,
        [string]$CustomApiId,
        [string]$Token
    )

    if ($ForceRecreate) {
        Invoke-DataversePost -EntitySet 'customapiresponseproperties' -Token $Token -Body $Definition | Out-Null
        return
    }

    $existing = Get-RecordByUniqueName -EntitySet 'customapiresponseproperties' -UniqueName $Definition.uniquename -PrimaryId 'customapiresponsepropertyid' -Token $Token -CustomApiId $CustomApiId
    if ($existing) {
        Invoke-DataversePatch -Path "customapiresponseproperties($($existing.customapiresponsepropertyid))" -Token $Token -Body $Definition | Out-Null
        return
    }

    Invoke-DataversePost -EntitySet 'customapiresponseproperties' -Token $Token -Body $Definition | Out-Null
}

function Sync-RequestParameters {
    param(
        [hashtable[]]$Definitions,
        [string]$CustomApiId,
        [string]$Token
    )

    $expected = @{}
    foreach ($definition in $Definitions) {
        $expected[$definition.UniqueName] = $true
    }

    $existing = Get-CustomApiChildren -EntitySet 'customapirequestparameters' -PrimaryId 'customapirequestparameterid' -CustomApiId $CustomApiId -Token $Token
    foreach ($record in $existing) {
        if (-not $expected.ContainsKey($record.uniquename)) {
            Invoke-DataverseDelete -Path "customapirequestparameters($($record.customapirequestparameterid))" -Token $Token
        }
    }
}

function Sync-ResponseProperties {
    param(
        [hashtable[]]$Definitions,
        [string]$CustomApiId,
        [string]$Token
    )

    $expected = @{}
    foreach ($definition in $Definitions) {
        $expected[$definition.UniqueName] = $true
    }

    $existing = Get-CustomApiChildren -EntitySet 'customapiresponseproperties' -PrimaryId 'customapiresponsepropertyid' -CustomApiId $CustomApiId -Token $Token
    foreach ($record in $existing) {
        if (-not $expected.ContainsKey($record.uniquename)) {
            Invoke-DataverseDelete -Path "customapiresponseproperties($($record.customapiresponsepropertyid))" -Token $Token
        }
    }
}

$token = Get-DataverseToken

$definitions = @(
    @{
        UniqueName = 'co_SharePointEmbeddedCreateContainer'
        Name = 'SharePointEmbeddedCreateContainer'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedCreateContainer'
        RequestParameters = @(
            @{ UniqueName = 'ContainerTypeId'; Name = 'ContainerTypeId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'DisplayName'; Name = 'DisplayName'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'Description'; Name = 'Description'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10 },
            @{ UniqueName = 'DriveId'; Name = 'DriveId'; Type = 10 },
            @{ UniqueName = 'DisplayName'; Name = 'DisplayName'; Type = 10 },
            @{ UniqueName = 'Description'; Name = 'Description'; Type = 10 },
            @{ UniqueName = 'Status'; Name = 'Status'; Type = 10 },
            @{ UniqueName = 'WebUrl'; Name = 'WebUrl'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedUploadFile'
        Name = 'SharePointEmbeddedUploadFile'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedUploadFile'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'FileName'; Name = 'FileName'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'FileContentBase64'; Name = 'FileContentBase64'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'FolderPath'; Name = 'FolderPath'; Type = 10; IsOptional = $true },
            @{ UniqueName = 'ContentType'; Name = 'ContentType'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'DriveId'; Name = 'DriveId'; Type = 10 },
            @{ UniqueName = 'DriveItemId'; Name = 'DriveItemId'; Type = 10 },
            @{ UniqueName = 'Name'; Name = 'Name'; Type = 10 },
            @{ UniqueName = 'SizeInBytes'; Name = 'SizeInBytes'; Type = 10 },
            @{ UniqueName = 'WebUrl'; Name = 'WebUrl'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedGrantAccess'
        Name = 'SharePointEmbeddedGrantAccess'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedGrantAccess'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'UserPrincipalName'; Name = 'UserPrincipalName'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'Role'; Name = 'Role'; Type = 10; IsOptional = $false }
        )
        ResponseProperties = @(
            @{ UniqueName = 'PermissionId'; Name = 'PermissionId'; Type = 10 },
            @{ UniqueName = 'UserPrincipalName'; Name = 'UserPrincipalName'; Type = 10 },
            @{ UniqueName = 'Role'; Name = 'Role'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedRevokeAccess'
        Name = 'SharePointEmbeddedRevokeAccess'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedRevokeAccess'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'PermissionId'; Name = 'PermissionId'; Type = 10; IsOptional = $true },
            @{ UniqueName = 'UserPrincipalName'; Name = 'UserPrincipalName'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'PermissionId'; Name = 'PermissionId'; Type = 10 },
            @{ UniqueName = 'UserPrincipalName'; Name = 'UserPrincipalName'; Type = 10 },
            @{ UniqueName = 'Role'; Name = 'Role'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedCreateContainerWithFile'
        Name = 'SharePointEmbeddedCreateContainerWithFile'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedCreateContainerWithFile'
        RequestParameters = @(
            @{ UniqueName = 'ContainerTypeId'; Name = 'ContainerTypeId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'DisplayName'; Name = 'DisplayName'; Type = 10; IsOptional = $true },
            @{ UniqueName = 'Description'; Name = 'Description'; Type = 10; IsOptional = $true },
            @{ UniqueName = 'FileName'; Name = 'FileName'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'FileContentBase64'; Name = 'FileContentBase64'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'ContentType'; Name = 'ContentType'; Type = 10; IsOptional = $true },
            @{ UniqueName = 'PermissionsJson'; Name = 'PermissionsJson'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10 },
            @{ UniqueName = 'DriveId'; Name = 'DriveId'; Type = 10 },
            @{ UniqueName = 'DriveItemId'; Name = 'DriveItemId'; Type = 10 },
            @{ UniqueName = 'DisplayName'; Name = 'DisplayName'; Type = 10 },
            @{ UniqueName = 'Description'; Name = 'Description'; Type = 10 },
            @{ UniqueName = 'Status'; Name = 'Status'; Type = 10 },
            @{ UniqueName = 'WebUrl'; Name = 'WebUrl'; Type = 10 },
            @{ UniqueName = 'FileName'; Name = 'FileName'; Type = 10 },
            @{ UniqueName = 'PermissionsJson'; Name = 'PermissionsJson'; Type = 10 },
            @{ UniqueName = 'ContainerJson'; Name = 'ContainerJson'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedListContainers'
        Name = 'SharePointEmbeddedListContainers'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedListContainers'
        RequestParameters = @(
            @{ UniqueName = 'ContainerTypeId'; Name = 'ContainerTypeId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'Top'; Name = 'Top'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainersJson'; Name = 'ContainersJson'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedListDeletedContainers'
        Name = 'SharePointEmbeddedListDeletedContainers'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedListDeletedContainers'
        RequestParameters = @(
            @{ UniqueName = 'ContainerTypeId'; Name = 'ContainerTypeId'; Type = 10; IsOptional = $false },
            @{ UniqueName = 'Top'; Name = 'Top'; Type = 10; IsOptional = $true }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainersJson'; Name = 'ContainersJson'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedGetContainerDetails'
        Name = 'SharePointEmbeddedGetContainerDetails'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedGetContainerDetails'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainerJson'; Name = 'ContainerJson'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedDeleteContainer'
        Name = 'SharePointEmbeddedDeleteContainer'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedGetContainerDetails'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10 },
            @{ UniqueName = 'Status'; Name = 'Status'; Type = 10 }
        )
    },
    @{
        UniqueName = 'co_SharePointEmbeddedRestoreContainer'
        Name = 'SharePointEmbeddedRestoreContainer'
        IsFunction = $false
        PluginTypeName = 'ManagedIdentityPlugin.SharePointEmbeddedGetContainerDetails'
        RequestParameters = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10; IsOptional = $false }
        )
        ResponseProperties = @(
            @{ UniqueName = 'ContainerId'; Name = 'ContainerId'; Type = 10 },
            @{ UniqueName = 'Status'; Name = 'Status'; Type = 10 },
            @{ UniqueName = 'ContainerJson'; Name = 'ContainerJson'; Type = 10 }
        )
    }
)

foreach ($definition in $definitions) {
    $pluginTypeId = Get-PluginTypeId -TypeName $definition.PluginTypeName -Token $token

    $customApiBody = @{
        uniquename = $definition.UniqueName
        name = $definition.Name
        displayname = $definition.Name
        description = $definition.Name
        bindingtype = 0
        isfunction = $definition.IsFunction
        isprivate = $false
        allowedcustomprocessingsteptype = 0
        'PluginTypeId@odata.bind' = "/plugintypes($pluginTypeId)"
    }

    $customApiId = Upsert-CustomApi -Definition $customApiBody -Token $token

    foreach ($parameter in $definition.RequestParameters) {
        $requestParameterBody = @{
            uniquename = $parameter.UniqueName
            name = $parameter.Name
            displayname = $parameter.Name
            description = $parameter.Name
            type = $parameter.Type
            isoptional = $parameter.IsOptional
            'CustomAPIId@odata.bind' = "/customapis($customApiId)"
        }

        Upsert-RequestParameter -Definition $requestParameterBody -CustomApiId $customApiId -Token $token
    }

    Sync-RequestParameters -Definitions $definition.RequestParameters -CustomApiId $customApiId -Token $token

    foreach ($property in $definition.ResponseProperties) {
        $responsePropertyBody = @{
            uniquename = $property.UniqueName
            name = $property.Name
            displayname = $property.Name
            description = $property.Name
            type = $property.Type
            'CustomAPIId@odata.bind' = "/customapis($customApiId)"
        }

        Upsert-ResponseProperty -Definition $responsePropertyBody -CustomApiId $customApiId -Token $token
    }

    Sync-ResponseProperties -Definitions $definition.ResponseProperties -CustomApiId $customApiId -Token $token

    Write-Host "Registered custom API: $($definition.UniqueName)"
}