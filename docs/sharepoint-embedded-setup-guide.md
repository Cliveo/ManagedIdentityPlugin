# SharePoint Embedded Setup Guide

This guide is the public-safe setup companion for the SharePoint Embedded components in this repo.

It is intentionally narrow. It focuses on the runtime model used here:

- Dataverse custom APIs are the contract.
- A plugin package uses managed identity at runtime.
- Microsoft Graph is called app-only.
- Container type creation and billing stay outside this repo.

## The moving parts

This solution uses two application identities.

### Owning app

The owning app is the Microsoft Entra application that owns the SharePoint Embedded container type and performs tenant registration.

### Runtime app

The runtime app is the managed identity associated with the Dataverse plugin package. It creates containers, uploads files, lists content, and manages container membership through Microsoft Graph.

## What you need before starting

Before using the repo, make sure you already have:

1. A working managed identity plugin package association in Dataverse.
2. An existing SharePoint Embedded container type.
3. An owning application that is allowed to register that container type in the target tenant.
4. The billing model already decided and configured outside this repo.

## Required permissions

### Owning app

The owning app needs the permission required to register the container type in the customer tenant.

- `FileStorageContainerTypeReg.Selected`

### Runtime managed identity

The managed identity used by the plugin package needs:

- `FileStorageContainer.Selected`
- `Files.ReadWrite.All`

The runtime app also needs an application permission grant on the registered container type. In this repo the grant is typically `full` so the custom APIs can create containers, upload files, and manage membership.

## Why both Graph permissions and container-type grants matter

These are separate checks.

- Graph permissions control which SharePoint Embedded APIs the application can request tokens for.
- Container type registration enables the container type in the target tenant.
- Container type application grants allow a specific application to use that registered container type.

If one of those three is missing, the runtime flow will fail.

## Recommended setup sequence

### 1. Set up the managed identity plugin package

If you need the earlier foundation work, start with these posts:

- https://www.clive-oldridge.com/azure/2024/10/14/set-up-managed-identity-for-power-platform-plugins.html
- https://www.clive-oldridge.com/azure/2024/11/22/power-platform-plugin-package-managed-identity.html

### 2. Create and register the container type outside this repo

This repository does not automate container type creation or billing attachment.

It does still automate target-tenant registration of an already existing container type through the runtime access helper script.

Use Microsoft documentation for:

- container type creation
- container type registration prerequisites
- billing configuration

### 3. Grant the runtime app access

Use the script in this repo once you know these values:

- tenant ID
- owning app ID
- managed identity app ID
- managed identity service principal ID
- container type ID
- owning app client secret

Command:

```powershell
pwsh ./scripts/GrantManagedIdentitySharePointEmbeddedRuntimeAccess.ps1 \
  -TenantId '<tenant-id>' \
  -OwningAppId '<owning-app-id>' \
  -ManagedIdentityAppId '<managed-identity-app-id>' \
  -ManagedIdentityServicePrincipalId '<managed-identity-service-principal-id>' \
  -ContainerTypeId '<container-type-id>' \
  -OwningAppClientSecret '<client-secret>'
```

That script does two things:

- assigns the Graph application permissions needed by the runtime managed identity
- registers the existing container type in the target tenant and grants the runtime app access to that registration

### 4. Build and deploy the plugin package

Build the package:

```powershell
dotnet build ManagedIdentityPlugin.csproj
```

Deploy it using your normal Dataverse plugin package deployment flow.

### 5. Register the custom APIs in Dataverse

```powershell
pwsh ./scripts/RegisterSharePointEmbeddedCustomApis.ps1 -OrgUrl 'https://your-org.crm.dynamics.com'
```

### 6. Test through the local workbench or direct custom API calls

The local UI is a thin Dataverse client. It never calls Microsoft Graph directly.

```powershell
$env:DATAVERSE_URL = 'https://your-org.crm.dynamics.com'
node ./local-ui/auth-server.js
```

Then open `http://localhost:3001`.

## Custom API pattern used in this repo

The main building block is `co_SharePointEmbeddedCreateContainerWithFile`.

That API:

1. Creates a SharePoint Embedded container.
2. Applies initial container membership.
3. Uploads the first file.
4. Returns a serialized container snapshot.

This repo recommends one container per file. That keeps the sharing model simple because container membership becomes the boundary for access.

## Troubleshooting

### Access denied from Microsoft Graph

Check all three layers:

- Graph app permissions are assigned.
- Admin consent has been granted where required.
- The runtime app has a container-type application permission grant.

### Container creation fails but permissions look correct

Review SharePoint Embedded billing and tenant registration. Those parts are outside this repo and are common causes of setup drift.

### File upload works inconsistently

This sample uses the simple `PUT /content` upload path. Keep uploads within the standard simple-upload limits.

### Dataverse custom API registration fails

Deploy the latest plugin package first, then rerun `scripts/RegisterSharePointEmbeddedCustomApis.ps1`.