# Power Platform Plugin Package – Managed identity

This branch adds a SharePoint Embedded extension to the managed identity plugin foundation that already exists in the public repo.

The main idea is simple:

- Dataverse custom APIs stay as the public contract.
- The plugin package uses managed identity at runtime to call Microsoft Graph.
- SharePoint Embedded operations remain app-only.
- The front-end sample calls only Dataverse, not Graph directly.

## What is included

- SharePoint Embedded custom APIs for container create, upload, access management, listing, inspection, delete, and restore.
- A reusable Graph client that acquires its token from `IManagedIdentityService`.
- A Dataverse custom API registration script.
- A small script that grants the runtime managed identity the Graph roles and container-type application permission it needs.
- A local workbench that exercises the custom APIs through Dataverse.

The PowerShell helpers now live under `scripts/` so the repo root stays focused on the plugin package itself.

## What is intentionally not included

- Container type creation.
- Billing setup automation.
- Tenant-specific IDs, URLs, or internal proof artifacts.

Tenant registration is still covered by the lightweight runtime access script because the target tenant must register the container type before the managed identity can use it.

For container type setup, registration prerequisites, and billing, follow the Microsoft documentation and use the generic onboarding notes in [docs/sharepoint-embedded-setup-guide.md](docs/sharepoint-embedded-setup-guide.md).

## Architecture

The runtime flow is:

1. A caller invokes a Dataverse custom API.
2. The plugin acquires a Microsoft Graph token through managed identity.
3. The plugin calls SharePoint Embedded through Microsoft Graph v1.0.
4. The plugin returns a Dataverse-friendly response payload.

The recommended pattern in this repo is one container per file. That keeps the access model simple because container membership becomes the sharing boundary.

## Source layout

The code is grouped by domain:

- `Domains/BlobStorage`: blob example plugin.
- `Domains/KeyVault`: Key Vault example plugin.
- `Domains/PowerAutomate`: flow example plugin.
- `Domains/SharePointEmbedded`: SharePoint Embedded custom APIs and Graph client.
- `Infrastructure`: shared plugin base and token credential plumbing.
- `scripts`: Dataverse registration and managed identity helper scripts.

## Custom APIs

The public SharePoint Embedded API surface in this branch is:

- `co_SharePointEmbeddedCreateContainer`
- `co_SharePointEmbeddedUploadFile`
- `co_SharePointEmbeddedGrantAccess`
- `co_SharePointEmbeddedRevokeAccess`
- `co_SharePointEmbeddedCreateContainerWithFile`
- `co_SharePointEmbeddedListContainers`
- `co_SharePointEmbeddedListDeletedContainers`
- `co_SharePointEmbeddedGetContainerDetails`
- `co_SharePointEmbeddedDeleteContainer`
- `co_SharePointEmbeddedRestoreContainer`

The combined create flow is usually the best place to start:

- `co_SharePointEmbeddedCreateContainerWithFile` creates the container.
- It optionally grants container access to selected users.
- It uploads the initial file.
- It returns both the file result and a serialized container snapshot.

## Minimal setup

1. Start from the managed identity plugin package setup described in the existing public blog posts.
2. Ensure you already have a SharePoint Embedded container type and the owning app registration required to register it in the target tenant.
3. Grant the runtime managed identity access:

```powershell
pwsh ./scripts/GrantManagedIdentitySharePointEmbeddedRuntimeAccess.ps1 \
	-TenantId '<tenant-id>' \
	-OwningAppId '<owning-app-id>' \
	-ManagedIdentityAppId '<managed-identity-app-id>' \
	-ManagedIdentityServicePrincipalId '<managed-identity-service-principal-id>' \
	-ContainerTypeId '<container-type-id>' \
	-OwningAppClientSecret '<client-secret>'
```

4. Build and deploy the plugin package.
5. Register the Dataverse custom APIs:

```powershell
pwsh ./scripts/RegisterSharePointEmbeddedCustomApis.ps1 -OrgUrl 'https://your-org.crm.dynamics.com'
```

That helper script assigns the runtime Graph app roles, registers the existing container type in the target tenant, and then grants the managed identity application permission on that registration.

## Local workbench

The sample UI is a thin Dataverse client for the custom APIs. It does not call Graph directly.

Run it like this:

```powershell
$env:DATAVERSE_URL = 'https://your-org.crm.dynamics.com'
node ./local-ui/auth-server.js
```

Then open `http://localhost:3001`.

The page uses your Azure CLI token for Dataverse, loads active users, creates containers with an initial file, updates container membership, and browses active or deleted containers.

## Background

If you want the earlier setup material that this extension builds on, start here:

- https://www.clive-oldridge.com/azure/2024/10/14/set-up-managed-identity-for-power-platform-plugins.html
- https://www.clive-oldridge.com/azure/2024/11/22/power-platform-plugin-package-managed-identity.html
