const DATAVERSE_PROXY_BASE = '/api/data/v9.2/';
const DEFAULT_CONTAINER_TYPE_ID = '';
const state = {
  token: null,
  tokenPromise: null,
  appConfig: {
    dataverseUrl: 'Not configured',
    authMode: 'Azure CLI',
    isConfigured: false
  },
  users: [],
  selectedUsers: [],
  containers: [],
  deletedContainers: [],
  recycleBinVisible: false,
  recycleBinLoaded: false,
  selectedContainerIds: []
};

const roleOptions = ['reader', 'writer', 'manager', 'owner'];

const elements = {
  containerTypeId: document.getElementById('containerTypeId'),
  displayName: document.getElementById('displayName'),
  description: document.getElementById('description'),
  fileInput: document.getElementById('fileInput'),
  fileSummary: document.getElementById('fileSummary'),
  userSearch: document.getElementById('userSearch'),
  availableUsers: document.getElementById('availableUsers'),
  selectedUsers: document.getElementById('selectedUsers'),
  createButton: document.getElementById('createButton'),
  refreshUsersButton: document.getElementById('refreshUsersButton'),
  refreshContainersButton: document.getElementById('refreshContainersButton'),
  toggleRecycleBinButton: document.getElementById('toggleRecycleBinButton'),
  selectAllContainers: document.getElementById('selectAllContainers'),
  bulkDeleteButton: document.getElementById('bulkDeleteButton'),
  containers: document.getElementById('containers'),
  recycleBinPanel: document.getElementById('recycleBinPanel'),
  refreshDeletedContainersButton: document.getElementById('refreshDeletedContainersButton'),
  deletedContainers: document.getElementById('deletedContainers'),
  statusOutput: document.getElementById('statusOutput'),
  statusSummary: document.getElementById('statusSummary'),
  tokenState: document.getElementById('tokenState'),
  orgUrl: document.getElementById('orgUrl'),
  authModeLabel: document.getElementById('authModeLabel'),
  userCountPill: document.getElementById('userCountPill'),
  visibleUserCountPill: document.getElementById('visibleUserCountPill'),
  selectionCountPill: document.getElementById('selectionCountPill'),
  topCount: document.getElementById('topCount'),
  overviewUsers: document.getElementById('overviewUsers'),
  overviewAssignments: document.getElementById('overviewAssignments'),
  selectedFileName: document.getElementById('selectedFileName'),
  selectedFileMeta: document.getElementById('selectedFileMeta'),
  containerCountPill: document.getElementById('containerCountPill'),
  containerHeaderCount: document.getElementById('containerHeaderCount'),
  deletedContainerHeaderCount: document.getElementById('deletedContainerHeaderCount')
};

const setStatusSummary = (title, detail) => {
  const titleElement = elements.statusSummary.querySelector('strong');
  const detailElement = elements.statusSummary.querySelector('small');

  if (titleElement) {
    titleElement.textContent = title;
  }

  if (detailElement) {
    detailElement.textContent = detail || '';
  }
};

const setTokenState = (stateName, label) => {
  elements.tokenState.className = `status-indicator ${stateName}`;
  elements.tokenState.textContent = label;
};

const updateOverview = () => {
  elements.overviewUsers.textContent = String(state.users.length);
  elements.overviewAssignments.textContent = String(state.selectedUsers.length);

  elements.containerCountPill.textContent = String(state.containers.length);
  elements.containerHeaderCount.textContent = `${state.containers.length} loaded`;
  elements.deletedContainerHeaderCount.textContent = `${state.deletedContainers.length} deleted`;
  elements.toggleRecycleBinButton.textContent = state.recycleBinVisible
    ? `Hide Recycle Bin (${state.deletedContainers.length})`
    : `Recycle Bin (${state.deletedContainers.length})`;

  const file = elements.fileInput.files[0];
  if (!file) {
    elements.selectedFileName.textContent = 'No file selected';
    elements.selectedFileMeta.textContent = 'Choose a file to prefill the display name and build the create payload.';
    elements.fileSummary.className = 'file-summary-pill hidden';
    elements.fileSummary.textContent = '';
    return;
  }

  const fileMeta = `${formatBytes(file.size)} · ${file.type || 'application/octet-stream'}`;
  elements.selectedFileName.textContent = file.name;
  elements.selectedFileMeta.textContent = fileMeta;
  elements.fileSummary.className = 'file-summary-pill';
  elements.fileSummary.innerHTML = `
    <span class="meta-label">Selected file</span>
    <strong>${escapeHtml(file.name)}</strong>
    <small>${escapeHtml(fileMeta)}</small>
  `;
};

const writeStatus = (message, payload) => {
  const lines = [`[${new Date().toLocaleTimeString()}] ${message}`];
  if (payload !== undefined) {
    lines.push(typeof payload === 'string' ? payload : JSON.stringify(payload, null, 2));
  }
  elements.statusOutput.textContent = `${lines.join('\n')}\n\n${elements.statusOutput.textContent}`.trim();
  setStatusSummary(message, payload === undefined ? 'Latest action recorded.' : 'Latest action includes details in the run log below.');
};

const extractErrorMessage = (text, fallbackMessage) => {
  if (!text) {
    return fallbackMessage;
  }

  try {
    const parsed = JSON.parse(text);
    if (typeof parsed === 'string' && parsed.trim()) {
      return parsed;
    }

    if (parsed?.error?.message) {
      return parsed.error.message;
    }

    if (parsed?.Message) {
      return parsed.Message;
    }
  } catch {
    // Fall back to the raw response body when the payload is not JSON.
  }

  return text || fallbackMessage;
};

const loadConfig = async () => {
  const response = await fetch('/config');
  const config = await response.json();

  state.appConfig = {
    dataverseUrl: config.dataverseUrl || 'Not configured',
    authMode: config.authMode || 'Azure CLI',
    isConfigured: Boolean(config.isConfigured)
  };

  elements.orgUrl.textContent = state.appConfig.dataverseUrl;
  elements.authModeLabel.textContent = state.appConfig.authMode;
};

const setBusy = (button, busy, label) => {
  button.disabled = busy;
  if (busy && label) {
    button.dataset.defaultLabel = button.dataset.defaultLabel || button.textContent;
    button.textContent = label;
    return;
  }

  if (!busy && button.dataset.defaultLabel) {
    button.textContent = button.dataset.defaultLabel;
  }
};

const updateContainerSelectionUi = () => {
  const currentContainerIds = state.containers.map((container) => container.ContainerId);
  state.selectedContainerIds = state.selectedContainerIds.filter((containerId) => currentContainerIds.includes(containerId));

  const selectedCount = state.selectedContainerIds.length;
  const totalCount = currentContainerIds.length;

  elements.bulkDeleteButton.disabled = selectedCount === 0;
  elements.bulkDeleteButton.textContent = selectedCount > 0 ? `Delete selected (${selectedCount})` : 'Delete selected';

  elements.selectAllContainers.checked = totalCount > 0 && selectedCount === totalCount;
  elements.selectAllContainers.indeterminate = selectedCount > 0 && selectedCount < totalCount;
};

const getToken = async () => {
  if (!state.appConfig.isConfigured) {
    throw new Error('Set DATAVERSE_URL before using the local workbench.');
  }

  if (state.token) {
    return state.token;
  }

  if (state.tokenPromise) {
    return state.tokenPromise;
  }

  state.tokenPromise = (async () => {
    setTokenState('busy', 'Fetching token');
    setStatusSummary('Fetching Dataverse token', 'Using the local Azure CLI-backed proxy.');
    const response = await fetch('/token');
    const text = await response.text();
    if (!response.ok) {
      setTokenState('error', 'Token failed');
      throw new Error(text || 'Token request failed.');
    }

    state.token = text;
    setTokenState('done', 'Token ready');
    setStatusSummary('Dataverse token ready', 'The local proxy is ready to send requests.');
    return state.token;
  })();

  try {
    return await state.tokenPromise;
  } finally {
    state.tokenPromise = null;
  }
};

const dataverseFetch = async (path, options = {}) => {
  await getToken();
  const headers = {
    Accept: 'application/json',
    'Content-Type': 'application/json',
    'OData-Version': '4.0',
    'OData-MaxVersion': '4.0',
    ...(options.headers || {})
  };

  const response = await fetch(`${DATAVERSE_PROXY_BASE}${path}`, {
    method: options.method || 'GET',
    headers,
    body: options.body
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(extractErrorMessage(text, `Dataverse request failed for ${path}`));
  }

  return text ? JSON.parse(text) : {};
};

const loadUsers = async () => {
  setBusy(elements.refreshUsersButton, true, 'Loading...');
  try {
    const result = await dataverseFetch("systemusers?$select=systemuserid,fullname,internalemailaddress,domainname,azureactivedirectoryobjectid,applicationid&$filter=isdisabled eq false and applicationid eq null and azureactivedirectoryobjectid ne null and fullname ne null");
    state.users = (result.value || [])
      .map((user) => ({
        id: user.systemuserid,
        name: user.fullname || user.domainname || user.internalemailaddress,
        upn: user.domainname || user.internalemailaddress,
        aadId: user.azureactivedirectoryobjectid,
        applicationId: user.applicationid
      }))
      .filter((user) => user.upn && user.name && user.name.trim() && !user.applicationId)
      .sort((left, right) => left.name.localeCompare(right.name));

    elements.userCountPill.textContent = `${state.users.length} loaded`;
    updateOverview();
    writeStatus('Loaded named Dataverse users.', { count: state.users.length });
    renderUsers();
  } finally {
    setBusy(elements.refreshUsersButton, false);
  }
};

const renderUsers = () => {
  const filter = elements.userSearch.value.trim().toLowerCase();
  const selectedUpns = new Set(state.selectedUsers.map((user) => user.userPrincipalName.toLowerCase()));
  const users = state.users.filter((user) => {
    if (selectedUpns.has(user.upn.toLowerCase())) {
      return false;
    }

    if (!filter) {
      return true;
    }

    return user.name.toLowerCase().includes(filter) || user.upn.toLowerCase().includes(filter);
  });

  elements.visibleUserCountPill.textContent = `${users.length} visible`;
  elements.visibleUserCountPill.classList.toggle('hidden', !filter);

  if (!users.length) {
    elements.availableUsers.innerHTML = '<div class="empty-message">No matching users.</div>';
    return;
  }

  elements.availableUsers.innerHTML = users.slice(0, 40).map((user) => `
    <div class="user-item">
      <div class="user-info">
        <strong class="user-name">${escapeHtml(user.name)}</strong>
        <small class="user-upn">${escapeHtml(user.upn)}</small>
      </div>
      <button class="btn btn-ghost btn-sm" data-add-upn="${escapeAttribute(user.upn)}" data-add-name="${escapeAttribute(user.name)}">Add</button>
    </div>
  `).join('');
};

const renderSelectedUsers = () => {
  elements.selectionCountPill.textContent = String(state.selectedUsers.length);
  updateOverview();
  if (!state.selectedUsers.length) {
    elements.selectedUsers.className = 'list-container';
    elements.selectedUsers.innerHTML = '<div class="empty-message">No users assigned yet.</div>';
    return;
  }

  elements.selectedUsers.className = 'list-container';
  elements.selectedUsers.innerHTML = state.selectedUsers.map((user) => `
    <div class="user-item">
      <div class="user-info">
        <strong class="user-name">${escapeHtml(user.name)}</strong>
        <small class="user-upn">${escapeHtml(user.userPrincipalName)}</small>
      </div>
      <div class="selected-controls">
        <select class="role-select" data-role-upn="${escapeAttribute(user.userPrincipalName)}">
          ${roleOptions.map((role) => `<option value="${role}" ${role === user.role ? 'selected' : ''}>${role}</option>`).join('')}
        </select>
        <button class="btn btn-ghost btn-sm" data-remove-upn="${escapeAttribute(user.userPrincipalName)}">Remove</button>
      </div>
    </div>
  `).join('');
};

const addSelectedUser = (name, userPrincipalName) => {
  if (state.selectedUsers.some((user) => user.userPrincipalName.toLowerCase() === userPrincipalName.toLowerCase())) {
    return;
  }

  state.selectedUsers.push({ name, userPrincipalName, role: 'reader' });
  renderUsers();
  renderSelectedUsers();
};

const removeSelectedUser = (userPrincipalName) => {
  state.selectedUsers = state.selectedUsers.filter((user) => user.userPrincipalName !== userPrincipalName);
  renderUsers();
  renderSelectedUsers();
};

const readFileAsBase64 = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => {
    const result = String(reader.result || '');
    const commaIndex = result.indexOf(',');
    resolve(commaIndex >= 0 ? result.slice(commaIndex + 1) : result);
  };
  reader.onerror = () => reject(reader.error || new Error('Unable to read file.'));
  reader.readAsDataURL(file);
});

const createContainer = async () => {
  const file = elements.fileInput.files[0];
  const containerTypeId = elements.containerTypeId.value.trim();
  if (!containerTypeId) {
    throw new Error('Container type id is required.');
  }
  if (!file) {
    throw new Error('Choose a file before creating the container.');
  }

  const fileContentBase64 = await readFileAsBase64(file);
  const payload = {
    ContainerTypeId: containerTypeId,
    DisplayName: elements.displayName.value.trim() || file.name.replace(/\.[^.]+$/, ''),
    Description: elements.description.value.trim(),
    FileName: file.name,
    FileContentBase64: fileContentBase64,
    ContentType: file.type || 'application/octet-stream',
    PermissionsJson: JSON.stringify(state.selectedUsers.map((user) => ({
      userPrincipalName: user.userPrincipalName,
      role: user.role
    })))
  };

  setBusy(elements.createButton, true, 'Creating...');
  try {
    const result = await dataverseFetch('co_SharePointEmbeddedCreateContainerWithFile', {
      method: 'POST',
      body: JSON.stringify(payload)
    });

    const details = JSON.parse(result.ContainerJson || '{}');

    writeStatus('Created container and uploaded file.', result);
    await refreshContainers();
  } finally {
    setBusy(elements.createButton, false);
  }
};

const refreshContainers = async () => {
  const containerTypeId = elements.containerTypeId.value.trim();
  if (!containerTypeId) {
    writeStatus('Skipping container refresh because no container type id is set.');
    return;
  }

  setBusy(elements.refreshContainersButton, true, 'Refreshing...');
  try {
    const result = await dataverseFetch('co_SharePointEmbeddedListContainers', {
      method: 'POST',
      body: JSON.stringify({
        ContainerTypeId: containerTypeId,
        Top: elements.topCount.value.trim() || '20'
      })
    });

    state.containers = JSON.parse(result.ContainersJson || '[]');
    updateContainerSelectionUi();
    renderContainers();
    updateOverview();
    writeStatus('Loaded container list.', { count: state.containers.length });
  } finally {
    setBusy(elements.refreshContainersButton, false);
  }
};

const refreshDeletedContainers = async ({ writeLog = true } = {}) => {
  const containerTypeId = elements.containerTypeId.value.trim();
  if (!containerTypeId) {
    return;
  }

  setBusy(elements.refreshDeletedContainersButton, true, 'Refreshing...');
  try {
    const result = await dataverseFetch('co_SharePointEmbeddedListDeletedContainers', {
      method: 'POST',
      body: JSON.stringify({
        ContainerTypeId: containerTypeId,
        Top: elements.topCount.value.trim() || '20'
      })
    });

    state.deletedContainers = JSON.parse(result.ContainersJson || '[]');
    state.recycleBinLoaded = true;
    renderDeletedContainers();
    updateOverview();
    if (writeLog) {
      writeStatus('Loaded deleted container list.', { count: state.deletedContainers.length });
    }
  } finally {
    setBusy(elements.refreshDeletedContainersButton, false);
  }
};

const deleteContainer = async (containerId) => dataverseFetch('co_SharePointEmbeddedDeleteContainer', {
  method: 'POST',
  body: JSON.stringify({
    ContainerId: containerId
  })
});

const restoreContainer = async (containerId) => dataverseFetch('co_SharePointEmbeddedRestoreContainer', {
  method: 'POST',
  body: JSON.stringify({
    ContainerId: containerId
  })
});

const setContainerSelected = (containerId, isSelected) => {
  const selectedIds = new Set(state.selectedContainerIds);
  if (isSelected) {
    selectedIds.add(containerId);
  } else {
    selectedIds.delete(containerId);
  }

  state.selectedContainerIds = [...selectedIds];
  updateContainerSelectionUi();
};

const setAllContainersSelected = (isSelected) => {
  state.selectedContainerIds = isSelected ? state.containers.map((container) => container.ContainerId) : [];
  updateContainerSelectionUi();
  renderContainers();
};

const deleteSelectedContainers = async (containerIds) => {
  if (!containerIds.length) {
    return;
  }

  const uniqueIds = [...new Set(containerIds)];
  await Promise.all(uniqueIds.map((containerId) => deleteContainer(containerId)));
  state.selectedContainerIds = state.selectedContainerIds.filter((containerId) => !uniqueIds.includes(containerId));
};

const grantContainerAccess = async (containerId, userPrincipalName, role) => {
  const normalizedUpn = String(userPrincipalName || '').trim();
  if (!normalizedUpn) {
    throw new Error('User principal name is required.');
  }

  return dataverseFetch('co_SharePointEmbeddedGrantAccess', {
    method: 'POST',
    body: JSON.stringify({
      ContainerId: containerId,
      UserPrincipalName: normalizedUpn,
      Role: role
    })
  });
};

const revokeContainerAccess = async (containerId, permissionId, userPrincipalName) => dataverseFetch('co_SharePointEmbeddedRevokeAccess', {
  method: 'POST',
  body: JSON.stringify({
    ContainerId: containerId,
    PermissionId: permissionId || undefined,
    UserPrincipalName: userPrincipalName || undefined
  })
});

const inspectContainer = async (containerId) => {
  const result = await dataverseFetch('co_SharePointEmbeddedGetContainerDetails', {
    method: 'POST',
    body: JSON.stringify({ ContainerId: containerId })
  });
  setStatusSummary('Container inspected', `Fetched live details for ${containerId}.`);
  writeStatus(`Container details for ${containerId}`, JSON.parse(result.ContainerJson || '{}'));
};

const renderDeletedContainers = () => {
  elements.recycleBinPanel.classList.toggle('hidden', !state.recycleBinVisible);

  if (!state.recycleBinVisible) {
    return;
  }

  if (!state.deletedContainers.length) {
    elements.deletedContainers.className = 'results-grid recycle-bin-grid';
    elements.deletedContainers.innerHTML = '<div class="empty-message">Recycle bin is empty for this container type.</div>';
    return;
  }

  elements.deletedContainers.className = 'results-grid recycle-bin-grid';
  elements.deletedContainers.innerHTML = state.deletedContainers.map((container) => `
    <article class="container-item deleted-container-item">
      <div class="container-head">
        <div class="container-head-info">
          <div class="container-title-group">
            <div class="container-title-row">
              <strong class="container-name">${escapeHtml(container.DisplayName || container.ContainerId)}</strong>
              <div class="info-flyout-wrapper" tabindex="0">
                <button class="info-icon-btn" title="View Metadata">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>
                </button>
                <div class="info-flyout">
                  <div class="info-flyout-row">
                    <span class="info-flyout-label">Container ID</span>
                    <span class="info-flyout-value copyable" title="Click to copy">${escapeHtml(container.ContainerId)}</span>
                  </div>
                  <div class="info-flyout-row">
                    <span class="info-flyout-label">Container Type ID</span>
                    <span class="info-flyout-value copyable" title="Click to copy">${escapeHtml(container.ContainerTypeId || 'Unknown')}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="container-head-actions">
          <button class="btn btn-secondary btn-sm" data-restore-container="${escapeAttribute(container.ContainerId)}">Restore</button>
        </div>
      </div>
      <div class="container-body-grid">
        <div class="container-section">
          <small class="muted">Deleted</small>
          <span class="deleted-meta">${escapeHtml(container.DeletedDateTime || 'Recently deleted')}</span>
        </div>
        ${container.Description ? `
          <div class="container-section">
            <small class="muted">Description</small>
            <span class="deleted-meta">${escapeHtml(container.Description)}</span>
          </div>
        ` : ''}
      </div>
    </article>
  `).join('');
};

const renderContainers = () => {
  if (!state.containers.length) {
    elements.containers.className = 'results-grid';
    elements.containers.innerHTML = '<div class="empty-message">No containers returned for this container type.</div>';
    updateContainerSelectionUi();
    updateOverview();
    return;
  }

  elements.containers.className = 'results-grid';
  elements.containers.innerHTML = state.containers.map((container) => {
    const file = (container.Files || []).find((item) => !item.IsFolder) || container.Files?.[0];
    const permissions = container.Permissions || [];
    const isSelected = state.selectedContainerIds.includes(container.ContainerId);

    return `
      <article class="container-item">
        <div class="container-head">
          <div class="container-head-info">
            <input class="container-checkbox" type="checkbox" data-select-container="${escapeAttribute(container.ContainerId)}" ${isSelected ? 'checked' : ''} aria-label="Select container">
            <div class="container-title-group">
              <div class="container-title-row">
                <strong class="container-name">${escapeHtml(container.DisplayName || container.ContainerId)}</strong>
                <div class="info-flyout-wrapper" tabindex="0">
                  <button class="info-icon-btn" title="View Metadata">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>
                  </button>
                  <div class="info-flyout">
                    <div class="info-flyout-row">
                      <span class="info-flyout-label">Container ID</span>
                      <span class="info-flyout-value copyable" title="Click to copy">${escapeHtml(container.ContainerId)}</span>
                    </div>
                    <div class="info-flyout-row">
                      <span class="info-flyout-label">Container Type ID</span>
                      <span class="info-flyout-value copyable" title="Click to copy">${escapeHtml(container.ContainerTypeId || 'Unknown')}</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div class="container-head-actions">
            <button class="btn btn-ghost btn-sm" data-inspect-container="${escapeAttribute(container.ContainerId)}">Inspect</button>
            <button class="btn btn-danger btn-sm" data-delete-container="${escapeAttribute(container.ContainerId)}">Delete</button>
          </div>
        </div>
        <div class="container-body-grid">
          <div class="container-section">
            <small class="muted">Primary File</small>
            ${file ? `
              <div class="file-compact">
                <span class="file-name" title="${escapeHtml(file.Name)}">${escapeHtml(file.Name)}</span>
                <span class="file-meta">${formatBytes(file.Size)} &middot; ${escapeHtml(file.ContentType || 'unknown type')}</span>
                ${file.WebUrl ? `<a class="container-link" href="${escapeAttribute(file.WebUrl)}" target="_blank" rel="noreferrer">Open raw file</a>` : ''}
              </div>
            ` : '<span class="empty-state">No file found in root.</span>'}
          </div>
          <div class="container-section permissions-section">
            <details class="permissions-details">
              <summary class="permissions-summary">
                <small class="muted">Privileges</small> 
                <span class="badge primary-subtle">${permissions.length} user${permissions.length === 1 ? '' : 's'}</span>
              </summary>
              <div class="permission-editor mt-2">
                ${permissions.length ? permissions.map((permission) => `
                  <div class="permission-row">
                    <div class="permission-meta">
                      <strong>${escapeHtml(permission.UserPrincipalName)}</strong>
                    </div>
                    <div class="permission-actions flex-wrap">
                      <select class="role-select" data-permission-role="${escapeAttribute(container.ContainerId)}|${escapeAttribute(permission.UserPrincipalName)}">
                        ${roleOptions.map((role) => `<option value="${role}" ${role === permission.Role ? 'selected' : ''}>${role}</option>`).join('')}
                      </select>
                      <button class="btn btn-secondary btn-sm" data-save-permission="${escapeAttribute(container.ContainerId)}" data-save-upn="${escapeAttribute(permission.UserPrincipalName)}">Save</button>
                      <button class="btn btn-ghost btn-sm" data-remove-permission="${escapeAttribute(container.ContainerId)}" data-remove-id="${escapeAttribute(permission.PermissionId || '')}" data-remove-upn="${escapeAttribute(permission.UserPrincipalName)}">Remove</button>
                    </div>
                  </div>
                `).join('') : '<div class="empty-message compact-empty">No explicit users</div>'}
                <div class="permission-add-row flex-wrap">
                  <input class="input-modern permission-input" data-add-upn-input="${escapeAttribute(container.ContainerId)}" placeholder="Add user by UPN">
                  <select class="role-select" data-add-role-input="${escapeAttribute(container.ContainerId)}">
                    ${roleOptions.map((role) => `<option value="${role}" ${role === 'reader' ? 'selected' : ''}>${role}</option>`).join('')}
                  </select>
                  <button class="btn btn-secondary btn-sm" data-add-permission="${escapeAttribute(container.ContainerId)}">Add</button>
                </div>
              </div>
            </details>
          </div>
        </div>
      </article>
    `;
  }).join('');
  updateContainerSelectionUi();
};

const escapeHtml = (value) => String(value || '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

const escapeAttribute = (value) => escapeHtml(value).replace(/'/g, '&#39;');

const formatBytes = (value) => {
  const size = Number(value || 0);
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
};

elements.refreshUsersButton.addEventListener('click', async () => {
  try {
    await loadUsers();
  } catch (error) {
    writeStatus('Failed to load users.', error.message);
  }
});

elements.userSearch.addEventListener('input', renderUsers);

elements.availableUsers.addEventListener('click', (event) => {
  const button = event.target.closest('[data-add-upn]');
  if (!button) {
    return;
  }

  addSelectedUser(button.getAttribute('data-add-name'), button.getAttribute('data-add-upn'));
});

elements.selectedUsers.addEventListener('click', (event) => {
  const button = event.target.closest('[data-remove-upn]');
  if (!button) {
    return;
  }

  removeSelectedUser(button.getAttribute('data-remove-upn'));
});

elements.selectedUsers.addEventListener('change', (event) => {
  const select = event.target.closest('[data-role-upn]');
  if (!select) {
    return;
  }

  const user = state.selectedUsers.find((entry) => entry.userPrincipalName === select.getAttribute('data-role-upn'));
  if (user) {
    user.role = select.value;
  }
});

elements.createButton.addEventListener('click', async () => {
  try {
    await createContainer();
  } catch (error) {
    writeStatus('Create flow failed.', error.message);
  }
});

elements.refreshContainersButton.addEventListener('click', async () => {
  try {
    await refreshContainers();
  } catch (error) {
    writeStatus('Container refresh failed.', error.message);
  }
});

elements.toggleRecycleBinButton.addEventListener('click', async () => {
  state.recycleBinVisible = !state.recycleBinVisible;

  try {
    if (state.recycleBinVisible && !state.recycleBinLoaded) {
      await refreshDeletedContainers();
      return;
    }

    renderDeletedContainers();
    updateOverview();
  } catch (error) {
    writeStatus('Recycle bin load failed.', error.message);
  }
});

elements.refreshDeletedContainersButton.addEventListener('click', async () => {
  try {
    await refreshDeletedContainers();
  } catch (error) {
    writeStatus('Recycle bin refresh failed.', error.message);
  }
});

elements.selectAllContainers.addEventListener('change', () => {
  setAllContainersSelected(elements.selectAllContainers.checked);
});

elements.bulkDeleteButton.addEventListener('click', async () => {
  const containerIds = [...state.selectedContainerIds];
  if (!containerIds.length) {
    return;
  }

  if (!window.confirm(`Delete ${containerIds.length} selected container${containerIds.length === 1 ? '' : 's'}? This cannot be undone.`)) {
    return;
  }

  try {
    setBusy(elements.bulkDeleteButton, true, 'Deleting...');
    await deleteSelectedContainers(containerIds);
    writeStatus('Deleted selected containers.', { count: containerIds.length, containerIds });
    await refreshContainers();
    await refreshDeletedContainers({ writeLog: false });
  } catch (error) {
    writeStatus('Bulk delete failed.', error.message);
  } finally {
    setBusy(elements.bulkDeleteButton, false);
  }
});

elements.containers.addEventListener('click', async (event) => {
  const button = event.target.closest('[data-inspect-container]');
  if (button) {
    try {
      await inspectContainer(button.getAttribute('data-inspect-container'));
    } catch (error) {
      writeStatus('Container inspection failed.', error.message);
    }
    return;
  }

  const deleteButton = event.target.closest('[data-delete-container]');
  if (deleteButton) {
    const containerId = deleteButton.getAttribute('data-delete-container');
    if (!window.confirm(`Delete container ${containerId}? This cannot be undone.`)) {
      return;
    }

    try {
      setBusy(deleteButton, true, 'Deleting...');
      await deleteSelectedContainers([containerId]);
      writeStatus('Deleted container.', { containerId });
      await refreshContainers();
      await refreshDeletedContainers({ writeLog: false });
    } catch (error) {
      writeStatus('Container delete failed.', error.message);
    } finally {
      setBusy(deleteButton, false);
    }
    return;
  }

  const saveButton = event.target.closest('[data-save-permission]');
  if (saveButton) {
    const containerId = saveButton.getAttribute('data-save-permission');
    const userPrincipalName = saveButton.getAttribute('data-save-upn');
    const roleSelect = elements.containers.querySelector(`[data-permission-role="${CSS.escape(`${containerId}|${userPrincipalName}`)}"]`);

    try {
      setBusy(saveButton, true, 'Saving...');
      await grantContainerAccess(containerId, userPrincipalName, roleSelect?.value || 'reader');
      writeStatus('Updated container permission.', { containerId, userPrincipalName, role: roleSelect?.value || 'reader' });
      await refreshContainers();
    } catch (error) {
      writeStatus('Permission update failed.', error.message);
    } finally {
      setBusy(saveButton, false);
    }
    return;
  }

  const removeButton = event.target.closest('[data-remove-permission]');
  if (removeButton) {
    const containerId = removeButton.getAttribute('data-remove-permission');
    const permissionId = removeButton.getAttribute('data-remove-id');
    const userPrincipalName = removeButton.getAttribute('data-remove-upn');

    try {
      setBusy(removeButton, true, 'Removing...');
      await revokeContainerAccess(containerId, permissionId, userPrincipalName);
      writeStatus('Removed container permission.', { containerId, userPrincipalName });
      await refreshContainers();
    } catch (error) {
      writeStatus('Permission removal failed.', error.message);
    } finally {
      setBusy(removeButton, false);
    }
    return;
  }

  const addButton = event.target.closest('[data-add-permission]');
  if (addButton) {
    const containerId = addButton.getAttribute('data-add-permission');
    const upnInput = elements.containers.querySelector(`[data-add-upn-input="${CSS.escape(containerId)}"]`);
    const roleInput = elements.containers.querySelector(`[data-add-role-input="${CSS.escape(containerId)}"]`);
    const userPrincipalName = upnInput?.value.trim();

    try {
      if (!userPrincipalName) {
        throw new Error('Enter a user UPN before adding a permission.');
      }

      setBusy(addButton, true, 'Adding...');
      await grantContainerAccess(containerId, userPrincipalName, roleInput?.value || 'reader');
      writeStatus('Added container permission.', { containerId, userPrincipalName, role: roleInput?.value || 'reader' });
      await refreshContainers();
    } catch (error) {
      writeStatus('Permission add failed.', error.message);
    } finally {
      setBusy(addButton, false);
    }
  }
});

elements.containers.addEventListener('change', (event) => {
  const checkbox = event.target.closest('[data-select-container]');
  if (!checkbox) {
    return;
  }

  setContainerSelected(checkbox.getAttribute('data-select-container'), checkbox.checked);
});

elements.deletedContainers.addEventListener('click', async (event) => {
  const restoreButton = event.target.closest('[data-restore-container]');
  if (!restoreButton) {
    return;
  }

  const containerId = restoreButton.getAttribute('data-restore-container');

  try {
    setBusy(restoreButton, true, 'Restoring...');
    await restoreContainer(containerId);
    writeStatus('Restored container from recycle bin.', { containerId });
    await refreshContainers();
    await refreshDeletedContainers();
  } catch (error) {
    writeStatus('Container restore failed.', error.message);
  } finally {
    setBusy(restoreButton, false);
  }
});

elements.fileInput.addEventListener('change', () => {
  const file = elements.fileInput.files[0];
  if (!file) {
    updateOverview();
    return;
  }

  if (!elements.displayName.value.trim()) {
    elements.displayName.value = file.name.replace(/\.[^.]+$/, '');
  }

  updateOverview();
});

elements.containerTypeId.value = '';
elements.topCount.value = localStorage.getItem('topCount') || '20';

const persistContainerTypeId = () => {
  localStorage.setItem('containerTypeId', elements.containerTypeId.value.trim());
  updateOverview();
};

elements.containerTypeId.addEventListener('input', persistContainerTypeId);
elements.containerTypeId.addEventListener('change', persistContainerTypeId);

elements.topCount.addEventListener('change', () => {
  localStorage.setItem('topCount', elements.topCount.value.trim());
});

const initializeDefaults = () => {
  const storedContainerTypeId = localStorage.getItem('containerTypeId') || '';
  elements.containerTypeId.value = storedContainerTypeId || DEFAULT_CONTAINER_TYPE_ID;
  elements.containerTypeId.readOnly = false;
  elements.containerTypeId.title = 'Enter the SharePoint Embedded container type to target in this session.';
  if (elements.containerTypeId.value.trim()) {
    localStorage.setItem('containerTypeId', elements.containerTypeId.value.trim());
  }
  updateOverview();
};

const initializePage = async () => {
  await loadConfig();
  renderUsers();
  renderSelectedUsers();
  renderContainers();
  renderDeletedContainers();
  initializeDefaults();

  if (!state.appConfig.isConfigured) {
    setTokenState('idle', 'Configuration needed');
    setStatusSummary('Configuration needed', 'Set DATAVERSE_URL and restart the local proxy to connect this demo to Dataverse.');
    writeStatus('UI ready. Configure DATAVERSE_URL before loading users or container inventory.');
    return;
  }

  writeStatus('UI ready. Loading users and container inventory for the selected container type.');

  const results = await Promise.allSettled([
    loadUsers(),
    refreshContainers(),
    refreshDeletedContainers({ writeLog: false })
  ]);

  const failures = results.filter((result) => result.status === 'rejected');
  if (!failures.length) {
    setStatusSummary('Initial data loaded', 'Users, inventory, and recycle bin counts are ready.');
    return;
  }

  failures.forEach((result) => {
    writeStatus('Initial load failed.', result.reason?.message || String(result.reason || 'Unknown error'));
  });
};

initializePage().catch((error) => {
  writeStatus('Startup failed.', error.message);
});
document.addEventListener('click', (e) => {
  if (e.target.classList.contains('copyable')) {
    const text = e.target.textContent;
    navigator.clipboard.writeText(text).then(() => {
      const originalText = text;
      e.target.textContent = 'Copied!';
      setTimeout(() => e.target.textContent = originalText, 1500);
    });
  }
});
