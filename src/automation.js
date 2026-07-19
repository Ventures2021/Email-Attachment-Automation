const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');

const DEFAULT_NOTION_FIELDS = {
  category: 'Category',
  externalAccessVerified: 'External Access Verified',
  extractionNotes: 'Extraction Notes',
  fileHash: 'File Hash',
  fileSize: 'File Size',
  folderPath: 'OneDrive Folder Path',
  linkAccess: 'Link Access',
  linkExpiration: 'Link Expiration',
  linkVerifiedAt: 'Link Verified At',
  messageId: 'Email Message ID',
  oneDriveLink: 'OneDrive File Link',
  originalFilename: 'Original Attachment Filename',
  processingKey: 'Processing Key',
  received: 'Received',
  sender: 'Sender',
  storageStatus: 'Storage Status',
  subject: 'Email Subject',
  sourceFolder: 'Outlook Source Folder',
};

function sha256(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function ensureDirectory(filePath) {
  return fs.mkdir(path.dirname(filePath), { recursive: true });
}

function sanitizePathSegment(value) {
  return String(value || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled';
}

function normalizeOneDrivePath(basePath, folderSegments = []) {
  const pieces = [basePath, ...folderSegments.map(sanitizePathSegment)]
    .join('/')
    .split('/')
    .filter(Boolean);
  return `/${pieces.join('/')}`;
}

function encodeGraphPath(value) {
  return value
    .split('/')
    .map((segment) => encodeURIComponent(segment))
    .join('/');
}

function buildProcessingKey({ attachmentId, fileHash, filename, messageId }) {
  return [messageId || 'missing-message-id', attachmentId, filename, fileHash].join('::');
}

function versionFilename(filename, stamp) {
  const extension = path.extname(filename);
  const base = path.basename(filename, extension);
  return `${base} (${stamp})${extension}`;
}

function buildVersionStamp(receivedAt) {
  return new Date(receivedAt || Date.now()).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');
}

function detectPasswordProtectedAttachment(filename, buffer) {
  const lower = String(filename || '').toLowerCase();
  if (lower.endsWith('.pdf')) {
    return buffer.subarray(0, Math.min(buffer.length, 4096)).toString('latin1').includes('/Encrypt');
  }

  const isZipFamily = ['.zip', '.docx', '.xlsx', '.pptx'].some((extension) => lower.endsWith(extension));
  if (isZipFamily && buffer.length >= 8 && buffer.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
    const generalPurposeBitFlag = buffer.readUInt16LE(6);
    return (generalPurposeBitFlag & 0x0001) === 0x0001;
  }

  return false;
}

function maskSecrets(headers = {}) {
  const result = { ...headers };
  for (const name of Object.keys(result)) {
    if (/authorization/i.test(name)) {
      result[name] = '***';
    }
  }
  return result;
}

async function requestJson(url, { body, headers = {}, method = 'GET', okStatuses = [200], parseAsText = false } = {}) {
  const response = await fetch(url, {
    body: body === undefined ? undefined : JSON.stringify(body),
    headers: {
      ...headers,
      ...(body === undefined ? {} : { 'content-type': 'application/json' }),
    },
    method,
  });

  if (!okStatuses.includes(response.status)) {
    const errorBody = await response.text();
    throw new Error(`Request failed (${response.status}) ${url}: ${errorBody}`);
  }

  if (parseAsText) {
    return response.text();
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

class JsonStateStore {
  constructor(filePath) {
    this.filePath = filePath;
    this.loaded = false;
    this.recordsByFolderAndFilename = new Map();
    this.recordsByProcessingKey = new Map();
    this.state = { records: [] };
  }

  async load() {
    if (this.loaded) {
      return;
    }

    try {
      const content = await fs.readFile(this.filePath, 'utf8');
      this.state = JSON.parse(content);
    } catch (error) {
      if (error.code !== 'ENOENT') {
        throw error;
      }
      this.state = { records: [] };
    }

    this.recordsByFolderAndFilename.clear();
    this.recordsByProcessingKey.clear();
    for (const record of this.state.records) {
      this.recordsByProcessingKey.set(record.processingKey, record);
      this.recordsByFolderAndFilename.set(`${record.folderPath}::${record.storedFilename}`, record);
    }
    this.loaded = true;
  }

  hasProcessingKey(processingKey) {
    return this.recordsByProcessingKey.has(processingKey);
  }

  findByFolderAndFilename(folderPath, filename) {
    return this.recordsByFolderAndFilename.get(`${folderPath}::${filename}`) || null;
  }

  add(record) {
    this.state.records.push(record);
    this.recordsByProcessingKey.set(record.processingKey, record);
    this.recordsByFolderAndFilename.set(`${record.folderPath}::${record.storedFilename}`, record);
  }

  async save() {
    await ensureDirectory(this.filePath);
    await fs.writeFile(this.filePath, `${JSON.stringify(this.state, null, 2)}\n`);
  }
}

class AuditLog {
  constructor(filePath) {
    this.filePath = filePath;
  }

  async append(entry) {
    await ensureDirectory(this.filePath);
    await fs.appendFile(this.filePath, `${JSON.stringify({ timestamp: new Date().toISOString(), ...entry })}\n`);
  }
}

class MicrosoftGraphClient {
  constructor(config) {
    this.config = config;
    this.graphBaseUrl = config.graphBaseUrl;
    this.cachedToken = null;
    this.cachedTokenExpiresAt = 0;
  }

  async getAccessToken() {
    if (this.cachedToken && Date.now() + 60_000 < this.cachedTokenExpiresAt) {
      return this.cachedToken;
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.config.microsoft.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.config.microsoft.clientId,
      client_secret: this.config.microsoft.clientSecret,
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default',
    });

    const response = await fetch(tokenUrl, {
      body,
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      method: 'POST',
    });

    if (!response.ok) {
      throw new Error(`Microsoft token request failed (${response.status}): ${await response.text()}`);
    }

    const payload = await response.json();
    this.cachedToken = payload.access_token;
    this.cachedTokenExpiresAt = Date.now() + (payload.expires_in * 1000);
    return this.cachedToken;
  }

  async request(pathname, { body, method = 'GET', okStatuses = [200], raw = false } = {}) {
    const accessToken = await this.getAccessToken();
    const response = await fetch(`${this.graphBaseUrl}${pathname}`, {
      body: body === undefined || body instanceof Buffer ? body : JSON.stringify(body),
      headers: {
        authorization: 'Bearer ' + accessToken,
        ...(body instanceof Buffer ? { 'content-type': 'application/octet-stream' } : {}),
        ...(body === undefined || body instanceof Buffer ? {} : { 'content-type': 'application/json' }),
      },
      method,
    });

    if (!okStatuses.includes(response.status)) {
      const errorBody = await response.text();
      throw new Error(`Graph request failed (${response.status}) ${pathname}: ${errorBody}`);
    }

    if (raw) {
      return response;
    }

    if (response.status === 204) {
      return null;
    }

    return response.json();
  }

  async paginate(pathname) {
    const items = [];
    let nextUrl = `${this.graphBaseUrl}${pathname}`;
    while (nextUrl) {
      const accessToken = await this.getAccessToken();
      const response = await fetch(nextUrl, {
        headers: { authorization: 'Bearer ' + accessToken },
      });
      if (!response.ok) {
        throw new Error(`Graph pagination failed (${response.status}) ${nextUrl}: ${await response.text()}`);
      }
      const payload = await response.json();
      items.push(...(payload.value || []));
      nextUrl = payload['@odata.nextLink'] || null;
    }
    return items;
  }

  driveBasePath() {
    return this.config.oneDrive.driveId
      ? `/drives/${this.config.oneDrive.driveId}`
      : `/users/${encodeURIComponent(this.config.mailboxUserId)}/drive`;
  }

  async getMailFolder(folderId) {
    return this.request(`/users/${encodeURIComponent(this.config.mailboxUserId)}/mailFolders/${encodeURIComponent(folderId)}?$select=id,displayName,parentFolderId`);
  }

  async listChildFolders(folderId) {
    return this.paginate(`/users/${encodeURIComponent(this.config.mailboxUserId)}/mailFolders/${encodeURIComponent(folderId)}/childFolders?$select=id,displayName,parentFolderId`);
  }

  async listMonitoredFolders() {
    const folders = [];
    const walk = async (folderId, pathSegments) => {
      const folder = await this.getMailFolder(folderId);
      const nextSegments = [...pathSegments, folder.displayName];
      folders.push({
        id: folder.id,
        pathSegments: nextSegments,
        relativePath: nextSegments.map(sanitizePathSegment).join('/'),
      });
      const children = await this.listChildFolders(folder.id);
      for (const child of children) {
        await walk(child.id, nextSegments);
      }
    };

    for (const folderId of this.config.outlookRootFolderIds) {
      await walk(folderId, []);
    }

    return folders;
  }

  async listMessages(folderId) {
    return this.paginate(`/users/${encodeURIComponent(this.config.mailboxUserId)}/mailFolders/${encodeURIComponent(folderId)}/messages?$select=id,internetMessageId,subject,receivedDateTime,from,hasAttachments`);
  }

  async listAttachments(messageId) {
    const items = await this.paginate(`/users/${encodeURIComponent(this.config.mailboxUserId)}/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,size,contentType,isInline,@odata.type`);
    return items.filter((item) => item.isInline !== true && item['@odata.type'] !== '#microsoft.graph.itemAttachment');
  }

  async downloadAttachment(messageId, attachmentId) {
    const response = await this.request(`/users/${encodeURIComponent(this.config.mailboxUserId)}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`, {
      raw: true,
    });
    return Buffer.from(await response.arrayBuffer());
  }

  async findDriveItemByPath(filePath) {
    const encodedPath = encodeGraphPath(filePath.replace(/^\//, ''));
    try {
      return await this.request(`${this.driveBasePath()}/root:/${encodedPath}`);
    } catch (error) {
      if (String(error.message).includes('Graph request failed (404)')) {
        return null;
      }
      throw error;
    }
  }

  async ensureDriveFolder(folderPath) {
    const segments = folderPath.split('/').filter(Boolean);
    let parentId = 'root';
    let cumulative = '';
    for (const segment of segments) {
      cumulative = `${cumulative}/${segment}`;
      const existing = await this.findDriveItemByPath(cumulative);
      if (existing) {
        parentId = existing.id;
        continue;
      }

      const created = await this.request(`${this.driveBasePath()}/items/${parentId}/children`, {
        body: {
          '@microsoft.graph.conflictBehavior': 'fail',
          folder: {},
          name: segment,
        },
        method: 'POST',
        okStatuses: [201],
      });
      parentId = created.id;
    }
  }

  async uploadFile(folderPath, filename, buffer) {
    const fullPath = `${folderPath.replace(/\/$/, '')}/${filename}`;
    const encodedPath = encodeGraphPath(fullPath.replace(/^\//, ''));
    return this.request(`${this.driveBasePath()}/root:/${encodedPath}:/content`, {
      body: buffer,
      method: 'PUT',
      okStatuses: [200, 201],
    });
  }

  async createAnonymousViewLink(itemId, expirationDateTime) {
    const payload = await this.request(`${this.driveBasePath()}/items/${itemId}/createLink`, {
      body: {
        expirationDateTime,
        scope: 'anonymous',
        type: 'view',
      },
      method: 'POST',
      okStatuses: [200, 201],
    });

    return {
      expirationDateTime,
      url: payload.link?.webUrl,
    };
  }

  async verifyAnonymousLink(url) {
    try {
      const response = await fetch(url, {
        headers: { 'user-agent': 'email-attachment-automation/1.0' },
        redirect: 'manual',
      });
      const location = response.headers.get('location') || '';
      const redirectsToMicrosoftLogin = /login\.microsoftonline\.com|login\.live\.com/i.test(location);
      return response.status < 400 && !redirectsToMicrosoftLogin;
    } catch (error) {
      return false;
    }
  }
}

class NotionClient {
  constructor(config) {
    this.config = config;
    this.databaseSchema = null;
  }

  headers() {
    return {
      authorization: 'Bearer ' + this.config.notion.token,
      'notion-version': '2022-06-28',
    };
  }

  async request(pathname, options = {}) {
    return requestJson(`${this.config.notion.baseUrl}${pathname}`, {
      ...options,
      headers: {
        ...this.headers(),
        ...(options.headers || {}),
      },
    });
  }

  async getDatabaseSchema() {
    if (this.databaseSchema) {
      return this.databaseSchema;
    }
    this.databaseSchema = await this.request(`/databases/${this.config.notion.databaseId}`);
    return this.databaseSchema;
  }

  getTitlePropertyName(schema) {
    if (this.config.notion.titleProperty) {
      return this.config.notion.titleProperty;
    }

    return Object.entries(schema.properties || {}).find(([, property]) => property.type === 'title')?.[0] || null;
  }

  buildTextProperty(type, value) {
    const content = String(value || '').slice(0, 2000);
    if (!content) {
      return { [type]: [] };
    }
    return {
      [type]: [{ type: 'text', text: { content } }],
    };
  }

  buildPropertyValue(type, value) {
    if (value === undefined) {
      return null;
    }

    switch (type) {
      case 'title':
        return this.buildTextProperty('title', value);
      case 'rich_text':
        return this.buildTextProperty('rich_text', value);
      case 'url':
        return { url: value ? String(value) : null };
      case 'checkbox':
        return { checkbox: Boolean(value) };
      case 'date':
        return { date: value ? { start: new Date(value).toISOString() } : null };
      case 'number':
        return { number: Number.isFinite(value) ? value : null };
      case 'select':
        return { select: value ? { name: String(value).slice(0, 100) } : null };
      case 'status':
        return { status: value ? { name: String(value).slice(0, 100) } : null };
      case 'email':
        return { email: /\S+@\S+\.\S+/.test(String(value || '')) ? String(value) : null };
      default:
        return null;
    }
  }

  buildQueryFilter(schema, propertyName, value) {
    const property = schema.properties?.[propertyName];
    if (!property || value === undefined || value === null || value === '') {
      return null;
    }

    if (property.type === 'title') {
      return { property: propertyName, title: { equals: String(value).slice(0, 2000) } };
    }
    if (property.type === 'rich_text') {
      return { property: propertyName, rich_text: { equals: String(value).slice(0, 2000) } };
    }
    if (property.type === 'url') {
      return { property: propertyName, url: { equals: String(value) } };
    }
    if (property.type === 'number' && Number.isFinite(value)) {
      return { property: propertyName, number: { equals: value } };
    }

    return null;
  }

  async queryFirst(filter) {
    const payload = await this.request(`/databases/${this.config.notion.databaseId}/query`, {
      body: { filter, page_size: 1 },
      method: 'POST',
    });
    return payload.results?.[0] || null;
  }

  async findExistingPage(document) {
    const schema = await this.getDatabaseSchema();
    const candidates = [
      [this.config.notion.fields.processingKey, document.processingKey],
      [this.config.notion.fields.messageId, document.messageId],
      [this.config.notion.fields.fileHash, document.fileHash],
    ];

    for (const [propertyName, value] of candidates) {
      const filter = this.buildQueryFilter(schema, propertyName, value);
      if (!filter) {
        continue;
      }
      const page = await this.queryFirst(filter);
      if (page) {
        return page;
      }
    }

    return null;
  }

  async buildProperties(document) {
    const schema = await this.getDatabaseSchema();
    const titlePropertyName = this.getTitlePropertyName(schema);
    const fieldValues = {
      [titlePropertyName]: document.originalFilename,
      [this.config.notion.fields.oneDriveLink]: document.oneDriveLink,
      [this.config.notion.fields.folderPath]: document.folderPath,
      [this.config.notion.fields.category]: document.category,
      [this.config.notion.fields.storageStatus]: document.storageStatus,
      [this.config.notion.fields.linkAccess]: document.linkAccess,
      [this.config.notion.fields.externalAccessVerified]: document.externalAccessVerified,
      [this.config.notion.fields.linkVerifiedAt]: document.linkVerifiedAt,
      [this.config.notion.fields.linkExpiration]: document.linkExpiration,
      [this.config.notion.fields.sender]: document.sender,
      [this.config.notion.fields.received]: document.received,
      [this.config.notion.fields.fileSize]: document.fileSize,
      [this.config.notion.fields.fileHash]: document.fileHash,
      [this.config.notion.fields.processingKey]: document.processingKey,
      [this.config.notion.fields.extractionNotes]: document.extractionNotes,
      [this.config.notion.fields.messageId]: document.messageId,
      [this.config.notion.fields.originalFilename]: document.originalFilename,
      [this.config.notion.fields.subject]: document.subject,
      [this.config.notion.fields.sourceFolder]: document.sourceFolder,
    };

    const properties = {};
    for (const [propertyName, value] of Object.entries(fieldValues)) {
      if (!propertyName || !schema.properties?.[propertyName]) {
        continue;
      }
      const propertyValue = this.buildPropertyValue(schema.properties[propertyName].type, value);
      if (propertyValue) {
        properties[propertyName] = propertyValue;
      }
    }

    return properties;
  }

  async upsertDocument(document) {
    const properties = await this.buildProperties(document);
    const existingPage = await this.findExistingPage(document);
    if (existingPage) {
      const updated = await this.request(`/pages/${existingPage.id}`, {
        body: { properties },
        method: 'PATCH',
      });
      return { action: 'updated', page: updated };
    }

    const created = await this.request('/pages', {
      body: {
        parent: { database_id: this.config.notion.databaseId },
        properties,
      },
      method: 'POST',
      okStatuses: [200],
    });
    return { action: 'created', page: created };
  }
}

class AttachmentAutomation {
  constructor({ auditLog, config, graphClient, notionClient, stateStore }) {
    this.auditLog = auditLog;
    this.config = config;
    this.graphClient = graphClient;
    this.notionClient = notionClient;
    this.stateStore = stateStore;
    this.running = false;
  }

  async start() {
    if (this.running) {
      return;
    }

    this.running = true;
    while (this.running) {
      await this.runOnce();
      await sleep(this.config.pollIntervalMs);
    }
  }

  stop() {
    this.running = false;
  }

  async runOnce() {
    await this.stateStore.load();
    const folders = await this.graphClient.listMonitoredFolders();

    for (const folder of folders) {
      const targetFolderPath = normalizeOneDrivePath(this.config.oneDrive.rootPath, folder.pathSegments);
      await this.graphClient.ensureDriveFolder(targetFolderPath);
      const messages = await this.graphClient.listMessages(folder.id);

      for (const message of messages) {
        if (!message.hasAttachments) {
          continue;
        }
        await this.processMessage(folder, targetFolderPath, message);
      }
    }

    await this.stateStore.save();
  }

  async processMessage(folder, targetFolderPath, message) {
    const attachments = await this.graphClient.listAttachments(message.id);
    for (const attachment of attachments) {
      try {
        await this.processAttachment(folder, targetFolderPath, message, attachment);
      } catch (error) {
        await this.auditLog.append({
          attachmentId: attachment.id,
          filename: attachment.name,
          messageId: message.internetMessageId || message.id,
          reason: error.message,
          status: 'failed',
        });
      }
    }
  }

  async processAttachment(folder, targetFolderPath, message, attachment) {
    const attachmentBuffer = await this.graphClient.downloadAttachment(message.id, attachment.id);
    if (detectPasswordProtectedAttachment(attachment.name, attachmentBuffer)) {
      await this.auditLog.append({
        attachmentId: attachment.id,
        filename: attachment.name,
        messageId: message.internetMessageId || message.id,
        status: 'password-protected',
      });
      return;
    }

    const fileHash = sha256(attachmentBuffer);
    const messageId = message.internetMessageId || message.id;
    const processingKey = buildProcessingKey({
      attachmentId: attachment.id,
      fileHash,
      filename: attachment.name,
      messageId,
    });

    if (this.stateStore.hasProcessingKey(processingKey)) {
      await this.auditLog.append({
        attachmentId: attachment.id,
        filename: attachment.name,
        messageId,
        processingKey,
        status: 'duplicate',
      });
      return;
    }

    const uploadDecision = await this.resolveUploadTarget({
      fileHash,
      filename: attachment.name,
      folderPath: targetFolderPath,
      receivedAt: message.receivedDateTime,
    });

    if (uploadDecision.duplicate) {
      await this.auditLog.append({
        filename: attachment.name,
        folderPath: targetFolderPath,
        messageId,
        processingKey,
        status: 'duplicate',
      });
      return;
    }

    const uploadedItem = await this.graphClient.uploadFile(targetFolderPath, uploadDecision.filename, attachmentBuffer);
    const expirationDateTime = new Date(Date.now() + (this.config.publicLinkExpiryDays * 24 * 60 * 60 * 1000)).toISOString();
    const link = await this.graphClient.createAnonymousViewLink(uploadedItem.id, expirationDateTime);
    const externalAccessVerified = await this.graphClient.verifyAnonymousLink(link.url);
    const note = uploadDecision.versioned
      ? `Filename collision detected. Stored as ${uploadDecision.filename} and flagged as a new version.`
      : 'Stored automatically from Outlook attachment.';
    const notionResult = await this.notionClient.upsertDocument({
      category: folder.pathSegments[0] || 'Uncategorized',
      externalAccessVerified,
      extractionNotes: note,
      fileHash,
      fileSize: attachment.size || attachmentBuffer.length,
      folderPath: targetFolderPath,
      linkAccess: externalAccessVerified ? 'Anonymous View' : 'Restricted',
      linkExpiration: expirationDateTime,
      linkVerifiedAt: new Date().toISOString(),
      messageId,
      oneDriveLink: link.url,
      originalFilename: attachment.name,
      processingKey,
      received: message.receivedDateTime,
      sender: message.from?.emailAddress?.address || message.from?.emailAddress?.name || 'Unknown',
      sourceFolder: folder.relativePath,
      storageStatus: uploadDecision.versioned ? 'Versioned' : 'Stored',
      subject: message.subject || '(no subject)',
    });

    const record = {
      externalAccessVerified,
      fileHash,
      folderPath: targetFolderPath,
      notionAction: notionResult.action,
      notionPageId: notionResult.page.id,
      processingKey,
      storedFilename: uploadDecision.filename,
    };
    this.stateStore.add(record);

    await this.auditLog.append({
      filename: uploadDecision.filename,
      folderPath: targetFolderPath,
      messageId,
      notionAction: notionResult.action,
      originalFilename: attachment.name,
      processingKey,
      status: externalAccessVerified ? 'processed' : 'restricted-link',
    });
  }

  async resolveUploadTarget({ fileHash, filename, folderPath, receivedAt }) {
    const stamp = buildVersionStamp(receivedAt);
    let candidate = filename;
    let versioned = false;
    let attempt = 0;

    while (true) {
      const existingRecord = this.stateStore.findByFolderAndFilename(folderPath, candidate);
      if (existingRecord) {
        if (candidate === filename && existingRecord.fileHash === fileHash) {
          return { duplicate: true, filename };
        }
      } else {
        const remoteItem = await this.graphClient.findDriveItemByPath(`${folderPath}/${candidate}`);
        if (!remoteItem) {
          return { duplicate: false, filename: candidate, versioned };
        }
      }

      attempt += 1;
      versioned = true;
      const suffix = attempt === 1 ? stamp : `${stamp}-${attempt}`;
      candidate = versionFilename(filename, suffix);
    }
  }
}

function required(name, env) {
  if (!env[name]) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return env[name];
}

function loadConfig(env = process.env) {
  return {
    auditLogFile: env.AUDIT_LOG_FILE || path.resolve(process.cwd(), 'data/audit-log.jsonl'),
    graphBaseUrl: env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
    mailboxUserId: required('MAILBOX_USER_ID', env),
    microsoft: {
      clientId: required('MICROSOFT_CLIENT_ID', env),
      clientSecret: required('MICROSOFT_CLIENT_SECRET', env),
      tenantId: required('MICROSOFT_TENANT_ID', env),
    },
    notion: {
      baseUrl: env.NOTION_BASE_URL || 'https://api.notion.com/v1',
      databaseId: required('NOTION_DATABASE_ID', env),
      fields: DEFAULT_NOTION_FIELDS,
      titleProperty: env.NOTION_TITLE_PROPERTY,
      token: required('NOTION_TOKEN', env),
    },
    oneDrive: {
      driveId: env.ONEDRIVE_DRIVE_ID,
      rootPath: env.ONEDRIVE_ROOT_PATH || '/Email Attachments',
    },
    outlookRootFolderIds: required('OUTLOOK_ROOT_FOLDER_IDS', env).split(',').map((value) => value.trim()).filter(Boolean),
    pollIntervalMs: Number(env.POLL_INTERVAL_MS || 300000),
    publicLinkExpiryDays: Number(env.PUBLIC_LINK_EXPIRY_DAYS || 90),
    runOnce: /^true$/i.test(env.RUN_ONCE || ''),
    stateFile: env.STATE_FILE || path.resolve(process.cwd(), 'data/processing-state.json'),
  };
}

module.exports = {
  AttachmentAutomation,
  AuditLog,
  DEFAULT_NOTION_FIELDS,
  JsonStateStore,
  MicrosoftGraphClient,
  NotionClient,
  buildProcessingKey,
  detectPasswordProtectedAttachment,
  loadConfig,
  maskSecrets,
  normalizeOneDrivePath,
  sha256,
  versionFilename,
};
