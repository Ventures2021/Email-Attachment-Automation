const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

const {
  AttachmentAutomation,
  AuditLog,
  JsonStateStore,
  buildProcessingKey,
  detectPasswordProtectedAttachment,
  normalizeOneDrivePath,
} = require('../src/automation');

test('buildProcessingKey keeps all dedupe identifiers', () => {
  assert.equal(
    buildProcessingKey({
      attachmentId: 'attachment-1',
      fileHash: 'abc123',
      filename: 'Loan Package.pdf',
      messageId: 'message-1',
    }),
    'message-1::attachment-1::Loan Package.pdf::abc123',
  );
});

test('detectPasswordProtectedAttachment identifies encrypted PDFs and Office zip containers', () => {
  const pdfBuffer = Buffer.from('%PDF-1.7 some content /Encrypt trailer', 'latin1');
  assert.equal(detectPasswordProtectedAttachment('secure.pdf', pdfBuffer), true);

  const zipBuffer = Buffer.alloc(8);
  zipBuffer.writeUInt32LE(0x04034b50, 0);
  zipBuffer.writeUInt16LE(1, 6);
  assert.equal(detectPasswordProtectedAttachment('secure.docx', zipBuffer), true);
  assert.equal(detectPasswordProtectedAttachment('plain.txt', Buffer.from('hello')), false);
});

test('normalizeOneDrivePath preserves the Outlook folder hierarchy', () => {
  assert.equal(
    normalizeOneDrivePath('/Email Attachments', ['Magnolia Homes', 'Draw Requests']),
    '/Email Attachments/Magnolia Homes/Draw Requests',
  );
});

test('AttachmentAutomation versions filename collisions and logs duplicates', async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'email-attachment-automation-'));
  const stateStore = new JsonStateStore(path.join(tempDir, 'state.json'));
  const auditLog = new AuditLog(path.join(tempDir, 'audit.jsonl'));

  const folder = {
    id: 'folder-1',
    pathSegments: ['Magnolia Homes'],
    relativePath: 'Magnolia Homes',
  };
  const messageOne = {
    from: { emailAddress: { address: 'sender@example.com' } },
    hasAttachments: true,
    id: 'message-graph-1',
    internetMessageId: 'message-1',
    receivedDateTime: '2026-07-19T22:00:00.000Z',
    subject: 'Initial package',
  };
  const messageTwo = {
    ...messageOne,
    id: 'message-graph-2',
    internetMessageId: 'message-2',
    receivedDateTime: '2026-07-20T22:00:00.000Z',
    subject: 'Updated package',
  };

  const graphClient = {
    createAnonymousViewLink: async (itemId, expirationDateTime) => ({ expirationDateTime, url: `https://example.com/${itemId}` }),
    downloadAttachment: async (_messageId, attachmentId) => Buffer.from(attachmentId),
    ensureDriveFolder: async () => {},
    findDriveItemByPath: async () => null,
    listAttachments: async (messageGraphId) => [{
      id: messageGraphId === 'message-graph-1' ? 'attachment-a' : 'attachment-b',
      name: 'Loan Package.pdf',
      size: 10,
    }],
    listMessages: async (folderId) => (folderId === 'folder-1' ? [messageOne, messageTwo, messageTwo] : []),
    listMonitoredFolders: async () => [folder],
    uploadFile: async (_folderPath, filename) => ({ id: filename }),
    verifyAnonymousLink: async () => true,
  };

  const notionClient = {
    upsertDocument: async () => ({ action: 'created', page: { id: 'page-1' } }),
  };

  const automation = new AttachmentAutomation({
    auditLog,
    config: {
      oneDrive: { rootPath: '/Email Attachments' },
      pollIntervalMs: 1,
      publicLinkExpiryDays: 90,
    },
    graphClient,
    notionClient,
    stateStore,
  });

  await automation.runOnce();

  const savedState = JSON.parse(await fs.readFile(path.join(tempDir, 'state.json'), 'utf8'));
  assert.equal(savedState.records.length, 2);
  assert.equal(savedState.records[0].storedFilename, 'Loan Package.pdf');
  assert.match(savedState.records[1].storedFilename, /^Loan Package \(\d{8}T\d{6}Z\)\.pdf$/);

  const auditEntries = (await fs.readFile(path.join(tempDir, 'audit.jsonl'), 'utf8'))
    .trim()
    .split('\n')
    .map((line) => JSON.parse(line));
  assert.equal(auditEntries.filter((entry) => entry.status === 'processed').length, 2);
  assert.equal(auditEntries.filter((entry) => entry.status === 'duplicate').length, 1);
});
