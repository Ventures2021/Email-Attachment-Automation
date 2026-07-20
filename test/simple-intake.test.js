const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

const {
  buildSimpleIntakeRows,
  csvEscape,
  runSimpleIntake,
  toCsv,
} = require('../src/simple-intake');

test('csvEscape quotes values only when needed', () => {
  assert.equal(csvEscape('plain value'), 'plain value');
  assert.equal(csvEscape('value, with comma'), '"value, with comma"');
  assert.equal(csvEscape('value "with quote"'), '"value ""with quote"""');
});

test('buildSimpleIntakeRows creates Notion-ready rows from a local folder', async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-intake-'));
  const folder = path.join(tempDir, 'Magnolia Homes', 'Draw Requests');
  await fs.mkdir(folder, { recursive: true });
  await fs.writeFile(path.join(folder, 'Loan Package.pdf'), 'loan package');

  const rows = await buildSimpleIntakeRows(tempDir);

  assert.equal(rows.length, 1);
  assert.equal(rows[0].Document, 'Loan Package.pdf');
  assert.equal(rows[0].Category, 'Magnolia Homes');
  assert.equal(rows[0]['OneDrive Folder Path'], '/Magnolia Homes/Draw Requests');
  assert.equal(rows[0]['OneDrive File Link'], '');
  assert.equal(rows[0]['Storage Status'], 'Needs Link');
  assert.match(rows[0]['Processing Key'], /^local-file::Magnolia Homes\/Draw Requests\/Loan Package\.pdf::/);
});

test('runSimpleIntake writes a CSV file', async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-intake-'));
  await fs.writeFile(path.join(tempDir, 'Document.csv'), 'hello');
  const outputFile = path.join(tempDir, 'out', 'notion-import.csv');

  const result = await runSimpleIntake({ inputDirectory: tempDir, outputFile });
  const csv = await fs.readFile(outputFile, 'utf8');

  assert.equal(result.rowCount, 1);
  assert.match(csv, /^Document,Original Attachment Filename,OneDrive File Link,/);
  assert.match(csv, /Document\.csv/);
});

test('toCsv includes a trailing newline for spreadsheet imports', () => {
  assert.equal(toCsv([{ Document: 'A' }], ['Document']), 'Document\nA\n');
});
