const fs = require('node:fs/promises');
const path = require('node:path');

const { sha256 } = require('./automation');

const DEFAULT_COLUMNS = [
  'Document',
  'Original Attachment Filename',
  'OneDrive File Link',
  'OneDrive Folder Path',
  'Local OneDrive Path',
  'Category',
  'Storage Status',
  'Link Access',
  'External Access Verified',
  'File Size',
  'File Hash',
  'Processing Key',
  'Extraction Notes',
];

function csvEscape(value) {
  const text = String(value ?? '');
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function toCsv(rows, columns = DEFAULT_COLUMNS) {
  const lines = [columns.map(csvEscape).join(',')];
  for (const row of rows) {
    lines.push(columns.map((column) => csvEscape(row[column])).join(','));
  }
  return `${lines.join('\n')}\n`;
}

async function listFilesRecursively(rootDirectory) {
  const entries = await fs.readdir(rootDirectory, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    if (entry.name.startsWith('.')) {
      continue;
    }

    const fullPath = path.join(rootDirectory, entry.name);
    if (entry.isDirectory()) {
      files.push(...await listFilesRecursively(fullPath));
    } else if (entry.isFile()) {
      files.push(fullPath);
    }
  }

  return files.sort((left, right) => left.localeCompare(right));
}

function buildSimpleProcessingKey(relativePath, fileHash) {
  return `local-file::${relativePath.replaceAll(path.sep, '/')}::${fileHash}`;
}

async function buildSimpleIntakeRows(inputDirectory) {
  const rootDirectory = path.resolve(inputDirectory);
  const files = await listFilesRecursively(rootDirectory);
  const rows = [];

  for (const filePath of files) {
    const buffer = await fs.readFile(filePath);
    const fileHash = sha256(buffer);
    const relativePath = path.relative(rootDirectory, filePath);
    const relativeFolder = path.dirname(relativePath) === '.' ? '' : path.dirname(relativePath);
    const folderParts = relativeFolder.split(path.sep).filter(Boolean);
    const filename = path.basename(filePath);

    rows.push({
      Category: folderParts[0] || 'Uncategorized',
      Document: filename,
      'External Access Verified': false,
      'Extraction Notes': 'Simple intake: upload/sync file in OneDrive, paste the sharing link into Notion, then mark verified after testing externally.',
      'File Hash': fileHash,
      'File Size': buffer.length,
      'Link Access': 'Needs Link',
      'Local OneDrive Path': filePath,
      'OneDrive File Link': '',
      'OneDrive Folder Path': relativeFolder ? `/${relativeFolder.replaceAll(path.sep, '/')}` : '/',
      'Original Attachment Filename': filename,
      'Processing Key': buildSimpleProcessingKey(relativePath, fileHash),
      'Storage Status': 'Needs Link',
    });
  }

  return rows;
}

async function runSimpleIntake({ inputDirectory, outputFile }) {
  if (!inputDirectory) {
    throw new Error('Usage: npm run simple -- <OneDrive folder path> [output.csv]');
  }

  const rows = await buildSimpleIntakeRows(inputDirectory);
  const outputPath = path.resolve(outputFile || 'notion-import.csv');
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, toCsv(rows));
  return { outputPath, rowCount: rows.length };
}

async function runSimpleIntakeCli(argv) {
  const [, , inputDirectory, outputFile] = argv;
  const result = await runSimpleIntake({ inputDirectory, outputFile });
  console.log(`Created ${result.outputPath} with ${result.rowCount} document row(s).`);
  console.log('Import the CSV into Notion, then paste OneDrive sharing links into the blank OneDrive File Link column.');
}

module.exports = {
  DEFAULT_COLUMNS,
  buildSimpleIntakeRows,
  buildSimpleProcessingKey,
  csvEscape,
  runSimpleIntake,
  runSimpleIntakeCli,
  toCsv,
};
