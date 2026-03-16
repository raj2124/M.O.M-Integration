const path = require('path');
const dotenv = require('dotenv');

const rootDir = path.resolve(__dirname, '..');
const envPath = path.join(rootDir, '.env');

// Force ElectroReports and any shared modules it imports to resolve dotenv from
// the ElectroReports folder instead of the workspace root.
if (process.cwd() !== rootDir) {
  process.chdir(rootDir);
}

dotenv.config({ path: envPath });

function toBool(value, defaultValue = false) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  return String(value).toLowerCase() === 'true';
}

const host = process.env.ELECTROREPORTS_HOST || '127.0.0.1';
const port = Number.parseInt(process.env.ELECTROREPORTS_PORT || '3011', 10);

module.exports = {
  app: {
    rootDir,
    publicDir: path.join(rootDir, 'public'),
    generatedDir: path.join(rootDir, 'generated-pdfs'),
    dataFilePath: path.join(rootDir, 'data', 'reports.json'),
    host,
    port
  },
  zoho: {
    useMock: toBool(process.env.ZOHO_USE_MOCK, true)
  },
  envPath
};
