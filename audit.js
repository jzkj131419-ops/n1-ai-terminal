const fs = require('fs');
const path = require('path');

const LOG_FILE = '/var/log/n1-audit.log';

function writeAudit({ operation, fileSize, modelName, uploaded, durationMs }) {
  const ts = new Date().toISOString();
  const line = `${ts} | ${operation} | ${fileSize}B | ${modelName} | ${uploaded ? 1 : 0} | ${durationMs}ms\n`;
  fs.appendFileSync(LOG_FILE, line, 'utf8');
}

module.exports = { writeAudit };
