const fs = require('fs');
const path = require('path');
const axios = require('axios');

/**
 * Download a Salesforce report in XLSX format.
 * @param {Object} auth - Authentication info from sfdcAuth.js
 * @param {string} auth.accessToken - OAuth access token
 * @param {string} auth.instanceUrl - Salesforce instance URL
 * @param {string} reportId - 18-character Salesforce report Id
 * @param {string} [destDir] - directory to save the file (defaults to current)
 * @returns {Promise<string>} - path to the downloaded XLSX file
 */
async function exportReport(auth, reportId, destDir = '.') {
  if (!auth || !auth.accessToken || !auth.instanceUrl) {
    throw new Error('Valid auth information required');
  }
  if (!/^[a-zA-Z0-9]{15,18}$/.test(reportId)) {
    throw new Error('Invalid report id');
  }

  const url = `${auth.instanceUrl}/services/data/v57.0/analytics/reports/${reportId}?export=1&enc=UTF-8&format=xlsx`;
  const { data } = await axios.get(url, {
    responseType: 'arraybuffer',
    headers: { Authorization: `Bearer ${auth.accessToken}` },
  });

  const outPath = path.join(destDir, `${reportId}.xlsx`);
  fs.writeFileSync(outPath, data);
  return outPath;
}

if (require.main === module) {
  const [token, instanceUrl, reportId] = process.argv.slice(2);
  exportReport({ accessToken: token, instanceUrl }, reportId)
    .then(fp => console.log(fp))
    .catch(err => {
      if (err.response && err.response.data) {
        console.error(err.response.data);
      } else {
        console.error(err.message);
      }
      process.exit(1);
    });
}

module.exports = exportReport;
