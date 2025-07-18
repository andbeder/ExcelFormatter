const fs = require('fs');
const path = require('path');
const jwt = require('jsonwebtoken');
const axios = require('axios');

/**
 * Authenticate to Salesforce using the JWT bearer flow via Okta.
 *
 * Required environment variables:
 *   SFDC_CLIENT_ID   - Connected app client id (consumer key)
 *   SFDC_PRIVATE_KEY - Path to RSA private key (PEM format)
 *   SFDC_USERNAME    - Salesforce username (default andbeder@salesforce.com)
 *   SFDC_TOKEN_URL   - OAuth token endpoint. Defaults to
 *                      https://login.salesforce.com/services/oauth2/token
 *   OKTA_DOMAIN      - Okta domain used as the audience (default beder.okta.com)
 */
async function authenticate() {
  const clientId = process.env.SFDC_CLIENT_ID;
  const username = process.env.SFDC_USERNAME || 'andbeder@salesforce.com';
  const keyPath = process.env.SFDC_PRIVATE_KEY;
  const tokenUrl = process.env.SFDC_TOKEN_URL ||
    'https://login.salesforce.com/services/oauth2/token';
  const oktaDomain = process.env.OKTA_DOMAIN || 'beder.okta.com';

  if (!clientId || !keyPath) {
    throw new Error('SFDC_CLIENT_ID and SFDC_PRIVATE_KEY must be set');
  }

  const privateKey = fs.readFileSync(path.resolve(keyPath), 'utf8');

  const jwtToken = jwt.sign(
    {
      iss: clientId,
      sub: username,
      aud: `https://${oktaDomain}`,
    },
    privateKey,
    { algorithm: 'RS256', expiresIn: 3 * 60 }
  );

  const params = new URLSearchParams();
  params.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
  params.append('assertion', jwtToken);

  const { data } = await axios.post(tokenUrl, params.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  return {
    accessToken: data.access_token,
    instanceUrl: data.instance_url,
  };
}

if (require.main === module) {
  authenticate()
    .then(res => console.log(res))
    .catch(err => {
      if (err.response && err.response.data) {
        console.error(err.response.data);
      } else {
        console.error(err.message);
      }
      process.exit(1);
    });
}

module.exports = authenticate;
