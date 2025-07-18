const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
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
 *   KEY_PASS         - Passphrase to decrypt the private key
 *   KEY_PBKDF2       - Set to "1" if the key was encrypted with openssl using
 *                      the -pbkdf2 option
 */
async function authenticate() {
  const debug = process.env.SFDC_AUTH_DEBUG === '1';
  const clientId = process.env.SFDC_CLIENT_ID;
  const username = process.env.SFDC_USERNAME || 'andbeder@salesforce.com';
  const keyPath = process.env.SFDC_PRIVATE_KEY;
  const keyPass = process.env.KEY_PASS;
  const tokenUrl = process.env.SFDC_TOKEN_URL ||
    'https://login.salesforce.com/services/oauth2/token';
  const oktaDomain = process.env.OKTA_DOMAIN || 'beder.okta.com';

  if (debug) {
    console.log('SFDC auth debug info:');
    console.log('  tokenUrl:', tokenUrl);
    console.log('  username:', username);
    console.log('  clientId:', clientId);
    console.log('  using PBKDF2:', process.env.KEY_PBKDF2 === '1');
  }

  if (!clientId || !keyPath) {
    throw new Error('SFDC_CLIENT_ID and SFDC_PRIVATE_KEY must be set');
  }

  if (!keyPass) {
    throw new Error('KEY_PASS must be set');
  }

  const usePbkdf2 = process.env.KEY_PBKDF2 === '1';
  const privateKey = decryptKey(
    fs.readFileSync(path.resolve(keyPath)),
    keyPass,
    usePbkdf2
  ).toString('utf8');

  const jwtToken = jwt.sign(
    {
      iss: clientId,
      sub: username,
      aud: `https://${oktaDomain}`,
    },
    privateKey,
    { algorithm: 'RS256', expiresIn: 3 * 60 }
  );

  if (debug) {
    const parts = jwtToken.split('.');
    const decode = str => JSON.parse(Buffer.from(str, 'base64').toString('utf8'));
    console.log('  jwt header:', decode(parts[0]));
    console.log('  jwt payload:', decode(parts[1]));
  }

  const params = new URLSearchParams();
  params.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
  params.append('assertion', jwtToken);

  if (debug) {
    console.log('  request body:', params.toString());
  }

  let data;
  try {
    const resp = await axios.post(tokenUrl, params.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    data = resp.data;
  } catch (err) {
    if (debug && err.response) {
      console.error('  status:', err.response.status);
      console.error('  response:', err.response.data);
    }
    throw err;
  }

  if (debug) {
    console.log('  accessToken:', data.access_token);
    console.log('  instanceUrl:', data.instance_url);
  }

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

function decryptKey(buf, pass, usePbkdf2) {
  const magic = Buffer.from('Salted__');
  if (buf.slice(0, magic.length).compare(magic) !== 0) {
    throw new Error('Invalid encrypted key file');
  }
  const salt = buf.slice(magic.length, magic.length + 8);
  const enc = buf.slice(magic.length + 8);

  let key, iv;
  if (usePbkdf2) {
    const derived = crypto.pbkdf2Sync(
      Buffer.from(pass, 'utf8'),
      salt,
      10000,
      48,
      'sha256'
    );
    key = derived.slice(0, 32);
    iv = derived.slice(32, 48);
  } else {
    ({ key, iv } = evpKdf(Buffer.from(pass, 'utf8'), salt, 32, 16));
  }
  const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
  return Buffer.concat([decipher.update(enc), decipher.final()]);
}

function evpKdf(password, salt, keyLen, ivLen) {
  let data = Buffer.alloc(0);
  let prev = Buffer.alloc(0);
  while (data.length < keyLen + ivLen) {
    const md5 = crypto.createHash('md5');
    md5.update(Buffer.concat([prev, password, salt]));
    prev = md5.digest();
    data = Buffer.concat([data, prev]);
  }
  return {
    key: data.slice(0, keyLen),
    iv: data.slice(keyLen, keyLen + ivLen),
  };
}
