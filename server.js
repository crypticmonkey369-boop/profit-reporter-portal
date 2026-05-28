require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { google } = require('googleapis');
const { runReportForDate } = require('./src/reporter');
const { getGoogleAuthClient } = require('./src/config-sheet');
const { DateTime } = require('luxon');

const app = express();
const PORT = process.env.PORT || 3000;
const TOKENS_PATH = path.join(__dirname, 'tokens.json');

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser(process.env.COOKIE_SECRET || 'fallback-cookie-secret'));

// Route protection middleware (Authentication Gate)
function requireAuth(req, res, next) {
  if (req.signedCookies && req.signedCookies.auth === 'true') {
    next();
  } else {
    if (req.path.startsWith('/api/') || req.path.startsWith('/oauth/')) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
    res.redirect('/login.html');
  }
}

// Serve login page and public assets
app.get('/login.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public/login.html'));
});
app.get('/style.css', (req, res) => {
  res.sendFile(path.join(__dirname, 'public/style.css'));
});

// Protect all dashboard routes
app.use('/index.html', requireAuth);
app.get('/index.html', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, 'public/index.html'));
});
app.get('/', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, 'public/index.html'));
});

// API Login endpoint
app.post('/api/login', (req, res) => {
  const { password } = req.body;
  const portalPassword = process.env.PORTAL_PASSWORD || 'admin';
  
  if (password === portalPassword) {
    res.cookie('auth', 'true', {
      signed: true,
      httpOnly: true,
      maxAge: 24 * 60 * 60 * 1000,
      sameSite: 'lax',
      secure: process.env.NODE_ENV === 'production'
    });
    res.json({ success: true });
  } else {
    res.status(401).json({ success: false, error: 'Invalid password' });
  }
});

// API Logout endpoint
app.post('/api/logout', (req, res) => {
  res.clearCookie('auth');
  res.json({ success: true });
});

// Helper to read tokens safely
function readTokens() {
  if (!fs.existsSync(TOKENS_PATH)) {
    fs.writeFileSync(TOKENS_PATH, JSON.stringify({ google: {}, shopify: {}, meta: {}, klaviyo: {} }), 'utf8');
  }
  const data = JSON.parse(fs.readFileSync(TOKENS_PATH, 'utf8'));
  if (!data.google) data.google = {};
  return data;
}

// Helper to write tokens safely
function writeTokens(tokens) {
  fs.writeFileSync(TOKENS_PATH, JSON.stringify(tokens, null, 2), 'utf8');
  syncTokensToSheet(tokens).catch(err => {
    console.error('[SYNC ERROR] Async token sync to sheet failed:', err.message);
  });
}

// Sync tokens from sheet to local tokens.json
async function syncTokensFromSheet() {
  console.log('[SYNC] Restoring tokens from Google Sheet sys_tokens tab...');
  try {
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId || spreadsheetId === 'mock_google_sheet_id') {
      console.log('[SYNC] No Google Sheet configured or set to mock. Skipping remote token restoration.');
      return;
    }

    let authClient;
    try {
      authClient = getGoogleAuthClient();
    } catch (e) {
      console.log('[SYNC] Google auth client could not be initialized yet. Skipping remote token restoration:', e.message);
      return;
    }

    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'sys_tokens!A1:B1',
      });
      const rows = response.data.values;
      if (rows && rows[0] && rows[0][1]) {
        const jsonStr = rows[0][1];
        const tokens = JSON.parse(jsonStr);
        fs.writeFileSync(TOKENS_PATH, JSON.stringify(tokens, null, 2), 'utf8');
        console.log('[SYNC] Successfully restored tokens from Google Sheet.');
      } else {
        console.log('[SYNC] sys_tokens tab found but empty.');
      }
    } catch (err) {
      if (err.message.includes('NOT_FOUND') || err.message.includes('parse')) {
        console.log('[SYNC] sys_tokens tab not found. Creating it...');
        await createSysTokensTab(sheets, spreadsheetId);
      } else {
        console.error('[SYNC ERROR] Failed to read sys_tokens tab:', err.message);
      }
    }
  } catch (err) {
    console.error('[SYNC ERROR] Token sync from sheet failed:', err.message);
  }
}

// Write tokens to sheet
async function syncTokensToSheet(tokens) {
  try {
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId || spreadsheetId === 'mock_google_sheet_id') {
      return;
    }

    let authClient;
    try {
      authClient = getGoogleAuthClient();
    } catch (e) {
      return;
    }

    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: 'sys_tokens!A1:B1',
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [['token_data', JSON.stringify(tokens)]]
        }
      });
      console.log('[SYNC] Saved tokens to Google Sheet sys_tokens tab.');
    } catch (err) {
      if (err.message.includes('NOT_FOUND')) {
        await createSysTokensTab(sheets, spreadsheetId);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: 'sys_tokens!A1:B1',
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [['token_data', JSON.stringify(tokens)]]
          }
        });
        console.log('[SYNC] Created tab and saved tokens to Google Sheet.');
      } else {
        console.error('[SYNC ERROR] Failed to update sys_tokens tab:', err.message);
      }
    }
  } catch (err) {
    console.error('[SYNC ERROR] Token sync to sheet failed:', err.message);
  }
}

async function createSysTokensTab(sheets, spreadsheetId) {
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: 'sys_tokens'
              }
            }
          }
        ]
      }
    });
    console.log('[SYNC] Created sys_tokens tab in spreadsheet.');
  } catch (err) {
    console.error('[SYNC ERROR] Failed to create sys_tokens tab:', err.message);
  }
}

// Connection Status API
app.get('/api/status', requireAuth, async (req, res) => {
  const tokens = readTokens();
  
  const status = {
    google: {
      connected: !!(tokens.google && tokens.google.refresh_token),
      email: tokens.google && tokens.google.refresh_token ? 'Connected Account' : null
    },
    shopify: {
      connected: !!(tokens.shopify && tokens.shopify.access_token),
      shop: tokens.shopify ? tokens.shopify.shop : null
    },
    meta: {
      connected: !!(tokens.meta && tokens.meta.access_token),
      systemTokenConfigured: false
    },
    klaviyo: {
      connected: !!(tokens.klaviyo && tokens.klaviyo.access_token)
    }
  };

  // Check if Meta System User Token is set in Sheet Config (only if Google Sheets is connected)
  const isMockGoogle = !status.google.connected && (!process.env.GOOGLE_CLIENT_ID || process.env.GOOGLE_CLIENT_ID.startsWith('mock_'));
  if (!isMockGoogle && status.google.connected && process.env.GOOGLE_SHEET_ID && process.env.GOOGLE_SHEET_ID !== 'mock_google_sheet_id') {
    try {
      let oauth2Client;
      if (tokens.google && tokens.google.refresh_token) {
        oauth2Client = new google.auth.OAuth2(
          process.env.GOOGLE_CLIENT_ID,
          process.env.GOOGLE_CLIENT_SECRET,
          process.env.GOOGLE_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/google`
        );
        oauth2Client.setCredentials({
          refresh_token: tokens.google.refresh_token,
          access_token: tokens.google.access_token
        });
      } else {
        oauth2Client = new google.auth.OAuth2(
          process.env.GOOGLE_CLIENT_ID,
          process.env.GOOGLE_CLIENT_SECRET,
          process.env.GOOGLE_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/google`
        );
        oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
      }
      const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
      
      const configResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'Config!A1:B50',
      });
      
      const rows = configResponse.data.values;
      if (rows) {
        const configMap = {};
        for (const row of rows) {
          if (row.length >= 2) {
            configMap[row[0].trim()] = row[1].trim();
          }
        }
        if (configMap['Meta System User Token'] || configMap['meta_system_user_token']) {
          status.meta.systemTokenConfigured = true;
        }
      }
    } catch (err) {
      console.error('Error fetching Meta system token state from sheets:', err.message);
    }
  }

  res.json(status);
});

// Helper HTML to close the popup and notify parent window
function renderPopupSuccess(platform) {
  return `
    <!DOCTYPE html>
    <html>
    <head><title>Authentication Success</title></head>
    <body style="font-family: sans-serif; text-align: center; padding-top: 50px; background: #0d0f14; color: #f3f4f6;">
      <h2>Connection Successful!</h2>
      <p style="color: #10b981; font-weight: bold;">Platform: ${platform.toUpperCase()}</p>
      <p style="color: #9ca3af; font-size: 0.9rem;">This window will close automatically...</p>
      <script>
        if (window.opener) {
          window.opener.postMessage({ type: 'oauth-success', platform: '${platform}' }, '*');
        }
        setTimeout(() => { window.close(); }, 1500);
      </script>
    </body>
    </html>
  `;
}

/* ==========================================================================
   SHOPIFY OAUTH & MOCK ROUTES
   ========================================================================== */
app.get('/oauth/shopify', requireAuth, (req, res) => {
  let { shop } = req.query;
  if (!shop) {
    return res.status(400).send('Missing shop parameter');
  }
  
  shop = shop.replace(/^https?:\/\//, '').replace(/\/$/, '');
  if (!shop.includes('.myshopify.com')) {
    shop = `${shop}.myshopify.com`;
  }

  if (process.env.OAUTH_PROXY_URL) {
    const redirectUri = process.env.SHOPIFY_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/shopify`;
    return res.redirect(`${process.env.OAUTH_PROXY_URL}/shopify?shop=${encodeURIComponent(shop)}&redirect_uri=${encodeURIComponent(redirectUri)}`);
  }

  const isMock = !process.env.SHOPIFY_CLIENT_ID || process.env.SHOPIFY_CLIENT_ID.startsWith('mock_');
  
  if (isMock) {
    return res.send(`
      <!DOCTYPE html>
      <html>
      <head><title>Mock Shopify Install</title></head>
      <body style="font-family: sans-serif; background: #0d0f14; color: #f3f4f6; padding: 40px; text-align: center;">
        <div style="max-width: 400px; margin: 0 auto; background: #161c29; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 4px 20px rgba(0,0,0,0.5);">
          <div style="font-size: 2.5rem; margin-bottom: 15px;">🛍️</div>
          <h1 style="font-size: 1.5rem; margin-bottom: 10px;">Simulated Shopify Install</h1>
          <p style="color: #9ca3af; font-size: 0.9rem; margin-bottom: 25px;">Grant Profit Reporter permission to read order lines from store <strong>${shop}</strong>.</p>
          <a href="/oauth/callback/shopify?code=mock_shopify_code&shop=${encodeURIComponent(shop)}" style="display: block; padding: 12px; background: #10b981; color: white; text-decoration: none; border-radius: 8px; font-weight: bold;">Install Mock App</a>
        </div>
      </body>
      </html>
    `);
  }

  const clientId = process.env.SHOPIFY_CLIENT_ID;
  const redirectUri = process.env.SHOPIFY_REDIRECT_URI;
  const scopes = 'read_orders';
  
  const authUrl = `https://${shop}/admin/oauth/authorize?client_id=${clientId}&scope=${scopes}&redirect_uri=${encodeURIComponent(redirectUri)}&state=${shop}`;
  
  res.redirect(authUrl);
});

app.get('/oauth/callback/shopify', async (req, res) => {
  const { code, shop, access_token } = req.query;
  const isMock = !process.env.SHOPIFY_CLIENT_ID || process.env.SHOPIFY_CLIENT_ID.startsWith('mock_');
  
  try {
    let accessToken = access_token;
    if (!accessToken) {
      if (isMock) {
        accessToken = 'mock_shopify_offline_access_token';
      } else {
        const clientId = process.env.SHOPIFY_CLIENT_ID;
        const clientSecret = process.env.SHOPIFY_CLIENT_SECRET;
        
        const tokenUrl = `https://${shop}/admin/oauth/access_token`;
        const response = await axios.post(tokenUrl, {
          client_id: clientId,
          client_secret: clientSecret,
          code
        });
        accessToken = response.data.access_token;
      }
    }

    const allTokens = readTokens();
    allTokens.shopify = {
      shop,
      access_token: accessToken
    };
    writeTokens(allTokens);

    res.send(renderPopupSuccess('shopify'));
  } catch (err) {
    console.error('Shopify OAuth exchange error:', err.response ? err.response.data : err.message);
    res.status(500).send(`Shopify Authentication Failed: ${err.response ? JSON.stringify(err.response.data) : err.message}`);
  }
});

/* ==========================================================================
   META ADS OAUTH & MOCK ROUTES
   ========================================================================== */
app.get('/oauth/meta', requireAuth, (req, res) => {
  if (process.env.OAUTH_PROXY_URL) {
    const redirectUri = process.env.META_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/meta`;
    return res.redirect(`${process.env.OAUTH_PROXY_URL}/meta?redirect_uri=${encodeURIComponent(redirectUri)}`);
  }

  const isMock = !process.env.META_CLIENT_ID || process.env.META_CLIENT_ID.startsWith('mock_');
  
  if (isMock) {
    return res.send(`
      <!DOCTYPE html>
      <html>
      <head><title>Mock Meta Login</title></head>
      <body style="font-family: sans-serif; background: #0d0f14; color: #f3f4f6; padding: 40px; text-align: center;">
        <div style="max-width: 400px; margin: 0 auto; background: #161c29; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 4px 20px rgba(0,0,0,0.5);">
          <div style="font-size: 2.5rem; margin-bottom: 15px;">🎯</div>
          <h1 style="font-size: 1.5rem; margin-bottom: 10px;">Simulated Meta Login</h1>
          <p style="color: #9ca3af; font-size: 0.9rem; margin-bottom: 25px;">Grant Profit Reporter permission to read Facebook campaign marketing costs.</p>
          <a href="/oauth/callback/meta?code=mock_meta_code" style="display: block; padding: 12px; background: #1877f2; color: white; text-decoration: none; border-radius: 8px; font-weight: bold;">Authorize Mock Meta Ads</a>
        </div>
      </body>
      </html>
    `);
  }

  const clientId = process.env.META_CLIENT_ID;
  const redirectUri = process.env.META_REDIRECT_URI;
  const scopes = 'ads_read,read_insights';
  
  const authUrl = `https://www.facebook.com/v19.0/dialog/oauth?client_id=${clientId}&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${scopes}&response_type=code`;
  
  res.redirect(authUrl);
});

app.get('/oauth/callback/meta', async (req, res) => {
  const { code, access_token, expires_in } = req.query;
  const isMock = !process.env.META_CLIENT_ID || process.env.META_CLIENT_ID.startsWith('mock_');
  
  try {
    let longLivedToken = access_token;
    let secondsToExpire = expires_in ? parseInt(expires_in) : 5184000;

    if (!longLivedToken) {
      if (isMock) {
        longLivedToken = 'mock_meta_long_lived_token';
        secondsToExpire = 5184000; // 60 days
      } else {
        const clientId = process.env.META_CLIENT_ID;
        const clientSecret = process.env.META_CLIENT_SECRET;
        const redirectUri = process.env.META_REDIRECT_URI;

        const tokenUrl = `https://graph.facebook.com/v19.0/oauth/access_token`;
        const response = await axios.get(tokenUrl, {
          params: {
            client_id: clientId,
            redirect_uri: redirectUri,
            client_secret: clientSecret,
            code
          }
        });

        const shortLivedToken = response.data.access_token;

        const exchangeResponse = await axios.get(tokenUrl, {
          params: {
            grant_type: 'fb_exchange_token',
            client_id: clientId,
            client_secret: clientSecret,
            fb_exchange_token: shortLivedToken
          }
        });

        longLivedToken = exchangeResponse.data.access_token;
        secondsToExpire = exchangeResponse.data.expires_in;
      }
    }

    const allTokens = readTokens();
    allTokens.meta = {
      access_token: longLivedToken,
      expires_at: DateTime.now().plus({ seconds: secondsToExpire }).toISO()
    };
    writeTokens(allTokens);

    res.send(renderPopupSuccess('meta'));
  } catch (err) {
    console.error('Meta OAuth exchange error:', err.response ? err.response.data : err.message);
    res.status(500).send(`Meta Ads Authentication Failed: ${err.response ? JSON.stringify(err.response.data) : err.message}`);
  }
});

/* ==========================================================================
   KLAVIYO OAUTH & MOCK ROUTES
   ========================================================================== */
app.get('/oauth/klaviyo', requireAuth, (req, res) => {
  if (process.env.OAUTH_PROXY_URL) {
    const redirectUri = process.env.KLAVIYO_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/klaviyo`;
    return res.redirect(`${process.env.OAUTH_PROXY_URL}/klaviyo?redirect_uri=${encodeURIComponent(redirectUri)}`);
  }

  const isMock = !process.env.KLAVIYO_CLIENT_ID || process.env.KLAVIYO_CLIENT_ID.startsWith('mock_');
  
  if (isMock) {
    return res.send(`
      <!DOCTYPE html>
      <html>
      <head><title>Mock Klaviyo Login</title></head>
      <body style="font-family: sans-serif; background: #0d0f14; color: #f3f4f6; padding: 40px; text-align: center;">
        <div style="max-width: 400px; margin: 0 auto; background: #161c29; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 4px 20px rgba(0,0,0,0.5);">
          <div style="font-size: 2.5rem; margin-bottom: 15px;">✉️</div>
          <h1 style="font-size: 1.5rem; margin-bottom: 10px;">Simulated Klaviyo Auth</h1>
          <p style="color: #9ca3af; font-size: 0.9rem; margin-bottom: 25px;">Grant Profit Reporter permission to read Klaviyo metrics and subscribers.</p>
          <a href="/oauth/callback/klaviyo?code=mock_klaviyo_code" style="display: block; padding: 12px; background: #a855f7; color: white; text-decoration: none; border-radius: 8px; font-weight: bold;">Authorize Mock Klaviyo</a>
        </div>
      </body>
      </html>
    `);
  }

  const clientId = process.env.KLAVIYO_CLIENT_ID;
  const redirectUri = process.env.KLAVIYO_REDIRECT_URI;
  const scopes = 'accounts:read metrics:read';
  
  const authUrl = `https://www.klaviyo.com/oauth/authorize?response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scopes)}&state=klaviyo`;
  
  res.redirect(authUrl);
});

app.get('/oauth/callback/klaviyo', async (req, res) => {
  const { code, access_token, refresh_token, expires_in } = req.query;
  const isMock = !process.env.KLAVIYO_CLIENT_ID || process.env.KLAVIYO_CLIENT_ID.startsWith('mock_');
  
  try {
    let accessToken = access_token;
    let refreshToken = refresh_token;
    let secondsToExpire = expires_in ? parseInt(expires_in) : 3600;

    if (!accessToken) {
      if (isMock) {
        accessToken = 'mock_klaviyo_access_token';
        refreshToken = 'mock_klaviyo_refresh_token';
        secondsToExpire = 3600;
      } else {
        const clientId = process.env.KLAVIYO_CLIENT_ID;
        const clientSecret = process.env.KLAVIYO_CLIENT_SECRET;
        const redirectUri = process.env.KLAVIYO_REDIRECT_URI;

        const tokenUrl = `https://a.klaviyo.com/oauth/token`;
        
        const params = new URLSearchParams();
        params.append('grant_type', 'authorization_code');
        params.append('code', code);
        params.append('client_id', clientId);
        params.append('client_secret', clientSecret);
        params.append('redirect_uri', redirectUri);

        const response = await axios.post(tokenUrl, params, {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
          }
        });

        accessToken = response.data.access_token;
        refreshToken = response.data.refresh_token;
        secondsToExpire = response.data.expires_in;
      }
    }

    const allTokens = readTokens();
    allTokens.klaviyo = {
      access_token: accessToken,
      refresh_token: refreshToken,
      expires_at: DateTime.now().plus({ seconds: secondsToExpire }).toISO()
    };
    writeTokens(allTokens);

    res.send(renderPopupSuccess('klaviyo'));
  } catch (err) {
    console.error('Klaviyo OAuth exchange error:', err.response ? err.response.data : err.message);
    res.status(500).send(`Klaviyo Authentication Failed: ${err.response ? JSON.stringify(err.response.data) : err.message}`);
  }
});

/* ==========================================================================
   GOOGLE OAUTH & MOCK ROUTES
   ========================================================================== */
app.get('/oauth/google', requireAuth, (req, res) => {
  if (process.env.OAUTH_PROXY_URL) {
    const redirectUri = process.env.GOOGLE_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/google`;
    return res.redirect(`${process.env.OAUTH_PROXY_URL}/google?redirect_uri=${encodeURIComponent(redirectUri)}`);
  }

  const isMock = !process.env.GOOGLE_CLIENT_ID || process.env.GOOGLE_CLIENT_ID.startsWith('mock_');
  if (isMock) {
    return res.send(`
      <!DOCTYPE html>
      <html>
      <head><title>Mock Google Login</title></head>
      <body style="font-family: sans-serif; background: #0d0f14; color: #f3f4f6; padding: 40px; text-align: center;">
        <div style="max-width: 400px; margin: 0 auto; background: #161c29; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 4px 20px rgba(0,0,0,0.5);">
          <div style="font-size: 2.5rem; margin-bottom: 15px;">🔑</div>
          <h1 style="font-size: 1.5rem; margin-bottom: 10px;">Simulated Google Login</h1>
          <p style="color: #9ca3af; font-size: 0.9rem; margin-bottom: 25px;">Grant Profit Reporter permission to read/write Sheets and send Gmail messages.</p>
          <a href="/oauth/callback/google?code=mock_google_code" style="display: block; padding: 12px; background: #4285f4; color: white; text-decoration: none; border-radius: 8px; font-weight: bold;">Authorize Mock Google Account</a>
        </div>
      </body>
      </html>
    `);
  }

  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/google`
  );

  const scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
  ];

  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: scopes,
    prompt: 'consent'
  });

  res.redirect(url);
});

app.get('/oauth/callback/google', async (req, res) => {
  const { code, access_token, refresh_token, expires_in } = req.query;
  const isMock = !process.env.GOOGLE_CLIENT_ID || process.env.GOOGLE_CLIENT_ID.startsWith('mock_');

  try {
    let accessToken = access_token;
    let refreshToken = refresh_token;
    let expiryDate = expires_in ? Date.now() + parseInt(expires_in) * 1000 : Date.now() + 3600 * 1000;

    if (!accessToken) {
      if (isMock) {
        accessToken = 'mock_google_access_token';
        refreshToken = 'mock_google_refresh_token';
        expiryDate = Date.now() + 3600 * 1000;
      } else {
        const oauth2Client = new google.auth.OAuth2(
          process.env.GOOGLE_CLIENT_ID,
          process.env.GOOGLE_CLIENT_SECRET,
          process.env.GOOGLE_REDIRECT_URI || `http://localhost:${PORT}/oauth/callback/google`
        );
        const { tokens } = await oauth2Client.getToken(code);
        accessToken = tokens.access_token;
        refreshToken = tokens.refresh_token;
        expiryDate = tokens.expiry_date;
      }
    }

    const allTokens = readTokens();
    allTokens.google = {
      access_token: accessToken,
      refresh_token: refreshToken || (allTokens.google && allTokens.google.refresh_token),
      expiry_date: expiryDate
    };
    writeTokens(allTokens);

    res.send(renderPopupSuccess('google'));
  } catch (err) {
    console.error('Google OAuth callback error:', err);
    res.status(500).send(`Google Authentication Failed: ${err.message}`);
  }
});

/* ==========================================================================
   REPORT TRIGGER ENDPOINT
   ========================================================================== */
app.post('/api/trigger-report', requireAuth, async (req, res) => {
  const { type, date } = req.body;
  let targetDateStr;

  if (type === 'custom') {
    if (!date) {
      return res.status(400).json({ success: false, error: 'Date is required for custom type.' });
    }
    targetDateStr = date;
  } else {
    targetDateStr = DateTime.now().setZone('Australia/Melbourne').minus({ days: 1 }).toFormat('yyyy-MM-dd');
  }

  const result = await runReportForDate(targetDateStr);
  if (result.success) {
    res.json(result);
  } else {
    res.status(500).json(result);
  }
});

// Background Scheduler Engine
const cron = require('node-cron');
cron.schedule('55 23 * * *', async () => {
  console.log('[SCHEDULER] Triggering daily profit report at 11:55 PM Melbourne time...');
  const targetDateStr = DateTime.now().setZone('Australia/Melbourne').toFormat('yyyy-MM-dd');
  const result = await runReportForDate(targetDateStr);
  if (result.success) {
    console.log(`[SCHEDULER] Profit report for ${targetDateStr} completed successfully.`);
  } else {
    console.error(`[SCHEDULER] Profit report for ${targetDateStr} failed:`, result.error);
  }
}, {
  scheduled: true,
  timezone: 'Australia/Melbourne'
});

// Start Server
app.listen(PORT, '0.0.0.0', async () => {
  console.log(`Server listening on port ${PORT}`);
  console.log(`Access Portal: http://localhost:${PORT}`);
  
  // Sync tokens from sheet on startup
  await syncTokensFromSheet();
});
