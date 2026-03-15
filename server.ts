import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import cookieParser from "cookie-parser";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cookieParser());

  const getRedirectUri = (req: express.Request) => {
    const baseUrl = process.env.APP_URL || `https://${req.get('host')}`;
    return `${baseUrl}/auth/callback`;
  };

  app.get("/api/auth/status", (req, res) => {
    const hasToken = !!req.cookies.google_access_token;
    const hasConfig = !!process.env.GOOGLE_CLIENT_ID && !!process.env.GOOGLE_CLIENT_SECRET;
    res.json({ isAuthenticated: hasToken, isConfigured: hasConfig });
  });

  app.get("/api/auth/url", (req, res) => {
    const redirectUri = getRedirectUri(req);
    const clientId = process.env.GOOGLE_CLIENT_ID;

    if (!clientId) {
      return res.status(500).json({ error: "GOOGLE_CLIENT_ID is not configured" });
    }

    const params = new URLSearchParams({
      client_id: clientId,
      redirect_uri: redirectUri,
      response_type: 'code',
      scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
      access_type: 'offline',
      prompt: 'consent'
    });

    res.json({ url: `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}` });
  });

  app.get(["/auth/callback", "/auth/callback/"], async (req, res) => {
    const { code } = req.query;
    const redirectUri = getRedirectUri(req);
    const clientId = process.env.GOOGLE_CLIENT_ID;
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

    if (!code || typeof code !== 'string') {
      return res.status(400).send("Missing authorization code");
    }

    try {
      const tokenResponse = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          code,
          client_id: clientId!,
          client_secret: clientSecret!,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code'
        })
      });

      const tokenData = await tokenResponse.json();

      if (!tokenResponse.ok) {
        console.error("Token exchange error:", tokenData);
        return res.status(500).send(`Authentication failed: ${tokenData.error_description || tokenData.error}`);
      }

      res.cookie('google_access_token', tokenData.access_token, {
        secure: true,
        sameSite: 'none',
        httpOnly: true,
        maxAge: tokenData.expires_in * 1000
      });

      res.send(`
        <html>
          <body>
            <script>
              if (window.opener) {
                window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
                window.close();
              } else {
                window.location.href = '/';
              }
            </script>
            <p>Authentication successful. This window should close automatically.</p>
          </body>
        </html>
      `);
    } catch (error) {
      console.error("OAuth callback error:", error);
      res.status(500).send("Internal Server Error during authentication");
    }
  });

  app.get("/api/sheet/metadata", async (req, res) => {
    const { sheetId } = req.query;
    const accessToken = req.cookies.google_access_token;

    if (!accessToken) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    try {
      const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`;
      const metaResponse = await fetch(metaUrl, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      const metaData = await metaResponse.json();
      
      if (!metaResponse.ok) {
        if (metaResponse.status === 401) res.clearCookie('google_access_token');
        return res.status(metaResponse.status).json({ error: metaData.error?.message || "Failed to fetch sheet metadata" });
      }

      const sheets = metaData.sheets?.map((s: any) => ({
        sheetId: s.properties?.sheetId,
        title: s.properties?.title
      })) || [];

      res.json({ sheets });
    } catch (error: any) {
      res.status(500).json({ error: error.message });
    }
  });

  app.get("/api/sheet", async (req, res) => {
    const { sheetId, sheetName, gid } = req.query;
    const accessToken = req.cookies.google_access_token;

    if (!accessToken) {
      return res.status(401).json({ error: "Not authenticated" });
    }

    try {
      let targetSheetName = sheetName as string;

      // If we don't have a specific sheetName, fetch metadata to resolve it via gid or pick the first sheet
      if (!targetSheetName) {
        const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`;
        const metaResponse = await fetch(metaUrl, {
          headers: { Authorization: `Bearer ${accessToken}` }
        });
        const metaData = await metaResponse.json();
        
        if (!metaResponse.ok) {
          if (metaResponse.status === 401) res.clearCookie('google_access_token');
          return res.status(metaResponse.status).json({ error: metaData.error?.message || "Failed to fetch sheet metadata" });
        }

        if (gid) {
          const targetGid = Number(gid);
          const sheet = metaData.sheets?.find((s: any) => s.properties?.sheetId === targetGid);
          if (!sheet) {
            return res.status(404).json({ error: `Sheet with gid=${gid} not found.` });
          }
          targetSheetName = sheet.properties.title;
        } else {
          // Default to the first sheet if no gid is provided
          targetSheetName = metaData.sheets?.[0]?.properties?.title;
        }
      }

      if (!targetSheetName) {
        return res.status(400).json({ error: "Could not determine sheet name." });
      }

      const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(targetSheetName)}`;
      const response = await fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      const data = await response.json();
      if (!response.ok) {
        if (response.status === 401) {
          res.clearCookie('google_access_token');
        }
        return res.status(response.status).json({ error: data.error?.message || "Failed to fetch sheet" });
      }

      res.json({ ...data, resolvedSheetName: targetSheetName });
    } catch (error: any) {
      res.status(500).json({ error: error.message });
    }
  });

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
