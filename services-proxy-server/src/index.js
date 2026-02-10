import path from "path";
import express from "express";
import cors from "cors";
import { fileURLToPath } from "url";
import { createProxyMiddleware } from "http-proxy-middleware";
import { DefaultAzureCredential } from "@azure/identity";
import { getRuntimeConfig, sanitizeRuntimeConfig } from "../../common/runtimeConfig.js";

// ── Managed Identity token service ──────────────────────────────────────────
const COGNITIVE_SERVICES_SCOPE = "https://cognitiveservices.azure.com/.default";
const TOKEN_REFRESH_MARGIN_MS = 9 * 60 * 1000; // refresh 1 min before 10-min expiry
const ICE_TOKEN_REFRESH_MS = 23 * 60 * 60 * 1000; // refresh every 23 hours

let miCredential = null;
let speechTokenCache = null;  // { value, expiresAt }
let iceTokenCache = null;     // { value, expiresAt }

function getMICredential() {
  if (!miCredential) {
    miCredential = new DefaultAzureCredential();
  }
  return miCredential;
}

/**
 * Get an Entra ID (AAD) token for Cognitive Services and compose it
 * in the format required by the Speech SDK: aad#{resourceId}#{aadToken}
 */
async function getManagedIdentitySpeechToken(speechResourceId) {
  const now = Date.now();
  if (speechTokenCache && speechTokenCache.expiresAt > now) {
    return speechTokenCache.value;
  }

  console.log("[ManagedIdentity] Fetching new Entra ID token...");
  const startTime = Date.now();
  const credential = getMICredential();

  let accessToken;
  try {
    accessToken = await credential.getToken(COGNITIVE_SERVICES_SCOPE);
  } catch (error) {
    console.error("[ManagedIdentity] Failed to get Entra ID token:", error.message);
    throw new Error(
      `Failed to authenticate with DefaultAzureCredential. ` +
      `Ensure you are logged in via 'az login' (local) or have Managed Identity configured (Azure). ` +
      `Original error: ${error.message}`
    );
  }

  const elapsed = Date.now() - startTime;
  const compoundToken = `aad#${speechResourceId}#${accessToken.token}`;

  speechTokenCache = { value: compoundToken, expiresAt: now + TOKEN_REFRESH_MARGIN_MS };
  console.log(`[ManagedIdentity] Token acquired in ${elapsed}ms (expires in ~9 min)`);

  return compoundToken;
}

/**
 * Get a raw Entra ID bearer token (without the aad# prefix).
 */
async function getRawBearerToken() {
  const credential = getMICredential();
  const accessToken = await credential.getToken(COGNITIVE_SERVICES_SCOPE);
  return accessToken.token;
}

/**
 * Exchange an Entra ID token for an STS-issued Speech token.
 * The ICE relay endpoint does NOT accept raw Entra ID tokens — it requires
 * a Speech STS token obtained via the issueToken endpoint.
 */
async function getStsSpeechToken(speechEndpoint) {
  const entraToken = await getRawBearerToken();
  const endpoint = speechEndpoint.replace(/^https?:\/\//, "").replace(/\/+$/, "");
  const url = `https://${endpoint}/sts/v1.0/issueToken`;

  console.log("[ManagedIdentity] Exchanging Entra ID token for STS Speech token...");
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${entraToken}`,
      "Content-Length": "0"
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Failed to exchange token via STS: ${response.status} ${response.statusText}. Body: ${body}`);
  }

  const stsToken = await response.text();
  console.log(`[ManagedIdentity] STS token acquired (length: ${stsToken.length})`);
  return stsToken;
}

/**
 * Get ICE relay credentials using an STS speech token (for managed identity mode).
 */
async function getManagedIdentityIceToken(speechRegion, speechEndpoint) {
  const now = Date.now();
  if (iceTokenCache && iceTokenCache.expiresAt > now) {
    return iceTokenCache.value;
  }

  console.log("[ManagedIdentity] Fetching ICE relay token...");
  const stsToken = await getStsSpeechToken(speechEndpoint);

  const url = `https://${speechRegion}.tts.speech.microsoft.com/cognitiveservices/avatar/relay/token/v1`;
  console.log(`[ManagedIdentity] ICE relay URL: ${url}`);

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${stsToken}`
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Failed to fetch ICE relay token: ${response.status} ${response.statusText}. Body: ${body}`);
  }

  const iceToken = await response.json();
  iceTokenCache = { value: iceToken, expiresAt: now + ICE_TOKEN_REFRESH_MS };
  console.log(`[ManagedIdentity] ICE relay token acquired (URLs: ${iceToken.Urls?.join(", ")})`);
  return iceToken;
}
// ─────────────────────────────────────────────────────────────────────────────

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = Number(process.env.SERVICES_PROXY_PORT ?? 4100);
const allowedOriginsRaw = process.env.SERVICES_PROXY_ALLOWED_ORIGINS ?? "";
const allowedOrigins = allowedOriginsRaw
  .split(",")
  .map((entry) => entry.trim())
  .filter((entry) => entry.length > 0);
const agentProxyTarget = process.env.AGENT_PROXY_INTERNAL_URL ?? "http://127.0.0.1:4000";
const agentProxyBasePathEnv = process.env.AGENT_PROXY_PUBLIC_PATH ?? "/agent";
const agentProxyBasePath = agentProxyBasePathEnv.endsWith("/") && agentProxyBasePathEnv !== "/"
  ? agentProxyBasePathEnv.slice(0, -1)
  : agentProxyBasePathEnv;
const enableAgentProxyRelay = (process.env.ENABLE_AGENT_PROXY_RELAY ?? "true").toLowerCase() !== "false";

app.use(
  cors({
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
        callback(null, true);
        return;
      }
      callback(new Error("Origin not allowed"));
    }
  })
);

if (enableAgentProxyRelay) {
  const pathRewrite = (path) => {
    if (path === agentProxyBasePath) {
      return "/";
    }
    if (path.startsWith(`${agentProxyBasePath}/`)) {
      const rewritten = path.slice(agentProxyBasePath.length);
      return rewritten !== "" ? rewritten : "/";
    }
    return path;
  };

  app.use(
    agentProxyBasePath,
    createProxyMiddleware({
      target: agentProxyTarget,
      changeOrigin: false,
      ws: true,
      pathRewrite,
      proxyTimeout: Number(process.env.AGENT_PROXY_TIMEOUT ?? 30000)
    })
  );
}

// Increase body size limit to handle large base64-encoded images (up to 50MB)
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

function buildPublicConfig(config) {
  const sanitized = sanitizeRuntimeConfig(config);
  const publicBaseRaw = (process.env.SERVICES_PROXY_PUBLIC_BASE_URL ?? "").trim();
  if (publicBaseRaw) {
    const normalizedBase = publicBaseRaw.endsWith("/") ? publicBaseRaw.slice(0, -1) : publicBaseRaw;
    sanitized.servicesProxyBaseUrl = normalizedBase;
    if (sanitized.agent) {
      sanitized.agent.apiUrl = `${normalizedBase}${agentProxyBasePath}`;
    }
  }
  return sanitized;
}

app.get("/config", (_req, res) => {
  const config = getRuntimeConfig();
  res.json(buildPublicConfig(config));
});

function resolveSpeechTokenEndpoint(config) {
  const privateEndpoint = (config.speech?.privateEndpoint ?? "").trim();
  const enablePrivateEndpoint = Boolean(config.speech?.enablePrivateEndpoint);

  if (enablePrivateEndpoint && privateEndpoint !== "") {
    const normalized = privateEndpoint.replace(/^https?:\/\//i, "");
    return {
      relay: `https://${normalized}/tts/cognitiveservices/avatar/relay/token/v1`,
      speech: `https://${normalized}/sts/v1.0/issuetoken`
    };
  }

  const region = (config.speech?.region ?? "").trim();
  if (!region) {
    throw new Error("Speech region is not configured");
  }
  return {
    relay: `https://${region}.tts.speech.microsoft.com/cognitiveservices/avatar/relay/token/v1`,
    speech: `https://${region}.api.cognitive.microsoft.com/sts/v1.0/issuetoken`
  };
}

app.post("/speech/token", async (_req, res) => {
  try {
    const config = getRuntimeConfig();
    const useManagedIdentity = Boolean(config.speech?.useManagedIdentity);
    const apiKey = (config.speech?.apiKey ?? "").trim();

    if (useManagedIdentity) {
      // ── Managed Identity / Entra ID flow ──
      const speechResourceId = (config.speech?.speechResourceId ?? "").trim();
      const speechEndpoint = (config.speech?.speechEndpoint ?? config.speech?.privateEndpoint ?? "").trim();
      const speechRegion = (config.speech?.region ?? "").trim();

      if (!speechResourceId) {
        res.status(500).json({ error: "speech.speechResourceId is required when useManagedIdentity is enabled" });
        return;
      }
      if (!speechEndpoint) {
        res.status(500).json({ error: "speech.speechEndpoint (custom domain) is required when useManagedIdentity is enabled" });
        return;
      }
      if (!speechRegion) {
        res.status(500).json({ error: "speech.region is required" });
        return;
      }

      const compoundToken = await getManagedIdentitySpeechToken(speechResourceId);
      const iceToken = await getManagedIdentityIceToken(speechRegion, speechEndpoint);

      res.json({
        speechToken: compoundToken,
        speechTokenExpiresInSeconds: 540,
        relay: iceToken,
        region: speechRegion,
        useManagedIdentity: true,
        speechEndpoint: speechEndpoint
      });
      return;
    }

    // ── API key flow (existing behavior) ──
    if (!apiKey) {
      res.status(500).json({ error: "Speech resource key is not configured and useManagedIdentity is not enabled" });
      return;
    }

    const endpoints = resolveSpeechTokenEndpoint(config);

    const relayResponse = await fetch(endpoints.relay, {
      method: "GET",
      headers: {
        "Ocp-Apim-Subscription-Key": apiKey
      }
    });

    if (!relayResponse.ok) {
      const text = await relayResponse.text();
      res.status(relayResponse.status).json({ error: "Failed to obtain speech relay token", details: text });
      return;
    }

    const relayPayload = await relayResponse.json();

    const speechTokenResponse = await fetch(endpoints.speech, {
      method: "POST",
      headers: {
        "Ocp-Apim-Subscription-Key": apiKey,
        "Content-Length": "0"
      }
    });

    if (!speechTokenResponse.ok) {
      const text = await speechTokenResponse.text();
      res.status(speechTokenResponse.status).json({ error: "Failed to obtain speech token", details: text });
      return;
    }

    const speechToken = await speechTokenResponse.text();

    res.json({
      speechToken,
      speechTokenExpiresInSeconds: 540,
      relay: relayPayload,
      region: (config.speech?.region ?? "").trim()
    });
  } catch (err) {
    console.error("/speech/token error", err);
    res.status(500).json({ error: err.message ?? String(err) });
  }
});

// Generic plugin API endpoint handler - handle both JSON and multipart/form-data
app.post("/plugin/:pluginId/api/:endpoint", async (req, res) => {
  try {
    const { pluginId, endpoint } = req.params;
    
    const config = getRuntimeConfig();
    
    // Get plugin config
    const pluginConfig = config.agent?.plugins?.[pluginId];
    if (!pluginConfig) {
      res.status(404).json({ error: `Plugin '${pluginId}' not found` });
      return;
    }

    // Get API handler configuration from plugin manifest
    const apiHandlers = pluginConfig.apiHandlers || {};
    const handlerConfig = apiHandlers[endpoint];
    
    if (!handlerConfig) {
      res.status(404).json({ error: `API endpoint '${endpoint}' not configured for plugin '${pluginId}'` });
      return;
    }

    // Check if handler is enabled
    if (handlerConfig.enabled === false) {
      res.status(400).json({ error: `API endpoint '${endpoint}' is disabled` });
      return;
    }

    // Helper function to replace template variables like {{imageGeneration.apiKey}}
    const replaceTemplateVars = (template, context) => {
      if (typeof template !== 'string') return template;
      
      return template.replace(/\{\{([^}]+)\}\}/g, (match, path) => {
        const keys = path.trim().split('.');
        let value = context;
        for (const key of keys) {
          if (value && typeof value === 'object' && key in value) {
            value = value[key];
          } else {
            console.warn(`Template variable not found: ${path}`);
            return match;
          }
        }
        return value;
      });
    };

    // Handle special targets
    if (handlerConfig.target === 'azure-blob-storage') {
      // Azure Blob Storage upload handler with Managed Identity support
      const { BlobServiceClient } = await import('@azure/storage-blob');
      const { DefaultAzureCredential } = await import('@azure/identity');
      
      // Resolve template variables in config
      const authMethod = replaceTemplateVars(handlerConfig.config?.authMethod || 'managedIdentity', pluginConfig);
      const accountName = replaceTemplateVars(handlerConfig.config?.accountName, pluginConfig);
      const containerName = replaceTemplateVars(handlerConfig.config?.containerName, pluginConfig);
      const connectionString = replaceTemplateVars(handlerConfig.config?.connectionString, pluginConfig);
      
      if (!accountName || accountName === 'YOUR_STORAGE_ACCOUNT_NAME') {
        res.status(500).json({ error: 'Blob storage account name not configured' });
        return;
      }

      try {
        // req.body is already parsed by express.json() middleware
        const { filename, contentType, data } = req.body;
        
        if (!filename || !data) {
          res.status(400).json({ error: 'Missing filename or data' });
          return;
        }

        // Convert base64 to buffer
        const buffer = Buffer.from(data.split(',')[1] || data, 'base64');
        
        // Create BlobServiceClient based on auth method
        let blobServiceClient;
        if (authMethod === 'managedIdentity') {
          const credential = new DefaultAzureCredential();
          const blobStorageUrl = `https://${accountName}.blob.core.windows.net`;
          blobServiceClient = new BlobServiceClient(blobStorageUrl, credential);
        } else if (authMethod === 'connectionString' && connectionString) {
          blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
        } else {
          res.status(500).json({ error: 'Invalid authentication method or missing credentials' });
          return;
        }
        
        const containerClient = blobServiceClient.getContainerClient(containerName);
        
        // Ensure container exists (private access - no public blob access allowed)
        await containerClient.createIfNotExists();
        
        const blockBlobClient = containerClient.getBlockBlobClient(filename);
        await blockBlobClient.upload(buffer, buffer.length, {
          blobHTTPHeaders: { blobContentType: contentType || 'image/png' }
        });
        
        // Return proxy URL instead of direct blob URL (no public access or SAS allowed)
        // The image will be served through our proxy endpoint using managed identity
        const proxyUrl = `/plugin/${pluginId}/blob/${containerName}/${filename}`;
        res.json({ url: proxyUrl, filename });
        return;
      } catch (error) {
        console.error('Blob upload error:', error);
        console.error('Error details:', {
          message: error.message,
          code: error.code,
          statusCode: error.statusCode,
          stack: error.stack
        });
        res.status(500).json({ 
          error: error.message,
          code: error.code,
          details: error.statusCode || error.code || 'Unknown error'
        });
        return;
      }
    }

    if (handlerConfig.target === 'qr-code-generator') {
      // QR Code generation handler
      const QRCode = (await import('qrcode')).default;
      
      try {
        // req.body is already parsed by express.json() middleware
        const { url } = req.body;
        
        if (!url) {
          res.status(400).json({ error: 'Missing URL' });
          return;
        }

        // Generate QR code as data URL
        const qrCodeDataUrl = await QRCode.toDataURL(url, {
          width: 400,
          margin: 2,
          color: {
            dark: '#000000',
            light: '#FFFFFF'
          }
        });
        
        res.json({ qrCode: qrCodeDataUrl, url });
        return;
      } catch (error) {
        console.error('QR code generation error:', error);
        res.status(500).json({ error: error.message });
        return;
      }
    }

    // Build target URL with template variable replacement
    let targetUrl = replaceTemplateVars(handlerConfig.target, pluginConfig);
    
    // Add query parameters if specified
    if (handlerConfig.query && typeof handlerConfig.query === 'object') {
      const queryParams = new URLSearchParams();
      for (const [key, value] of Object.entries(handlerConfig.query)) {
        const resolvedValue = replaceTemplateVars(value, pluginConfig);
        if (resolvedValue) {
          queryParams.append(key, resolvedValue);
        }
      }
      const queryString = queryParams.toString();
      if (queryString) {
        targetUrl += (targetUrl.includes('?') ? '&' : '?') + queryString;
      }
    }

    // Handle request body based on content type
    const method = handlerConfig.method || 'POST';
    const contentType = req.headers['content-type'] || '';
    let requestBody;

    // Build headers with template variable replacement
    const headers = {};
    if (handlerConfig.headers && typeof handlerConfig.headers === 'object') {
      for (const [key, value] of Object.entries(handlerConfig.headers)) {
        // Skip content-type for multipart, as FormData will set it with boundary
        if (key.toLowerCase() === 'content-type' && contentType.includes('multipart/form-data')) {
          continue;
        }
        headers[key] = replaceTemplateVars(value, pluginConfig);
      }
    }

    if (method !== 'GET') {
      if (contentType.includes('multipart/form-data')) {
        console.log(`[Plugin API] Handling multipart/form-data for ${pluginId}/${endpoint}`);
        
        // Parse multipart data
        const formidable = (await import('formidable')).default;
        const form = formidable({ multiples: true, maxFileSize: 50 * 1024 * 1024 });
        
        const [fields, files] = await new Promise((resolve, reject) => {
          form.parse(req, (err, fields, files) => {
            if (err) {
              console.error('[Plugin API] Formidable parse error:', err);
              reject(err);
            } else {
              console.log('[Plugin API] Parsed fields:', Object.keys(fields));
              console.log('[Plugin API] Parsed files:', Object.keys(files));
              resolve([fields, files]);
            }
          });
        });

        // Use global FormData and Blob (available in Node 18+)
        const fs = await import('fs');
        const formData = new FormData();
        
        // Add all fields
        for (const [key, value] of Object.entries(fields)) {
          const val = Array.isArray(value) ? value[0] : value;
          console.log(`[Plugin API] Adding field: ${key}`);
          formData.append(key, val);
        }
        
        // Add all files as Blobs
        for (const [key, fileArray] of Object.entries(files)) {
          const fileList = Array.isArray(fileArray) ? fileArray : [fileArray];
          for (const file of fileList) {
            console.log(`[Plugin API] Adding file: ${key} (${file.originalFilename}, ${file.size} bytes, ${file.mimetype})`);
            const fileBuffer = await fs.promises.readFile(file.filepath);
            const blob = new Blob([fileBuffer], { type: file.mimetype || 'image/png' });
            formData.append(key, blob, file.originalFilename || 'file.png');
          }
        }
        
        requestBody = formData;
        // Global FormData handles its own headers automatically
        
      } else {
        // For JSON or other content types, use express.json() parsed body
        await new Promise((resolve) => {
          express.json({ limit: '50mb' })(req, res, resolve);
        });
        requestBody = JSON.stringify(req.body);
        headers['content-type'] = 'application/json';
      }
    }

    // Forward the request
    console.log(`[Plugin API] Forwarding ${method} to ${targetUrl}`);
    console.log(`[Plugin API] Headers:`, Object.keys(headers));
    
    const response = await fetch(targetUrl, {
      method,
      headers,
      body: requestBody
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`Plugin API error [${pluginId}/${endpoint}]:`, response.status, errorText);
      res.status(response.status).json({ 
        error: `API request failed`, 
        details: errorText 
      });
      return;
    }

    const responseData = await response.json();
    res.json(responseData);
  } catch (err) {
    console.error('/plugin/:pluginId/api/:endpoint error', err);
    console.error('Error stack:', err.stack);
    res.status(500).json({ error: err.message ?? String(err) });
  }
});

// Blob proxy endpoint - serves private blobs using managed identity
app.get("/plugin/:pluginId/blob/:containerName/:filename", async (req, res) => {
  try {
    const { pluginId, containerName, filename } = req.params;
    
    const config = getRuntimeConfig();
    const pluginConfig = config.agent?.plugins?.[pluginId];
    
    if (!pluginConfig) {
      res.status(404).json({ error: `Plugin '${pluginId}' not found` });
      return;
    }

    // Get blob storage config from plugin
    const accountName = pluginConfig.blobStorage?.accountName;
    
    if (!accountName || accountName === 'YOUR_STORAGE_ACCOUNT_NAME') {
      res.status(500).json({ error: 'Blob storage not configured' });
      return;
    }

    // Use managed identity to fetch blob
    const { BlobServiceClient } = await import('@azure/storage-blob');
    const { DefaultAzureCredential } = await import('@azure/identity');
    
    const credential = new DefaultAzureCredential();
    const blobStorageUrl = `https://${accountName}.blob.core.windows.net`;
    const blobServiceClient = new BlobServiceClient(blobStorageUrl, credential);
    
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blockBlobClient = containerClient.getBlockBlobClient(filename);
    
    // Download blob
    const downloadResponse = await blockBlobClient.download();
    
    // Set content type
    const contentType = downloadResponse.contentType || 'image/png';
    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'public, max-age=31536000'); // Cache for 1 year
    
    // Pipe blob stream to response
    downloadResponse.readableStreamBody.pipe(res);
    
  } catch (error) {
    console.error('Blob proxy error:', error);
    res.status(500).json({ error: error.message });
  }
});

const staticRoot = path.resolve(__dirname, "..", "..");
const staticDir = process.env.STATIC_ASSETS_DIR
  ? path.resolve(process.env.STATIC_ASSETS_DIR)
  : path.resolve(staticRoot, "public");
const topLevelStaticDir = staticRoot;

if (process.env.SERVE_STATIC !== "false") {
  app.use(express.static(staticDir));
  app.use(express.static(topLevelStaticDir));
  app.get("/", (req, res, next) => {
    res.sendFile(path.join(staticDir, "chat_agent.html"), (err) => {
      if (err) {
        next(err);
      }
    });
  });
}

app.get("/healthz", (_req, res) => {
  res.json({ status: "ok" });
});

// Health check for Managed Identity authentication
app.get("/speech/health", async (_req, res) => {
  try {
    const config = getRuntimeConfig();
    const useManagedIdentity = Boolean(config.speech?.useManagedIdentity);
    if (!useManagedIdentity) {
      res.json({ status: "ok", authMode: "apiKey", message: "Using API key authentication" });
      return;
    }

    const speechResourceId = (config.speech?.speechResourceId ?? "").trim();
    const speechEndpoint = (config.speech?.speechEndpoint ?? "").trim();
    const speechRegion = (config.speech?.region ?? "").trim();

    const credential = getMICredential();
    const accessToken = await credential.getToken(COGNITIVE_SERVICES_SCOPE);

    res.json({
      status: "healthy",
      authMode: "managedIdentity",
      region: speechRegion,
      endpoint: speechEndpoint,
      resourceId: speechResourceId ? `${speechResourceId.substring(0, 20)}...` : "",
      tokenLength: accessToken.token.length,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(503).json({
      status: "unhealthy",
      authMode: "managedIdentity",
      error: error.message,
      timestamp: new Date().toISOString()
    });
  }
});

// Clear cached tokens (for debugging managed identity)
app.post("/speech/clear-cache", (_req, res) => {
  speechTokenCache = null;
  iceTokenCache = null;
  console.log("[ManagedIdentity] Token cache cleared");
  res.json({ message: "Token cache cleared" });
});

app.get("/favicon.ico", (req, res, next) => {
  const faviconPath = path.join(topLevelStaticDir, "favicon.ico");
  res.sendFile(faviconPath, (err) => {
    if (err) {
      next(err);
    }
  });
});

app.listen(port, () => {
  console.log(`Services proxy listening on http://localhost:${port}`);
});
