import express from "express";
import cors from "cors";
import { DefaultAzureCredential } from "@azure/identity";
import { AIProjectClient } from "@azure/ai-projects";
import { randomUUID } from "crypto";
import { setTimeout as delay } from "timers/promises";
import { getRuntimeConfig, getPublicRuntimeConfig } from "./runtimeConfig.js";

const app = express();
const port = process.env.PORT || 4000;
const DIRECT_LINE_DEFAULT_BASE_URL = "https://directline.botframework.com";
const DIRECT_LINE_POLL_INTERVAL_MS = Number(process.env.DIRECT_LINE_POLL_INTERVAL_MS ?? 1000);
const DIRECT_LINE_STREAM_TIMEOUT_MS = Number(process.env.DIRECT_LINE_STREAM_TIMEOUT_MS ?? 60000);

const directLineConversations = new Map();

function normalizeProvider(value) {
  if (typeof value !== "string") {
    return "";
  }
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return "";
  }
  if (["copilotstudio", "copilot-studio", "copilot_studio", "copilot", "directline", "direct-line", "direct_line"].includes(trimmed)) {
    return "copilot_studio";
  }
  if (["azure-ai-foundry", "azureaifoundry", "aifoundry", "azure"].includes(trimmed)) {
    return "azure_ai_foundry";
  }
  return trimmed;
}

function getAgentRuntimeContext(requestedPluginId) {
  const config = getRuntimeConfig();
  const pluginMap = config.agent?.plugins || {};
  const fallbackPluginId = requestedPluginId
    || config.agent?.activePluginId
    || config.agent?.selectedPluginId
    || config.agent?.defaultPluginId
    || null;
  const pluginRecord = fallbackPluginId ? pluginMap[fallbackPluginId] : null;
  const connection = pluginRecord?.connection || {};
  const provider = normalizeProvider(connection.provider || config.agent?.provider) || "azure_ai_foundry";
  return {
    config,
    pluginId: fallbackPluginId,
    pluginRecord,
    connection,
    provider
  };
}

function resolveDirectLineSecret(connection, agentConfig) {
  console.log('[DirectLine] Resolving Direct Line secret...');
  
  // Priority 1: Direct secret in connection (from plugin manifest)
  if (connection?.directLineSecret) {
    const trimmed = typeof connection.directLineSecret === 'string' ? connection.directLineSecret.trim() : '';
    if (trimmed) {
      console.log('[DirectLine] Found secret in connection.directLineSecret');
      return trimmed;
    }
  }
  
  // Priority 2: Secret via environment variable name (from plugin manifest)
  if (connection?.directLineSecretEnv) {
    const envVarName = connection.directLineSecretEnv;
    console.log(`[DirectLine] Looking for env var: ${envVarName}`);
    if (process.env[envVarName]) {
      const trimmed = process.env[envVarName].trim();
      if (trimmed) {
        console.log(`[DirectLine] Found secret in env var: ${envVarName}`);
        return trimmed;
      }
    }
    console.log(`[DirectLine] Env var ${envVarName} not set or empty`);
  }
  
  // Priority 3: Legacy fallback to global agent config
  const agentDirectLine = agentConfig?.directLine || {};
  if (agentDirectLine.secret) {
    const trimmed = typeof agentDirectLine.secret === 'string' ? agentDirectLine.secret.trim() : '';
    if (trimmed) {
      console.log('[DirectLine] Found secret in agentConfig.directLine.secret');
      return trimmed;
    }
  }
  
  // Priority 4: Legacy env var from global config
  if (agentDirectLine.secretEnv && process.env[agentDirectLine.secretEnv]) {
    const trimmed = process.env[agentDirectLine.secretEnv].trim();
    if (trimmed) {
      console.log(`[DirectLine] Found secret in env var: ${agentDirectLine.secretEnv}`);
      return trimmed;
    }
  }
  
  // Priority 5: Global fallback env var
  if (process.env.DIRECT_LINE_SECRET) {
    const trimmed = process.env.DIRECT_LINE_SECRET.trim();
    if (trimmed) {
      console.log('[DirectLine] Found secret in DIRECT_LINE_SECRET env var');
      return trimmed;
    }
  }
  
  console.warn('[DirectLine] No Direct Line secret found in any source');
  return null;
}

function resolveDirectLineContext(providerHint, requestedPluginId) {
  const context = getAgentRuntimeContext(requestedPluginId);
  const provider = normalizeProvider(providerHint) || context.provider;
  const agentDirectLine = context.config.agent?.directLine || {};
  const endpoint = context.connection.directLineEndpoint
    || context.connection.endpoint
    || agentDirectLine.endpoint
    || DIRECT_LINE_DEFAULT_BASE_URL;

  return {
    ...context,
    provider,
    directLine: {
      endpoint,
      botId: context.connection.directLineBotId
        || agentDirectLine.botId
        || context.connection.agentId
        || context.config.agent?.agentId
        || null,
      userId: context.connection.directLineUserId
        || agentDirectLine.userId
        || null,
      scope: context.connection.directLineScope
        || agentDirectLine.scope
        || null,
      region: context.connection.directLineRegion
        || agentDirectLine.region
        || null,
      secret: resolveDirectLineSecret(context.connection, context.config.agent)
    }
  };
}

function updateConversationExpiry(conversation, expiresInSeconds) {
  const safeSeconds = typeof expiresInSeconds === "number" && Number.isFinite(expiresInSeconds)
    ? Math.max(expiresInSeconds - 30, 30)
    : 1800;
  conversation.expiresAt = Date.now() + safeSeconds * 1000;
}

async function createDirectLineConversation(resolvedContext) {
  const { directLine, pluginId } = resolvedContext;
  
  console.log(`[DirectLine] Creating conversation for plugin: ${pluginId || '(none)'}`);
  console.log(`[DirectLine] Endpoint: ${directLine.endpoint || '(not set)'}`);
  console.log(`[DirectLine] BotId: ${directLine.botId || '(not set)'}`);
  console.log(`[DirectLine] Secret available: ${directLine.secret ? 'yes' : 'no'}`);
  
  if (!directLine.secret) {
    const errorMsg = `Direct Line secret is not configured for plugin '${pluginId || 'unknown'}'. Please ensure the manifest includes directLineSecretEnv pointing to a valid environment variable.`;
    console.error(`[DirectLine] ${errorMsg}`);
    throw new Error(errorMsg);
  }

  const baseUrl = directLine.endpoint.replace(/\/$/, "");
  const userId = directLine.userId || `user-${randomUUID()}`;

  const tokenPayload = {
    user: { id: userId }
  };

  if (directLine.botId) {
    tokenPayload.bot = { id: directLine.botId };
  }
  if (directLine.scope) {
    tokenPayload.scope = directLine.scope;
  }

  console.log(`[DirectLine] Requesting token from: ${baseUrl}/v3/directline/tokens/generate`);
  
  const tokenResponse = await fetch(`${baseUrl}/v3/directline/tokens/generate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${directLine.secret}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(tokenPayload)
  });

  if (!tokenResponse.ok) {
    const detail = await tokenResponse.text();
    throw new Error(`Direct Line token request failed (${tokenResponse.status}): ${detail}`);
  }

  const tokenResult = await tokenResponse.json();
  const conversationIdFromToken = tokenResult?.conversationId;
  const token = tokenResult?.token;
  let streamUrl = tokenResult?.streamUrl || null;
  let expirationSeconds = tokenResult?.expires_in ?? tokenResult?.expiresIn ?? 1800;

  if (!token) {
    throw new Error("Direct Line token response missing token value.");
  }

  let conversationId = conversationIdFromToken || null;
  try {
    const startResponse = await fetch(`${baseUrl}/v3/directline/conversations`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({})
    });
    if (startResponse.ok) {
      const startPayload = await startResponse.json();
      conversationId = startPayload?.conversationId || conversationId;
      streamUrl = startPayload?.streamUrl || streamUrl;
      if (startPayload?.expires_in) {
        expirationSeconds = startPayload.expires_in;
      }
    } else if (!conversationId) {
      const detail = await startResponse.text();
      throw new Error(`Failed to start Direct Line conversation (${startResponse.status}): ${detail}`);
    }
  } catch (err) {
    if (!conversationId) {
      throw err;
    }
    console.warn("Direct Line conversation start encountered an error but continuing with generated conversation ID", err);
  }

  if (!conversationId) {
    throw new Error("Direct Line conversation ID could not be determined.");
  }

  const conversation = {
    id: conversationId,
    token,
    endpoint: baseUrl,
    pluginId,
    secret: directLine.secret,
    botId: directLine.botId || null,
    userId,
    streamUrl,
    watermark: null,
    deliveredActivityIds: new Set()
  };

  updateConversationExpiry(conversation, expirationSeconds);
  directLineConversations.set(conversationId, conversation);
  return conversation;
}

async function refreshDirectLineToken(conversation) {
  try {
    const response = await fetch(`${conversation.endpoint}/v3/directline/tokens/refresh`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${conversation.token}`,
        "Content-Type": "application/json"
      }
    });
    if (response.ok) {
      const payload = await response.json();
      if (payload?.token) {
        conversation.token = payload.token;
      }
      updateConversationExpiry(conversation, payload?.expires_in ?? payload?.expiresIn ?? 1800);
      return;
    }
    const detail = await response.text();
    throw new Error(`Direct Line token refresh failed (${response.status}): ${detail}`);
  } catch (err) {
    console.error("Direct Line token refresh error", err);
    throw err;
  }
}

async function ensureDirectLineToken(conversation) {
  if (!conversation) {
    throw new Error("Conversation not found");
  }
  if (!conversation.expiresAt || conversation.expiresAt - Date.now() > 60000) {
    return conversation.token;
  }
  await refreshDirectLineToken(conversation);
  return conversation.token;
}

async function fetchDirectLineActivities(conversation) {
  await ensureDirectLineToken(conversation);
  const watermarkQuery = conversation.watermark ? `?watermark=${encodeURIComponent(conversation.watermark)}` : "";
  const response = await fetch(`${conversation.endpoint}/v3/directline/conversations/${conversation.id}/activities${watermarkQuery}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${conversation.token}`
    }
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Direct Line activities request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  if (payload?.watermark !== undefined) {
    conversation.watermark = payload.watermark;
  }
  return Array.isArray(payload?.activities) ? payload.activities : [];
}

function mapDirectLineAttachments(activity) {
  if (!activity || !Array.isArray(activity.attachments)) {
    return [];
  }
  return activity.attachments
    .map((attachment) => {
      if (!attachment) {
        return null;
      }
      const contentType = typeof attachment.contentType === "string" ? attachment.contentType : "";
      const url = typeof attachment.contentUrl === "string" ? attachment.contentUrl : "";
      const name = typeof attachment.name === "string" ? attachment.name : undefined;
      if (!url) {
        return null;
      }
      const isImage = contentType.startsWith("image/");
      return {
        type: isImage ? "image" : "attachment",
        url,
        title: name || (isImage ? "Assistant shared an image" : "Assistant shared a file"),
        contentType,
        name
      };
    })
    .filter(Boolean);
}

function transformActivityToDelta(activity) {
  // Check if activity has structured content array (e.g., from Copilot Studio)
  // This could be in channelData, suggestedActions, or the text field itself might be JSON
  let structuredContent = null;
  
  // Try to parse text as JSON array if it looks like structured content
  const rawText = typeof activity?.text === "string" ? activity.text : "";
  if (rawText.trim().startsWith("[") && rawText.trim().endsWith("]")) {
    try {
      const parsed = JSON.parse(rawText);
      if (Array.isArray(parsed) && parsed.some(item => item?.type)) {
        structuredContent = parsed;
        console.log(`[DirectLine] Detected structured content in activity.text with ${structuredContent.length} items`);
      }
    } catch {
      // Not JSON, treat as plain text
    }
  }
  
  // Check channelData for structured content
  if (!structuredContent && activity?.channelData?.content && Array.isArray(activity.channelData.content)) {
    structuredContent = activity.channelData.content;
    console.log(`[DirectLine] Detected structured content in channelData with ${structuredContent.length} items`);
  }
  
  // If we have structured content, filter for speakable text only (type: "text")
  let speakableText = "";
  if (structuredContent) {
    const textItems = structuredContent.filter(item => item?.type === "text" && typeof item?.text === "string");
    speakableText = textItems.map(item => item.text).join(" ").trim();
    console.log(`[DirectLine] Filtered ${textItems.length} text items for TTS from ${structuredContent.length} total items`);
  } else {
    // Fall back to plain text
    speakableText = rawText;
  }
  
  if (!speakableText) {
    return null;
  }
  
  return {
    id: activity.id,
    message_id: activity.id,
    delta: {
      content: [
        {
          type: "output_text",
          text: { value: speakableText }
        }
      ]
    }
  };
}

function transformActivityToCompletion(activity) {
  const text = typeof activity?.text === "string" ? activity.text : "";
  const attachments = mapDirectLineAttachments(activity);
  
  // Check if activity has structured content array (same logic as delta)
  let structuredContent = null;
  if (text.trim().startsWith("[") && text.trim().endsWith("]")) {
    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed) && parsed.some(item => item?.type)) {
        structuredContent = parsed;
        console.log(`[DirectLine] Completion: structured content with ${structuredContent.length} items`);
      }
    } catch {
      // Not JSON, treat as plain text
    }
  }
  
  if (!structuredContent && activity?.channelData?.content && Array.isArray(activity.channelData.content)) {
    structuredContent = activity.channelData.content;
    console.log(`[DirectLine] Completion: structured content from channelData with ${structuredContent.length} items`);
  }
  
  const content = [];
  
  // If we have structured content, preserve it fully in completion
  if (structuredContent) {
    structuredContent.forEach(item => {
      if (!item) return;
      
      if (item.type === "text" && typeof item.text === "string") {
        content.push({
          type: "output_text",
          text: { value: item.text }
        });
      } else if (item.type === "image" && typeof item.text === "string") {
        // Image URL in text field
        content.push({
          type: "image",
          url: item.text,
          title: item.title || "Assistant shared an image"
        });
      } else if (item.type && item.text) {
        // Other structured content types
        content.push({
          type: item.type,
          text: item.text,
          title: item.title
        });
      }
    });
    console.log(`[DirectLine] Completion: transformed to ${content.length} content items`);
  } else {
    // Fall back to plain text + attachments
    if (text) {
      content.push({
        type: "output_text",
        text: { value: text }
      });
    }
  }

  // Add traditional attachments from Direct Line
  attachments.forEach((attachment) => {
    const item = {
      type: attachment.type,
      url: attachment.url,
      title: attachment.title,
      contentType: attachment.contentType
    };
    if (attachment.name) {
      item.name = attachment.name;
    }
    content.push(item);
  });

  return {
    id: activity.id,
    role: "assistant",
    message: {
      id: activity.id,
      role: "assistant",
      content
    },
    data: {
      activity,
      content
    },
    content,
    attachments
  };
}

function normalizeDirectLineUserContent(content) {
  let text = "";
  const attachments = [];

  if (Array.isArray(content)) {
    content.forEach((item) => {
      if (!item || typeof item !== "object") {
        return;
      }
      if (item.type === "text" && typeof item.text === "string" && item.text.trim()) {
        text = item.text.trim();
      }
      if (item.type === "image_url" && item.image_url && typeof item.image_url.url === "string") {
        const attachmentUrl = item.image_url.url.trim();
        if (!attachmentUrl) {
          return;
        }

        let contentType = typeof item.image_url.content_type === "string" && item.image_url.content_type.trim()
          ? item.image_url.content_type.trim()
          : typeof item.image_url.mimeType === "string" && item.image_url.mimeType.trim()
            ? item.image_url.mimeType.trim()
            : "image/png";

        if (!contentType || contentType === "image/png") {
          const urlWithoutQuery = attachmentUrl.split("?")[0].toLowerCase();
          if (urlWithoutQuery.endsWith(".jpg") || urlWithoutQuery.endsWith(".jpeg")) {
            contentType = "image/jpeg";
          } else if (urlWithoutQuery.endsWith(".gif")) {
            contentType = "image/gif";
          } else if (urlWithoutQuery.endsWith(".webp")) {
            contentType = "image/webp";
          }
        }

        attachments.push({
          contentType,
          contentUrl: attachmentUrl,
          name: item.image_url.name || "image"
        });
      }
    });
  } else if (typeof content === "string") {
    text = content;
  } else if (content && typeof content === "object" && typeof content.text === "string") {
    text = content.text;
  }

  return { text, attachments };
}

async function streamDirectLineConversation(conversation, sendEvent, isClosed) {
  const delivered = conversation.deliveredActivityIds || new Set();
  conversation.deliveredActivityIds = delivered;
  const startTime = Date.now();

  while (!isClosed()) {
    const activities = await fetchDirectLineActivities(conversation);
    const newBotActivities = activities.filter((activity) => {
      if (!activity || !activity.id) {
        return false;
      }
      if (delivered.has(activity.id)) {
        return false;
      }
      if ((activity.type || "").toLowerCase() !== "message") {
        return false;
      }
      const fromId = activity.from?.id;
      return !fromId || fromId !== conversation.userId;
    });

    if (newBotActivities.length > 0) {
      for (const activity of newBotActivities) {
        delivered.add(activity.id);
        const deltaPayload = transformActivityToDelta(activity);
        if (deltaPayload) {
          sendEvent("message.delta", deltaPayload);
          sendEvent("thread.message.delta", deltaPayload);
        }
        const completionPayload = transformActivityToCompletion(activity);
        if (completionPayload) {
          sendEvent("message.completed", completionPayload);
          sendEvent("thread.message.completed", completionPayload);
        }
      }
      sendEvent("thread.run.completed", { status: "completed" });
      sendEvent("run.completed", { status: "completed" });
      sendEvent("done", {});
      return;
    }

    if (Date.now() - startTime > DIRECT_LINE_STREAM_TIMEOUT_MS) {
      sendEvent("thread.run.completed", { status: "timeout" });
      sendEvent("run.completed", { status: "timeout" });
      sendEvent("done", {});
      return;
    }

    await delay(DIRECT_LINE_POLL_INTERVAL_MS);
  }
}

app.use(cors());
app.use(express.json());

const shouldExposeConfigEndpoint = () => {
  const raw = process.env.EXPOSE_CONFIG_ENDPOINT ?? "";
  return ["1", "true", "yes", "y", "on"].includes(raw.toLowerCase());
};

if (shouldExposeConfigEndpoint()) {
  app.get("/config", (_req, res) => {
    res.json(getPublicRuntimeConfig());
  });
}

const credential = new DefaultAzureCredential();
const clientCache = new Map();

function createClient(baseEndpoint, projectId) {
  const normalizedEndpoint = baseEndpoint.replace(/\/$/, "");
  const cacheKey = `${normalizedEndpoint}|${projectId}`;

  if (clientCache.has(cacheKey)) {
    return clientCache.get(cacheKey);
  }

  const projectEndpoint = `${normalizedEndpoint}/api/projects/${projectId}`;
  const client = new AIProjectClient(projectEndpoint, credential);
  clientCache.set(cacheKey, client);
  return client;
}

app.post("/thread", async (req, res) => {
  try {
    const providerHint = req.body.provider;
    const pluginId = req.body.pluginId;
    const resolvedContext = resolveDirectLineContext(providerHint, pluginId);

    if (resolvedContext.provider === "copilot_studio") {
      const conversation = await createDirectLineConversation(resolvedContext);
      console.log(`Created Direct Line conversation, ID: ${conversation.id}`);
      res.json({ id: conversation.id });
      return;
    }

    const { endpoint, projectId } = req.body;
    if (!endpoint || !projectId) {
      return res.status(400).json({ error: "Missing endpoint or projectId" });
    }
    const client = createClient(endpoint, projectId);
    const thread = await client.agents.threads.create();
    console.log(`Created thread, ID: ${thread.id}`);
    res.json(thread);
  } catch (err) {
    console.error("/thread error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.delete("/thread/:threadId", async (req, res) => {
  try {
    const { threadId } = req.params;
    if (!threadId) {
      return res.status(400).json({ error: "Missing threadId" });
    }

    const providerHint = req.body.provider;
    const pluginId = req.body.pluginId;
    const runtimeContext = resolveDirectLineContext(providerHint, pluginId);

    if (runtimeContext.provider === "copilot_studio") {
      const conversation = directLineConversations.get(threadId);
      if (conversation && pluginId && conversation.pluginId && pluginId !== conversation.pluginId) {
        return res.status(409).json({ error: "Thread is associated with a different plugin." });
      }
      if (directLineConversations.delete(threadId)) {
        console.log(`Deleted Direct Line conversation, ID: ${threadId}`);
      }
      res.status(204).send();
      return;
    }

    const { endpoint, projectId } = req.body;
    if (!endpoint || !projectId) {
      return res.status(400).json({ error: "Missing endpoint or projectId" });
    }
    const client = createClient(endpoint, projectId);
    if (typeof client.agents?.threads?.delete !== "function") {
      return res.status(501).json({ error: "Thread deletion is not supported by the current SDK." });
    }
    await client.agents.threads.delete(threadId);
    console.log(`Deleted thread, ID: ${threadId}`);
    res.status(204).send();
  } catch (err) {
    console.error("/thread delete error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.post("/message", async (req, res) => {
  try {
    const { threadId, role = "user", content } = req.body;
    if (!threadId || content === undefined) {
      return res.status(400).json({ error: "Missing required fields" });
    }

    const providerHint = req.body.provider;
    const pluginId = req.body.pluginId;
    const runtimeContext = resolveDirectLineContext(providerHint, pluginId);

    if (runtimeContext.provider === "copilot_studio" || directLineConversations.has(threadId)) {
      const conversation = directLineConversations.get(threadId);
      if (!conversation) {
        return res.status(404).json({ error: "Conversation not found. Please start a new thread." });
      }
      if (pluginId && conversation.pluginId && pluginId !== conversation.pluginId) {
        return res.status(409).json({ error: "Thread is associated with a different plugin." });
      }

      await ensureDirectLineToken(conversation);
      const { text, attachments } = normalizeDirectLineUserContent(content);
      if (!text && attachments.length === 0) {
        return res.status(400).json({ error: "Message content is empty" });
      }

      const activity = {
        type: "message",
        from: { id: conversation.userId },
        locale: typeof req.body.locale === "string" ? req.body.locale : "en-US",
        channelData: {
          pluginId: conversation.pluginId,
          role
        }
      };

      if (text) {
        activity.text = text;
      }

      if (attachments.length > 0) {
        activity.attachments = attachments.map((attachment) => ({
          contentType: attachment.contentType || "application/octet-stream",
          contentUrl: attachment.contentUrl,
          name: attachment.name
        }));
      }

      const response = await fetch(`${conversation.endpoint}/v3/directline/conversations/${encodeURIComponent(threadId)}/activities`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${conversation.token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(activity)
      });

      if (!response.ok) {
        const detail = await response.text();
        return res.status(response.status).json({ error: "Direct Line activity failed", details: detail });
      }

      const payload = await response.json();
      res.json(payload);
      return;
    }

    const { endpoint, projectId } = req.body;
    if (!endpoint || !projectId) {
      return res.status(400).json({ error: "Missing endpoint or projectId" });
    }
    const client = createClient(endpoint, projectId);
    const message = await client.agents.messages.create(threadId, role, content);
    res.json(message);
  } catch (err) {
    console.error("/message error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.post("/run", async (req, res) => {
  try {
    const { endpoint, projectId, threadId, agentId } = req.body;
    if (!endpoint || !projectId || !threadId || !agentId) {
      return res.status(400).json({ error: "Missing required fields" });
    }
    const client = createClient(endpoint, projectId);
    const run = await client.agents.runs.create(threadId, agentId);
    res.json(run);
  } catch (err) {
    console.error("/run error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.get("/run/:threadId/:runId", async (req, res) => {
  try {
    const { endpoint, projectId } = req.query;
    const { threadId, runId } = req.params;
    if (!endpoint || !projectId || !threadId || !runId) {
      return res.status(400).json({ error: "Missing endpoint, projectId, threadId, or runId" });
    }
    const client = createClient(endpoint, projectId);
    const run = await client.agents.runs.get(threadId, runId);
    res.json(run);
  } catch (err) {
    console.error("/run get error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.get("/messages/:threadId", async (req, res) => {
  try {
    const { threadId } = req.params;
    const endpointQuery = req.query.endpoint;
    const projectIdQuery = req.query.projectId;
    const providerQuery = req.query.provider;
    const pluginIdQuery = req.query.pluginId;

    if (!threadId) {
      return res.status(400).json({ error: "Missing threadId" });
    }

    const provider = Array.isArray(providerQuery) ? providerQuery[0] : providerQuery;
    const normalizedProvider = normalizeProvider(provider);
    const pluginId = Array.isArray(pluginIdQuery) ? pluginIdQuery[0] : pluginIdQuery;

    if (normalizedProvider === "copilot_studio" || directLineConversations.has(threadId)) {
      const conversation = directLineConversations.get(threadId);
      if (!conversation) {
        return res.status(404).json({ error: "Conversation not found" });
      }
      if (pluginId && conversation.pluginId && pluginId !== conversation.pluginId) {
        return res.status(409).json({ error: "Thread is associated with a different plugin." });
      }
      return res.status(501).json({ error: "Listing messages is not supported for Copilot Studio conversations." });
    }

    const endpoint = Array.isArray(endpointQuery) ? endpointQuery[0] : endpointQuery;
    const projectId = Array.isArray(projectIdQuery) ? projectIdQuery[0] : projectIdQuery;
    if (!endpoint || !projectId) {
      return res.status(400).json({ error: "Missing endpoint or projectId" });
    }
    const client = createClient(endpoint, projectId);
    const messageIterator = await client.agents.messages.list(threadId, { order: "asc" });
    const allMessages = [];
    for await (const message of messageIterator) {
      allMessages.push(message);
    }
    res.json({ messages: allMessages });
  } catch (err) {
    console.error("/messages error", err);
    res.status(err.statusCode || 500).json({ error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
  }
});

app.get("/run-stream", async (req, res) => {
  const endpointQuery = req.query.endpoint;
  const projectIdQuery = req.query.projectId;
  const threadIdQuery = req.query.threadId;
  const agentIdQuery = req.query.agentId;
  const providerQuery = req.query.provider;
  const pluginIdQuery = req.query.pluginId;

  const endpoint = Array.isArray(endpointQuery) ? endpointQuery[0] : endpointQuery;
  const projectId = Array.isArray(projectIdQuery) ? projectIdQuery[0] : projectIdQuery;
  const threadId = Array.isArray(threadIdQuery) ? threadIdQuery[0] : threadIdQuery;
  const agentId = Array.isArray(agentIdQuery) ? agentIdQuery[0] : agentIdQuery;
  const provider = Array.isArray(providerQuery) ? providerQuery[0] : providerQuery;
  const normalizedProvider = normalizeProvider(provider);
  const pluginId = Array.isArray(pluginIdQuery) ? pluginIdQuery[0] : pluginIdQuery;

  if (!threadId) {
    res.status(400).json({ error: "Missing threadId" });
    return;
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  const sendEvent = (eventName, payload) => {
    res.write(`event: ${eventName}\n`);
    const serialized = JSON.stringify(payload ?? {});
    res.write(`data: ${serialized}\n\n`);
  };

  let stream;
  let closed = false;

  const cleanup = async () => {
    if (closed) {
      return;
    }
    closed = true;
    try {
      if (stream?.[Symbol.asyncDispose]) {
        await stream[Symbol.asyncDispose]();
      } else if (typeof stream?.return === "function") {
        await stream.return();
      }
    } catch (disposeError) {
      console.warn("Error disposing stream", disposeError);
    }
  };

  req.on("close", async () => {
    await cleanup();
  });

  try {
    if (normalizedProvider === "copilot_studio" || directLineConversations.has(threadId)) {
      const conversation = directLineConversations.get(threadId);
      if (!conversation) {
        sendEvent("error", { error: "Conversation not found. Please start a new thread." });
        await cleanup();
        res.end();
        return;
      }
      if (pluginId && conversation.pluginId && pluginId !== conversation.pluginId) {
        sendEvent("error", { error: "Thread is associated with a different plugin." });
        await cleanup();
        res.end();
        return;
      }

      try {
        await ensureDirectLineToken(conversation);
        await streamDirectLineConversation(conversation, sendEvent, () => closed);
      } catch (err) {
        console.error("/run-stream Direct Line error", err);
        sendEvent("error", { error: err.message ?? String(err) });
      }
      await cleanup();
      res.end();
      return;
    }

    if (!endpoint || !projectId || !agentId) {
      sendEvent("error", { error: "Missing endpoint, projectId, or agentId" });
      await cleanup();
      res.end();
      return;
    }

    const client = createClient(endpoint, projectId);
    const runResponse = client.agents.runs.create(threadId, agentId);
    stream = await runResponse.stream();

    for await (const event of stream) {
      sendEvent(event.event ?? "message", event.data ?? {});
    }

    const finalRun = await runResponse;
    sendEvent("run.completed", finalRun);
    sendEvent("done", {});
    await cleanup();
    res.end();
  } catch (err) {
    console.error("/run-stream error", err);
    sendEvent("error", { error: err.message ?? String(err), details: err.response?.parsedBody ?? null });
    await cleanup();
    res.end();
  }
});

app.listen(port, () => {
  console.log(`Agent proxy server listening on http://localhost:${port}`);
});
