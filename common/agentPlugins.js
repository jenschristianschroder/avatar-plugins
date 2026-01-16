import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const currentFilePath = fileURLToPath(import.meta.url);
const currentDir = path.dirname(currentFilePath);

const PLUGIN_DIRS = [
  path.resolve(currentDir, "..", "plugins")
];

function getManifestPriority(sourcePath) {
  const fileName = path.basename(sourcePath).toLowerCase();
  if (fileName.includes(".example.") || fileName.endsWith(".example.json")
    || fileName.includes(".sample.") || fileName.endsWith(".sample.json")
    || fileName.includes(".template.")) {
    return 0;
  }
  if (fileName.includes(".local.")) {
    return 3;
  }
  if (fileName === "manifest.json") {
    return 2;
  }
  return 1;
}

function walkDirectory(dirPath, collected = []) {
  if (!fs.existsSync(dirPath)) {
    return collected;
  }
  const entries = fs.readdirSync(dirPath, { withFileTypes: true });
  entries.forEach((entry) => {
    const entryPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      walkDirectory(entryPath, collected);
      return;
    }
    if (entry.isFile() && entry.name.toLowerCase().endsWith(".json")) {
      collected.push(entryPath);
    }
  });
  return collected;
}

function normalizeString(value) {
  if (typeof value !== "string") {
    return "";
  }
  return value.trim();
}

function normalizeStringArray(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeString(entry))
      .filter((entry) => entry !== "");
  }
  const normalized = normalizeString(value);
  if (!normalized) {
    return [];
  }
  return normalized
    .split(/[|,]/)
    .map((entry) => entry.trim())
    .filter((entry) => entry !== "");
}

function deepClone(value) {
  if (value === undefined) {
    return undefined;
  }
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (_) {
    return undefined;
  }
}

function sanitizeSection(value) {
  if (!value || typeof value !== "object") {
    return {};
  }
  return deepClone(value) ?? {};
}

function buildPluginRecord(manifest, sourcePath) {
  const pluginId = normalizeString(manifest.id) || path.basename(sourcePath, path.extname(sourcePath));
  if (!pluginId) {
    return null;
  }

  const label = normalizeString(manifest.label) || pluginId;
  const description = normalizeString(manifest.description);
  const defaultFlag = manifest.default === true;
  const badge = normalizeString(manifest.badge);
  const thumbnail = normalizeString(manifest.thumbnail);
  const image = normalizeString(manifest.image);

  const connectionRaw = manifest.connection && typeof manifest.connection === "object" ? manifest.connection : {};
  const providerRaw = normalizeString(
    connectionRaw.provider
      || connectionRaw.type
      || manifest.provider
      || manifest.backend
  );
  const normalizedProvider = providerRaw ? providerRaw.toLowerCase() : "";
  const connection = {
    provider: normalizedProvider || "azure_ai_foundry",
    endpoint: normalizeString(connectionRaw.endpoint),
    projectId: normalizeString(connectionRaw.projectId),
    agentId: normalizeString(connectionRaw.agentId || connectionRaw.deploymentId || manifest.agentId || manifest.deploymentId),
    apiUrl: normalizeString(connectionRaw.apiUrl),
    systemPrompt: normalizeString(connectionRaw.systemPrompt || manifest.systemPrompt),
    directLineEndpoint: normalizeString(connectionRaw.directLineEndpoint || connectionRaw.directLineBaseUrl),
    directLineSecret: normalizeString(connectionRaw.directLineSecret),
    directLineSecretEnv: normalizeString(connectionRaw.directLineSecretEnv),
    directLineBotId: normalizeString(connectionRaw.directLineBotId || connectionRaw.botId),
    directLineUserId: normalizeString(connectionRaw.directLineUserId || connectionRaw.userId),
    directLineScope: normalizeString(connectionRaw.directLineScope || connectionRaw.scope),
    directLineRegion: normalizeString(connectionRaw.directLineRegion)
  };

  if (!connection.agentId && connection.provider !== "copilot_studio") {
    return null;
  }

  // Normalize known properties
  const branding = sanitizeSection(manifest.branding);
  const avatar = sanitizeSection(manifest.avatar);
  const speech = sanitizeSection(manifest.speech);
  const features = sanitizeSection(manifest.features);
  const ui = sanitizeSection(manifest.ui);
  const conversation = sanitizeSection(manifest.conversation);
  const assets = sanitizeSection(manifest.assets);
  const config = sanitizeSection(manifest.config);
  const content = sanitizeSection(manifest.content ?? manifest.handlers);
  const languages = normalizeStringArray(manifest.languages ?? manifest.languageSupport ?? connectionRaw.languages);

  const priority = getManifestPriority(sourcePath);

  // Start with a deep clone of the entire manifest to preserve all custom properties
  const plugin = deepClone(manifest) || {};
  
  // Override with normalized known properties
  plugin.id = pluginId;
  plugin.label = label;
  plugin.description = description;
  plugin.default = defaultFlag;
  plugin.badge = badge;
  plugin.thumbnail = thumbnail;
  plugin.image = image;
  plugin.connection = connection;
  plugin.provider = connection.provider;
  plugin.branding = branding;
  plugin.avatar = avatar;
  plugin.speech = speech;
  plugin.features = features;
  plugin.ui = ui;
  plugin.conversation = conversation;
  plugin.assets = assets;
  plugin.config = config;
  plugin.content = content;
  plugin.languages = languages;
  
  // Remove any properties that might conflict with normalized ones
  delete plugin.agentId;
  delete plugin.deploymentId;
  delete plugin.backend;
  delete plugin.languageSupport;
  delete plugin.handlers;

  const optionBranding = deepClone(branding) ?? {};
  const optionConfig = deepClone(config) ?? {};

  const option = {
    key: pluginId,
    id: connection.agentId || pluginId,
    pluginId,
    label,
    description,
    badge,
    agentId: connection.agentId,
    projectId: connection.projectId,
    endpoint: connection.endpoint,
    apiUrl: connection.apiUrl,
    provider: connection.provider,
    directLine: deepClone({
      endpoint: connection.directLineEndpoint,
      botId: connection.directLineBotId,
      region: connection.directLineRegion
    }) ?? {},
    branding: optionBranding,
    config: optionConfig,
    default: defaultFlag === true
  };

  return { plugin, option, priority, sourcePath };
}

export function loadAgentPlugins() {
  const pluginRecordsById = new Map();
  let discoveryOrder = 0;

  PLUGIN_DIRS.forEach((dirPath) => {
    walkDirectory(dirPath).forEach((filePath) => {
      try {
        // Skip example, sample, and template files entirely
        const fileName = path.basename(filePath).toLowerCase();
        if (fileName.includes(".example.") || fileName.endsWith(".example.json")
          || fileName.includes(".sample.") || fileName.endsWith(".sample.json")
          || fileName.includes(".template.")) {
          return;
        }

        const raw = fs.readFileSync(filePath, "utf8");
        if (!raw.trim()) {
          return;
        }
        const parsed = JSON.parse(raw);
        const record = buildPluginRecord(parsed, filePath);
        if (!record) {
          console.warn(`Agent plugin manifest skipped: ${filePath}`);
          return;
        }
        record.order = discoveryOrder;
        discoveryOrder += 1;

        const existing = pluginRecordsById.get(record.plugin.id);
        if (!existing) {
          pluginRecordsById.set(record.plugin.id, record);
          return;
        }

        if (record.priority > existing.priority) {
          record.order = Math.min(record.order, existing.order);
          pluginRecordsById.set(record.plugin.id, record);
          return;
        }

        if (record.priority === existing.priority) {
          const keepExisting = existing.order <= record.order;
          if (!keepExisting) {
            record.order = Math.min(record.order, existing.order);
            pluginRecordsById.set(record.plugin.id, record);
          }
          return;
        }
        // Lower priority records are ignored but we still emit a warning for visibility.
        console.warn(`Agent plugin manifest '${filePath}' ignored in favor of higher priority manifest for '${record.plugin.id}'`);
      } catch (err) {
        console.warn(`Failed to load agent plugin manifest ${filePath}`, err);
      }
    });
  });

  const plugins = {};
  const options = [];
  let defaultPluginId = null;
  let defaultConnection = null;

  const pluginRecords = Array.from(pluginRecordsById.values()).sort((a, b) => a.order - b.order);

  pluginRecords.forEach(({ plugin, option }) => {
    plugins[plugin.id] = plugin;
    options.push(option);
    if (plugin.default && !defaultPluginId) {
      defaultPluginId = plugin.id;
      defaultConnection = plugin.connection;
    }
  });

  if (!defaultPluginId && pluginRecords.length > 0) {
    const first = pluginRecords[0];
    defaultPluginId = first.plugin.id;
    defaultConnection = first.plugin.connection;
  }

  return {
    plugins,
    options,
    defaultPluginId,
    defaultConnection
  };
}
