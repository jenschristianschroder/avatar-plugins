import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { loadAgentPlugins } from "./agentPlugins.js";

const currentFilePath = fileURLToPath(import.meta.url);
const currentDir = path.dirname(currentFilePath);
const CONFIG_DIRS = [
  path.resolve(currentDir, "..", "services-proxy-server", "config"),
  path.resolve(currentDir, "..", "agent-proxy-server", "config"),
  path.resolve(currentDir, "..", "config")
];
const CONFIG_FILES = ["settings.json"];
const DEFAULT_AGENT_PROVIDER = "azure_ai_foundry";

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

function mergeDirectLineConfig(target, source) {
  if (!target || typeof target !== "object") {
    target = {};
  }
  if (!source || typeof source !== "object") {
    return target;
  }

  const assign = (key, value) => {
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (trimmed) {
        target[key] = trimmed;
      }
    }
  };

  assign("endpoint", source.directLineEndpoint ?? source.endpoint);
  assign("botId", source.directLineBotId ?? source.botId ?? source.agentId);
  assign("userId", source.directLineUserId ?? source.userId);
  assign("secret", source.directLineSecret ?? source.secret);
  assign("secretEnv", source.directLineSecretEnv ?? source.secretEnv);
  assign("scope", source.directLineScope ?? source.scope);
  assign("region", source.directLineRegion ?? source.region);

  return target;
}

const defaultConfig = {
  speech: {
    region: "eastus2",
    apiKey: "",
    enablePrivateEndpoint: false,
    privateEndpoint: "",
    sttLocales: "en-US,da-DK",
    ttsVoice: "en-US-AvaMultilingualNeural",
    customVoiceEndpointId: ""
  },
  agent: {
    apiUrl: "http://localhost:4000",
    options: []
  },
  branding: {
    primaryColor: "#001f76",
    backgroundImage: "",
    logoUrl: "",
    logoAlt: ""
  },
  search: {
    enabled: false,
    endpoint: "",
    apiKey: "",
    indexName: ""
  },
  avatar: {
    character: "lisa",
    style: "casual-sitting",
    customized: false,
    useBuiltInVoice: false,
    autoReconnect: false,
    useLocalVideoForIdle: false,
    backgroundImage: "",
    transparentBackground: false
  },
  ui: {
    showSubtitles: false
  },
  conversation: {
    continuous: false
  },
  features: {
    quickReplyEnabled: false,
    quickReplyOptions: [
      "Let me take a look.",
      "Let me check.",
      "One moment, please."
    ]
  },
  servicesProxyBaseUrl: "http://localhost:4100"
};

function deepMerge(target, source) {
  if (!source || typeof source !== "object") {
    return target;
  }

  Object.entries(source).forEach(([key, value]) => {
    if (Array.isArray(value)) {
      target[key] = value.slice();
    } else if (value && typeof value === "object") {
      if (!target[key] || typeof target[key] !== "object") {
        target[key] = {};
      }
      deepMerge(target[key], value);
    } else if (value !== undefined) {
      target[key] = value;
    }
  });

  return target;
}

function readConfigFile(fileName) {
  for (const dir of CONFIG_DIRS) {
    const filePath = path.join(dir, fileName);
    if (!fs.existsSync(filePath)) {
      continue;
    }

    try {
      const content = fs.readFileSync(filePath, "utf8");
      if (!content.trim()) {
        return {};
      }
      return JSON.parse(content);
    } catch (err) {
      console.warn(`Failed to read config file ${fileName} in ${dir}`, err);
      return {};
    }
  }

  return {};
}

function applyAgentPluginConfig(config) {
  if (!config.agent || typeof config.agent !== "object") {
    config.agent = {};
  }

  const existingPluginOverrides = config.agent && typeof config.agent.plugins === "object"
    ? config.agent.plugins
    : null;

  const pluginPayload = loadAgentPlugins();
  const { plugins, options, defaultPluginId, defaultConnection } = pluginPayload;

  if (existingPluginOverrides) {
    Object.entries(existingPluginOverrides).forEach(([pluginId, override]) => {
      if (!override || typeof override !== "object") {
        return;
      }
      const pluginRecord = plugins?.[pluginId];
      if (!pluginRecord) {
        return;
      }

      if (override.config && typeof override.config === "object") {
        if (!pluginRecord.config || typeof pluginRecord.config !== "object") {
          pluginRecord.config = {};
        }
        deepMerge(pluginRecord.config, override.config);

        const optionRecord = Array.isArray(options)
          ? options.find((option) => option && (option.pluginId === pluginId || option.key === pluginId))
          : null;
        if (optionRecord) {
          if (!optionRecord.config || typeof optionRecord.config !== "object") {
            optionRecord.config = {};
          }
          deepMerge(optionRecord.config, override.config);
        }
      }
    });
  }

  config.agent.plugins = plugins;
  config.agent.options = options;

  if (defaultPluginId && !config.agent.defaultPluginId) {
    config.agent.defaultPluginId = defaultPluginId;
  }

  if (!config.agent.selectedPluginId && config.agent.defaultPluginId) {
    config.agent.selectedPluginId = config.agent.defaultPluginId;
  }

  if (defaultConnection) {
    const defaultProvider = normalizeProvider(defaultConnection.provider) || DEFAULT_AGENT_PROVIDER;
    config.agent.provider = defaultProvider;

    if (!config.agent.apiUrl && defaultConnection.apiUrl) {
      config.agent.apiUrl = defaultConnection.apiUrl;
    }
    if (!config.agent.systemPrompt && defaultConnection.systemPrompt) {
      config.agent.systemPrompt = defaultConnection.systemPrompt;
    }

    if (defaultProvider === "copilot_studio") {
      if (!config.agent.directLine || typeof config.agent.directLine !== "object") {
        config.agent.directLine = {};
      }
      mergeDirectLineConfig(config.agent.directLine, defaultConnection);
      if (!config.agent.endpoint && defaultConnection.endpoint) {
        config.agent.endpoint = defaultConnection.endpoint;
      }
    } else {
      if (!config.agent.endpoint && defaultConnection.endpoint) {
        config.agent.endpoint = defaultConnection.endpoint;
      }
      if (!config.agent.projectId && defaultConnection.projectId) {
        config.agent.projectId = defaultConnection.projectId;
      }
      if (!config.agent.agentId && defaultConnection.agentId) {
        config.agent.agentId = defaultConnection.agentId;
      }
    }
  }

  const selectedPluginId = config.agent.selectedPluginId;
  const pluginRecord = selectedPluginId ? plugins?.[selectedPluginId] : undefined;
  if (!pluginRecord) {
    return config;
  }

  const { connection = {}, branding, avatar, speech, features, ui, conversation, assets, content, languages } = pluginRecord;

  const provider = normalizeProvider(connection.provider || config.agent.provider) || DEFAULT_AGENT_PROVIDER;
  config.agent.provider = provider;

  if (connection.apiUrl) {
    config.agent.apiUrl = connection.apiUrl;
  }
  if (connection.systemPrompt) {
    config.agent.systemPrompt = connection.systemPrompt;
  }

  if (provider === "copilot_studio") {
    if (!config.agent.directLine || typeof config.agent.directLine !== "object") {
      config.agent.directLine = {};
    }
    mergeDirectLineConfig(config.agent.directLine, connection);
    const directLineEndpoint = connection.directLineEndpoint || connection.endpoint;
    if (directLineEndpoint) {
      config.agent.directLine.endpoint = directLineEndpoint;
    }
    config.agent.endpoint = directLineEndpoint || config.agent.endpoint || "";
    config.agent.projectId = "";
    config.agent.agentId = "";
  } else {
    if (connection.endpoint) {
      config.agent.endpoint = connection.endpoint;
    }
    if (connection.projectId) {
      config.agent.projectId = connection.projectId;
    }
    if (connection.agentId) {
      config.agent.agentId = connection.agentId;
    }
  }

  if (branding) {
    if (!config.branding || typeof config.branding !== "object") {
      config.branding = {};
    }
    deepMerge(config.branding, branding);
  }

  if (avatar) {
    if (!config.avatar || typeof config.avatar !== "object") {
      config.avatar = {};
    }
    deepMerge(config.avatar, avatar);
  }

  if (speech) {
    if (!config.speech || typeof config.speech !== "object") {
      config.speech = {};
    }
    deepMerge(config.speech, speech);
  }

  if (features) {
    if (!config.features || typeof config.features !== "object") {
      config.features = {};
    }
    deepMerge(config.features, features);
  }

  if (ui) {
    if (!config.ui || typeof config.ui !== "object") {
      config.ui = {};
    }
    deepMerge(config.ui, ui);
  }

  if (conversation) {
    if (!config.conversation || typeof config.conversation !== "object") {
      config.conversation = {};
    }
    deepMerge(config.conversation, conversation);
  }

  if (assets) {
    if (!config.assets || typeof config.assets !== "object") {
      config.assets = {};
    }
    deepMerge(config.assets, assets);
  }

  if (content) {
    if (!config.content || typeof config.content !== "object") {
      config.content = {};
    }
    deepMerge(config.content, content);
  }

  if (Array.isArray(languages) && languages.length > 0) {
    config.agent.languages = languages.slice();
  }

  return config;
}

function normalizeBoolean(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "string") {
    const lowered = value.trim().toLowerCase();
    if (["1", "true", "yes", "y", "on"].includes(lowered)) {
      return true;
    }
    if (["0", "false", "no", "n", "off"].includes(lowered)) {
      return false;
    }
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  return undefined;
}

function normalizeStringArray(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry)).filter((entry) => entry.trim() !== "");
  }

  if (typeof value === "string") {
    const parts = value
      .split(/[|,]/)
      .map((part) => part.trim())
      .filter((part) => part !== "");
    return parts.length > 0 ? parts : undefined;
  }

  return undefined;
}

function setDeep(target, pathSegments, value) {
  if (value === undefined) {
    return;
  }
  let cursor = target;
  for (let i = 0; i < pathSegments.length - 1; i += 1) {
    const segment = pathSegments[i];
    if (!cursor[segment] || typeof cursor[segment] !== "object") {
      cursor[segment] = {};
    }
    cursor = cursor[segment];
  }
  cursor[pathSegments[pathSegments.length - 1]] = value;
}

const ENVIRONMENT_MAPPINGS = [
  { env: "SPEECH_RESOURCE_REGION", path: ["speech", "region"] },
  { env: "SPEECH_RESOURCE_KEY", path: ["speech", "apiKey"] },
  { env: "SPEECH_PRIVATE_ENDPOINT", path: ["speech", "privateEndpoint"] },
  { env: "ENABLE_PRIVATE_ENDPOINT", path: ["speech", "enablePrivateEndpoint"], converter: normalizeBoolean },
  { env: "STT_LOCALES", path: ["speech", "sttLocales"] },
  { env: "TTS_VOICE", path: ["speech", "ttsVoice"] },
  { env: "CUSTOM_VOICE_ENDPOINT_ID", path: ["speech", "customVoiceEndpointId"] },
  { env: "AGENT_API_URL", path: ["agent", "apiUrl"] },
  { env: "ENABLE_ON_YOUR_DATA", path: ["search", "enabled"], converter: normalizeBoolean },
  { env: "COG_SEARCH_ENDPOINT", path: ["search", "endpoint"] },
  { env: "COG_SEARCH_API_KEY", path: ["search", "apiKey"] },
  { env: "COG_SEARCH_INDEX_NAME", path: ["search", "indexName"] },
  { env: "AVATAR_CHARACTER", path: ["avatar", "character"] },
  { env: "AVATAR_STYLE", path: ["avatar", "style"] },
  { env: "AVATAR_CUSTOMIZED", path: ["avatar", "customized"], converter: normalizeBoolean },
  { env: "AVATAR_USE_BUILT_IN_VOICE", path: ["avatar", "useBuiltInVoice"], converter: normalizeBoolean },
  { env: "AVATAR_AUTO_RECONNECT", path: ["avatar", "autoReconnect"], converter: normalizeBoolean },
  { env: "AVATAR_USE_LOCAL_VIDEO_FOR_IDLE", path: ["avatar", "useLocalVideoForIdle"], converter: normalizeBoolean },
  { env: "AVATAR_BACKGROUND_IMAGE", path: ["avatar", "backgroundImage"] },
  { env: "AVATAR_TRANSPARENT_BACKGROUND", path: ["avatar", "transparentBackground"], converter: normalizeBoolean },
  { env: "SHOW_SUBTITLES", path: ["ui", "showSubtitles"] },
  { env: "CONTINUOUS_CONVERSATION", path: ["conversation", "continuous"] },
  { env: "ENABLE_QUICK_REPLY", path: ["features", "quickReplyEnabled"], converter: normalizeBoolean },
  { env: "QUICK_REPLY_OPTIONS", path: ["features", "quickReplyOptions"], converter: normalizeStringArray },
  { env: "SERVICES_PROXY_PUBLIC_BASE_URL", path: ["servicesProxyBaseUrl"] },
  { env: "BRANDING_PRIMARY_COLOR", path: ["branding", "primaryColor"] },
  { env: "BRANDING_BACKGROUND_IMAGE", path: ["branding", "backgroundImage"] },
  { env: "BRANDING_LOGO_URL", path: ["branding", "logoUrl"] },
  { env: "BRANDING_LOGO_ALT", path: ["branding", "logoAlt"] }
];

function loadFileConfig() {
  const aggregate = {};
  CONFIG_FILES.forEach((fileName) => {
    const fileConfig = readConfigFile(fileName);
    if (fileConfig && typeof fileConfig === "object") {
      deepMerge(aggregate, fileConfig);
    }
  });
  return aggregate;
}

function applyEnvironmentOverrides(baseConfig) {
  const merged = JSON.parse(JSON.stringify(baseConfig));

  ENVIRONMENT_MAPPINGS.forEach(({ env, path: pathSegments, converter }) => {
    if (!(env in process.env)) {
      return;
    }
    const rawValue = process.env[env];
    const value = converter ? converter(rawValue) : rawValue;
    if (value === undefined) {
      return;
    }
    setDeep(merged, pathSegments, value);
  });

  return merged;
}

export function getRuntimeConfig() {
  const fileConfig = loadFileConfig();
  const baseConfig = deepMerge(JSON.parse(JSON.stringify(defaultConfig)), fileConfig);
  const withPlugins = applyAgentPluginConfig(baseConfig);
  return applyEnvironmentOverrides(withPlugins);
}

/**
 * Recursively removes API keys from nested objects
 * Handles dynamic plugin configurations
 */
function sanitizeNestedApiKeys(obj) {
  if (!obj || typeof obj !== "object") {
    return;
  }

  Object.keys(obj).forEach((key) => {
    // Remove any property with 'apiKey', 'api_key', 'secret', 'password', 'token' in the name
    const keyLower = key.toLowerCase();
    if (keyLower.includes('apikey') || keyLower.includes('api_key') || 
        keyLower === 'secret' || keyLower === 'password' || keyLower === 'token') {
      delete obj[key];
    } else if (obj[key] && typeof obj[key] === "object" && !Array.isArray(obj[key])) {
      // Recursively sanitize nested objects
      sanitizeNestedApiKeys(obj[key]);
    }
  });
}

export function sanitizeRuntimeConfig(config) {
  const clone = JSON.parse(JSON.stringify(config ?? {}));

  if (clone.speech) {
    delete clone.speech.apiKey;
  }

  if (clone.search) {
    delete clone.search.apiKey;
  }

  if (clone.agent && typeof clone.agent === "object") {
    if (clone.agent.directLine && typeof clone.agent.directLine === "object") {
      delete clone.agent.directLine.secret;
      delete clone.agent.directLine.token;
    }

    if (clone.agent.connection && typeof clone.agent.connection === "object") {
      delete clone.agent.connection.directLineSecret;
    }

    if (clone.agent.plugins && typeof clone.agent.plugins === "object") {
      Object.values(clone.agent.plugins).forEach((plugin) => {
        if (plugin && typeof plugin === "object") {
          // Remove secrets from connection
          if (plugin.connection && typeof plugin.connection === "object") {
            delete plugin.connection.directLineSecret;
          }
          
          // Recursively sanitize any API keys in plugin config
          // This handles dynamic plugin configurations like imageGeneration.apiKey
          sanitizeNestedApiKeys(plugin);
        }
      });
    }

    if (Array.isArray(clone.agent.options)) {
      clone.agent.options = clone.agent.options.map((option) => {
        if (!option || typeof option !== "object") {
          return option;
        }
        const sanitized = { ...option };
        if (sanitized.directLine && typeof sanitized.directLine === "object") {
          delete sanitized.directLine.secret;
        }
        return sanitized;
      });
    }
  }

  return clone;
}

export function getPublicRuntimeConfig() {
  return sanitizeRuntimeConfig(getRuntimeConfig());
}