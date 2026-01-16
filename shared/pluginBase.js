const OVERLAY_API_CANDIDATES = [
  "updateAttachmentOverlay",
  "setAttachmentOverlay",
  "setOverlayContent",
  "updateOverlayContent"
];

function deepClone(value) {
  if (value === null || typeof value !== "object") {
    return value;
  }
  if (typeof structuredClone === "function") {
    try {
      return structuredClone(value);
    } catch (_) {
      // Ignore structuredClone failures and fall back to JSON.
    }
  }
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (_) {
    return value;
  }
}

function normalizeManifest(manifest) {
  if (!manifest || typeof manifest !== "object") {
    return {};
  }
  return { ...manifest };
}

function normalizeApi(api) {
  if (!api || typeof api !== "object") {
    return {};
  }
  return { ...api };
}

function coerceText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value.trim();
  }
  return String(value).trim();
}

function attemptJson(value) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  if (!trimmed.startsWith("{") && !trimmed.startsWith("[")) {
    return null;
  }
  try {
    return JSON.parse(trimmed);
  } catch (_) {
    return null;
  }
}

function extractContentFromPayload(payload) {
  const queue = [payload];
  const visited = new WeakSet();

  while (queue.length > 0) {
    const current = queue.shift();
    if (current === null || current === undefined) {
      continue;
    }

    if (typeof current === "string") {
      const parsed = attemptJson(current);
      if (parsed !== null) {
        queue.push(parsed);
      }
      continue;
    }

    if (typeof current !== "object") {
      continue;
    }

    if (visited.has(current)) {
      continue;
    }
    visited.add(current);

    if (Object.prototype.hasOwnProperty.call(current, "content")) {
      return current.content;
    }

    const nestedKeys = [
      "data",
      "response",
      "message",
      "payload",
      "result",
      "value",
      "body",
      "output"
    ];

    for (const key of nestedKeys) {
      if (current[key] !== undefined) {
        queue.push(current[key]);
      }
    }
  }

  return undefined;
}

function enforceDomRestriction() {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return;
  }
  if (window.__PLUGIN_DOM_GUARD__) {
    return;
  }
  Object.defineProperty(window, "__PLUGIN_DOM_GUARD__", {
    value: true,
    configurable: false,
    enumerable: false,
    writable: false
  });

  const descriptor = {
    configurable: false,
    enumerable: false,
    get() {
      throw new Error("Access to DOM is forbidden in plugins.");
    }
  };

  try {
    Object.defineProperty(window, "document", descriptor);
  } catch (_) {
    // Silently ignore if the host has already locked the descriptor.
  }
}

export class PluginBase {
  constructor(manifest, api) {
    this._manifest = Object.freeze(normalizeManifest(manifest));
    this._api = normalizeApi(api);
    this._overlayContent = null;
    this._destroyed = false;
    this._agentContentUnsubscribe = null;

    enforceDomRestriction();
    this._bindAgentContentListener();
  }

  get manifest() {
    return this._manifest;
  }

  get id() {
    const candidate = this._manifest.id || this._manifest.name || this._manifest.key;
    return typeof candidate === "string" ? candidate : null;
  }

  get api() {
    return this._api;
  }

  get config() {
    const fromApi = this._api.config;
    if (fromApi && typeof fromApi === "object") {
      return fromApi;
    }
    const fromManifest = this._manifest.config;
    return fromManifest && typeof fromManifest === "object" ? fromManifest : {};
  }

  get isDestroyed() {
    return this._destroyed;
  }

  get overlayContent() {
    return this._overlayContent;
  }

  init() {
    throw new Error("init() must be implemented by the plugin.");
  }

  destroy() {
    if (this._destroyed) {
      return;
    }

    this.clearOverlayContent();

    if (typeof this._agentContentUnsubscribe === "function") {
      try {
        this._agentContentUnsubscribe();
      } catch (err) {
        this._log("warn", "Agent content unsubscribe failed", err);
      }
    }

    this._agentContentUnsubscribe = null;
    this._destroyed = true;
  }

  sendTextInput(text, metadata = {}) {
    if (this._destroyed) {
      this._log("warn", "sendTextInput ignored after destroy");
      return false;
    }

    const normalized = coerceText(text);
    if (!normalized) {
      return false;
    }

    const sender = this._api.sendTextInput
      || this._api.emitTextInput
      || this._api.sendUserText
      || this._api.submitText;

    if (typeof sender !== "function") {
      this._log("warn", "sendTextInput requested but host API is unavailable");
      return false;
    }

    const context = {
      pluginId: this.id,
      manifest: this._manifest,
      metadata: metadata && typeof metadata === "object" ? { ...metadata } : {},
      text: normalized
    };

    try {
      const result = sender(normalized, context);
      return result === undefined ? true : result;
    } catch (err) {
      this._log("error", "sendTextInput threw", err);
      throw err;
    }
  }

  setOverlayContent(contentDescriptor) {
    if (this._destroyed) {
      return;
    }
    this._overlayContent = contentDescriptor === undefined ? null : deepClone(contentDescriptor);
    this._dispatchOverlayUpdate();
  }

  clearOverlayContent() {
    if (this._overlayContent === null) {
      return;
    }
    this._overlayContent = null;
    this._dispatchOverlayUpdate();
  }

  onAgentContent(_content, _context) {
    // Plugins can override to react to agent-produced content payloads.
  }

  _dispatchOverlayUpdate() {
    if (!this._api) {
      return;
    }

    const payload = {
      pluginId: this.id,
      manifest: this._manifest,
      content: this._overlayContent
    };

    for (const key of OVERLAY_API_CANDIDATES) {
      const candidate = this._api[key];
      if (typeof candidate !== "function") {
        continue;
      }
      try {
        candidate(payload.content, payload);
        return;
      } catch (err) {
        this._log("error", `Overlay update via ${key} failed`, err);
      }
    }
  }

  _bindAgentContentListener() {
    const subscriber = this._api.onAgentContent
      || this._api.onAgentResponse
      || this._api.subscribeToAgentContent
      || this._api.registerAgentContentHandler;

    if (typeof subscriber !== "function") {
      return;
    }

    const handler = (payload) => {
      if (this._destroyed) {
        return;
      }
      const content = extractContentFromPayload(payload);
      if (content === undefined) {
        return;
      }
      try {
        this.onAgentContent(content, {
          pluginId: this.id,
          manifest: this._manifest,
          rawPayload: payload
        });
      } catch (err) {
        this._log("error", "onAgentContent handler threw", err);
      }
    };

    try {
      const unsubscribe = subscriber(handler, {
        pluginId: this.id,
        manifest: this._manifest
      });

      if (typeof unsubscribe === "function") {
        this._agentContentUnsubscribe = unsubscribe;
      } else if (unsubscribe && typeof unsubscribe.unsubscribe === "function") {
        this._agentContentUnsubscribe = () => unsubscribe.unsubscribe();
      }
    } catch (err) {
      this._log("error", "Failed to bind agent content listener", err);
    }
  }

  _log(level, message, detail) {
    const logger = this._api && this._api.log;
    if (typeof logger === "function") {
      try {
        if (detail !== undefined) {
          logger(`[PluginBase:${this.id || "unknown"}] ${message}`, { level, detail });
        } else {
          logger(`[PluginBase:${this.id || "unknown"}] ${message}`, { level });
        }
        return;
      } catch (_) {
        // Fall back to console below.
      }
    }

    if (typeof console !== "undefined" && typeof console[level] === "function") {
      console[level](`[PluginBase] ${message}`, detail);
      return;
    }

    if (typeof console !== "undefined" && typeof console.log === "function") {
      console.log(`[PluginBase] ${message}`, detail);
    }
  }
}

// Only the PluginBase class is implemented in this file as per architectural guidelines.