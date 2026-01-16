# Azure Talking Avatar Reference Experience

This repo hosts an end-to-end reference experience that blends the Azure AI Speech Talking Avatar service, Azure AI Foundry agents, Cognitive Search “on your data” enrichment, and a pluginable front-end that can render adaptive content in real time. Everything is wired through a thin pair of Node.js proxy services so secrets stay server-side while the static web client runs anywhere.

## What You Get

- **Immersive chat UI** rendered by [index.html](index.html) and bundled scripts in [dist/](dist). It supports microphone input, typed input, subtitles, rich attachments, idle video fallbacks, and agent switching with a carousel.
- **Speech + avatar streaming** driven by [js/chat-agent.js](js/chat-agent.js), which negotiates WebRTC with the Talking Avatar service, keeps STT/TTS in sync, and manages session state (reconnect, stop speaking, quick replies, chat history, etc.).
- **Configurable multi-agent routing** via plugin manifests stored beneath [plugins/](plugins). Each plugin can override avatar poses, voices, and UX branding while pointing to different Azure AI Foundry agents or Copilot Studio bots.
- **Two proxy services**:
  - [services-proxy-server](services-proxy-server) exposes `/config`, `/speech/token`, plugin utility APIs (QR generation, Blob uploads), and can relay `/agent` calls to the agent proxy.
  - [agent-proxy-server](agent-proxy-server) signs requests with `DefaultAzureCredential`, talks to Azure AI Projects, and optionally bridges Direct Line for Copilot Studio bots.
- **Deployment assets** including a multi-stage [Dockerfile](Dockerfile), automation scripts in [scripts/](scripts), and configuration templates under [config/](config).

## Repository Tour

- [css/](css) – Application styling for the landing experience, session UI, and overlay components.
- [js/](js) – Browser logic (`chat-agent.js`, `chat.js`, helpers) that orchestrates Speech SDK, WebRTC, plugins, and UI state.
- [common/](common) – Shared runtime config loader that merges JSON files, environment variables, and plugin manifests.
- [services-proxy-server/](services-proxy-server) – Express server that returns sanitized runtime config, mints Speech + relay tokens, relays agent calls, and hosts plugin helper APIs.
- [agent-proxy-server/](agent-proxy-server) – Express server that invokes Azure AI Projects (threads, runs, messages) or Copilot Studio Direct Line on behalf of the browser.
- [plugins/](plugins) – Houses the Innovation Hub sample plugin and serves as the contract for adding new personas (additional samples were removed to simplify the repo).
- [docs/](docs) – Deep dives such as architecture diagrams, session lifecycle, and Copilot Studio integration guidance.
- [scripts/](scripts) – Tooling for local dev (`start-container.js`) and Azure Container Apps deployments (`Set-EnvFromSettings.ps1`).

## Prerequisites

- Node.js 20+ and npm 10+.
- Azure subscription with:
  - Azure AI Speech resource (for Talking Avatar + token issuance).
  - Azure AI Foundry project + agent deployment **or** a Copilot Studio bot (Direct Line connection).
  - Optional: Azure Cognitive Search for “on your data” scenarios.
- Azure CLI for authentication (`az login`) and deployment automation.

## Configure Runtime Settings

1. Copy [config/settings.example.json](config/settings.example.json) to one of the following (in order of precedence):
   - `config/settings.local.json` (ignored by git)
   - `config/settings.dev.json`
   - `config/settings.json`
2. Populate the sections:
   - **speech** – Region, key, optional private endpoint, STT locales, default TTS voice, custom voice endpoint.
   - **agent** – `apiUrl` for the agent proxy, plus Azure AI Foundry endpoint/project/agent IDs or Copilot Studio Direct Line metadata.
   - **search** – Toggle `enabled` and provide Cognitive Search endpoint/key/index if you need Retrieval Augmented chats.
   - **avatar / ui / features** – Default character, style, subtitles, quick replies, overlay behavior.
   - **servicesProxyBaseUrl** – Public URL when hosting the services proxy behind an ingress.
3. Secrets can be overridden per-environment with environment variables. The mapping is documented in [scripts/Set-EnvFromSettings.ps1](scripts/Set-EnvFromSettings.ps1) and in [common/runtimeConfig.js](common/runtimeConfig.js).

## Install Dependencies and Build the Bundle

```bash
# 1. Front-end bundle + helper scripts
npm install
npx webpack --config webpack.config.js

# 2. Service proxies (install inside each folder)

npm install --prefix agent-proxy-server
```

`npx webpack` produces `dist/bundle.js`, which is referenced by [index.html](index.html). Re-run the command after UI changes.

## Run the Proxy Services Locally

```bash
# From the repo root
npm run start:container
```

`scripts/start-container.js` launches both proxies:

- **Services proxy** (defaults to `http://localhost:4100`)
  - `GET /config` → returns sanitized runtime config so the browser can auto-fill forms.
  - `POST /speech/token` → exchanges the Speech key for relay + STS tokens (handles private endpoints automatically).
  - `POST /plugin/:pluginId/api/:endpoint` → opinionated helpers for plugin uploads (Azure Blob Storage), QR codes, or forwarding HTTP requests.
  - `/:agent` (optional) → relays requests to the agent proxy when `ENABLE_AGENT_PROXY_RELAY` is true.

- **Agent proxy** (defaults to `http://localhost:4000`)
  - `POST /thread`, `POST /message`, `POST /run`, `GET /messages/:threadId` – wraps Azure AI Projects threads API.
  - Transparently switches to Copilot Studio Direct Line when a plugin manifest specifies `provider: "copilot_studio"` and supplies Direct Line credentials.
  - Uses `DefaultAzureCredential`, so authenticate with `az login`, a Managed Identity, or service principal env vars.

### Serving the Web Client

The repo is static, so you can open [index.html](index.html) with VS Code’s Live Server, `npx http-server .`, or any static host. Ensure the browser can reach the services proxy (`SERVICES_PROXY_PORT`) and agent proxy (`AGENT_PROXY_PORT`).

## Plugin System Summary

- Each plugin lives under `plugins/<id>/` with a `manifest.json`, `plugin.js`, optional scoped `styles.css`, prompts, and assets.
- Manifests can override speech, avatar, branding, and feature toggles. When the plugin is selected, its speech/stt settings supersede the host defaults (see [shared/README.md](shared/README.md)).
- [common/agentPlugins.js](common/agentPlugins.js) discovers manifests, merges overrides, and exposes them to the browser via `/config`.
- In the UI, the carousel in [index.html](index.html) lets users pick an agent; `js/chat-agent.js` hot-loads the plugin module, injects overlay HTML with `setOverlayContent`, and wires DOM events on the next animation frame.
- Example plugin:
  - [plugins/innovation-hub](plugins/innovation-hub) – End-to-end reference for overlay carousels, image assets, and command routing without relying on the legacy sample set.

## Deployment Notes

- **Container image**: The root [Dockerfile](Dockerfile) compiles the front-end in builder stage, copies static assets under `/app/static`, installs proxy dependencies with `--omit=dev`, aligns the selected config file via `CONFIG_FILE`, and finally runs `node scripts/start-container.js`. Publish the resulting image to ACR or Docker Hub.
- **Azure Container Apps / App Service**: Use [scripts/Set-EnvFromSettings.ps1](scripts/Set-EnvFromSettings.ps1) to push config as environment variables. Run with Managed Identity or supply `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `AZURE_CLIENT_SECRET` for the agent proxy.
- **Static hosting**: Any CDN or static site host can serve the files under `dist`, `css`, `js`, `image`, and `video`. Point the web app at your deployed services proxy URL via `servicesProxyBaseUrl`.

## Reference Documentation

- [docs/architecture-overview.md](docs/architecture-overview.md) – Full component architecture and data flow.
- [docs/avatar-session-lifecycle.md](docs/avatar-session-lifecycle.md) – How sessions progress from “Talk to Avatar” to shutdown.
- [docs/interaction-sequence.md](docs/interaction-sequence.md) – Sequence diagrams for Azure AI Foundry and Copilot Studio paths.
- [docs/copilot-studio-content-filtering.md](docs/copilot-studio-content-filtering.md) – Guidelines when using Copilot Studio as a provider.
- [shared/README.md](shared/README.md) – Plugin authoring guide with overlay patterns and CSS scoping tips.
- [agent-proxy-server/README.md](agent-proxy-server/README.md) – Deployment and API details for the agent proxy.

## Tips and Troubleshooting

- Set `EXPOSE_CONFIG_ENDPOINT=true` when running locally to populate the browser UI with non-secret defaults from `config/settings.*.json`.
- `services-proxy-server/src/index.js` automatically normalizes public base URLs and rewrites the agent proxy path when fronted by an ingress (`SERVICES_PROXY_PUBLIC_BASE_URL`, `AGENT_PROXY_PUBLIC_PATH`).
- For private Speech endpoints, keep `speech.enablePrivateEndpoint=true` and provide `speech.privateEndpoint`. The proxy computes the correct relay + STS URLs.
- Copilot Studio bots require Direct Line credentials in the plugin manifest or environment variables. The agent proxy refreshes tokens automatically but logs warnings if the secret is missing.
- Use `window.DEBUG_PLUGINS = true` in the browser console to dump plugin lifecycle logs buffered by `js/chat-agent.js`.

---

This README now reflects the actual structure of the repo (single-page chat experience, plugin host, services + agent proxies) and should be the starting point for anyone deploying or customizing the Azure Talking Avatar reference app.
- **Example Plugin**: Study `innovation-hub` for a production-ready implementation pattern
