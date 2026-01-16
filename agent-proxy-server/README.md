# Agent Proxy Server

This Node.js service brokers calls between the browser sample and Azure AI Projects. It uses `DefaultAzureCredential`, so the account running the server must be signed in (for example with `az login`) or have another supported credential provider available.

build: docker build -t avatar-agent-proxy -f Dockerfile .
run with: docker run -p 4000:4000 --env AZURE_CLIENT_ID=… --env AZURE_TENANT_ID=… --env AZURE_CLIENT_SECRET=… avatar-agent-proxy

or rely on managed identity in Azure App Service/Container Apps

## Setup

```powershell
cd server
npm install
```

Ensure you are authenticated:

```powershell
az login --use-device-code
```

## Run in development

```powershell
npm run dev
```

Or run normally:

```powershell
npm start
```

The server listens on `http://localhost:4000` by default. Adjust the `PORT` environment variable if needed.

### Configuration

Create a `server/config/settings.local.json` file (ignored by git) to provide development defaults, based on the structure in `server/config/settings.example.json`. When running in Azure Container Apps, supply the same values via environment variables instead:

- `SPEECH_RESOURCE_REGION`, `SPEECH_RESOURCE_KEY`, `SPEECH_PRIVATE_ENDPOINT`
- `AZURE_AI_FOUNDRY_ENDPOINT`, `AZURE_AI_FOUNDRY_AGENT_ID`, `AZURE_AI_FOUNDRY_PROJECT_ID`
- `AGENT_API_URL`, `SYSTEM_PROMPT`
- `ENABLE_PRIVATE_ENDPOINT`, `ENABLE_ON_YOUR_DATA`
- `COG_SEARCH_ENDPOINT`, `COG_SEARCH_API_KEY`, `COG_SEARCH_INDEX_NAME`
- `STT_LOCALES`, `TTS_VOICE`, `CUSTOM_VOICE_ENDPOINT_ID`

By default the server **does not** expose configuration data over HTTP. If you want the web client to auto-populate fields during local development, opt-in explicitly:

```powershell
$env:EXPOSE_CONFIG_ENDPOINT = "true"
npm run dev
```

When the endpoint is enabled it returns only non-secret values (API keys are removed).

## API

- `POST /thread` → creates an empty thread.
- `POST /message` → posts a message to a thread.
- `POST /run` → starts a run for the thread+agent.
- `GET /messages/:threadId` → lists messages in the thread.
- `GET /config` → (optional) returns a sanitized runtime configuration when `EXPOSE_CONFIG_ENDPOINT=true`.

Each request body must include:

```json
{
  "endpoint": "https://<your-project>.services.ai.azure.com",
  "projectId": "<your-project-id>",
  "...": "other fields as required"
}
```

The browser client should call these endpoints instead of the Azure SDK directly.
