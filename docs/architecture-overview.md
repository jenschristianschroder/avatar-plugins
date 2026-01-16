# Avatar Architecture Overview

This diagram shows the overall architecture of the avatar application, including all major components and their interactions.

```mermaid
flowchart TD
    subgraph Frontend["Frontend Layer"]
        Browser["Browser\nchat_agent.html + chat-agent.js"]
        PluginHost["Plugin Host\n(pluginBase.js)"]
        Overlay["Attachment Overlay\n(custom HTML rendering)"]
    end

    subgraph ServicesProxy["services-proxy-server (Express)"]
        ConfigLoader["Load config\ncommon/runtimeConfig.js"]
        ProxyEndpoint["/agent → http-proxy\n(WebSocket + HTTP)"]
        SpeechToken["/speech/token\nfetch Azure Speech tokens"]
        StaticAssets["Static Assets\n(HTML, JS, CSS)"]
    end

    subgraph AgentProxy["agent-proxy-server (Express)"]
        AuthProvider["DefaultAzureCredential\n(Managed Identity)"]
        FoundryClient["AIProjectClient\n(@azure/ai-projects)"]
        DirectLineClient["Direct Line Client\n(Copilot Studio)"]
        ProviderRouter["Provider Router\nazure_ai_foundry | copilot_studio"]
    end

    subgraph Azure["Azure Services"]
        Speech["Azure Speech Service\n(STT, TTS, Avatar relay)"]
        AgentAPI["Azure AI Foundry Project\nagents / threads / runs"]
        DirectLine["Direct Line API\nDirectLine Conversations"]
        CopilotStudio["Copilot Studio\nConversational Bots"]
        OptionalSearch["(optional) Azure Cognitive Search\nOn Your Data"]
    end

    subgraph Config["Configuration"]
        ConfigFiles["config/settings.json\n+ env vars"]
        PluginManifests["Plugin Manifests\n(plugins/*/manifest.json)"]
        AgentPlugins["Agent Plugins\n(common/agentPlugins.js)"]
    end

    Browser -- "GET static assets" --> StaticAssets
    Browser -- "GET /config" --> ConfigLoader
    ConfigLoader -- "merge defaults" --> ConfigFiles
    ConfigLoader -- "load plugins" --> AgentPlugins
    AgentPlugins -- "read manifests" --> PluginManifests

    Browser -- "POST /speech/token" --> SpeechToken
    SpeechToken -- "REST call\n(Ocp-Apim-Subscription-Key)" --> Speech
    Browser -- "WebRTC + WebSocket\n(using tokens)" --> Speech

    Browser -- "POST/GET /agent/*" --> ProxyEndpoint
    ProxyEndpoint -- "HTTP/SSE/WebSocket" --> ProviderRouter
    
    ProviderRouter -- "Azure AI Foundry" --> AuthProvider
    AuthProvider --> FoundryClient
    FoundryClient -- "threads, messages, runs\n(SSE streaming)" --> AgentAPI
    
    ProviderRouter -- "Copilot Studio" --> DirectLineClient
    DirectLineClient -- "conversations, activities\n(polling)" --> DirectLine
    DirectLine --> CopilotStudio

    FoundryClient -- "(when enabled)\nOn Your Data options" --> OptionalSearch

    Browser -- "loads & initializes" --> PluginHost
    PluginHost -- "setOverlayContent()" --> Overlay
    PluginHost -- "onAgentContent()" --> Browser
    
    ConfigLoader -- "serves sanitized config" --> Browser
    
    style Frontend fill:#e1f5ff
    style ServicesProxy fill:#fff4e1
    style AgentProxy fill:#ffe1f5
    style Azure fill:#e1ffe1
    style Config fill:#f5f5f5
```

## Key Components

### Frontend Layer
- **Browser UI**: Main application interface (chat_agent.html + chat-agent.js)
- **Plugin Host**: Manages plugin lifecycle and provides API for custom UI rendering
- **Attachment Overlay**: Renders custom HTML content (e.g., image carousels, interactive elements)

### services-proxy-server
- Serves static assets and handles client-facing endpoints
- Proxies agent requests to agent-proxy-server
- Manages Azure Speech token acquisition
- Merges configuration from files and environment variables

### agent-proxy-server
- Routes requests to appropriate provider (Azure AI Foundry or Copilot Studio)
- Handles authentication via DefaultAzureCredential for Azure services
- Manages Direct Line conversations for Copilot Studio integration
- Streams agent responses via Server-Sent Events (SSE)

### Azure Services
- **Azure Speech Service**: STT, TTS, and avatar relay for real-time communication
- **Azure AI Foundry**: Agent runtime with threads, messages, and runs API
- **Direct Line API**: Protocol for Copilot Studio bot communication
- **Azure Cognitive Search**: Optional On Your Data integration

### Configuration
- **settings.json**: Base configuration with defaults
- **Plugin Manifests**: Per-plugin configuration including connection details, branding, voice settings
- **Environment Variables**: Runtime overrides for sensitive values
