# Avatar Interaction Sequence

This diagram shows the detailed sequence of interactions during a typical user conversation with the avatar, including both Azure AI Foundry and Copilot Studio provider flows.

## Azure AI Foundry Flow

```mermaid
sequenceDiagram
    participant User
    participant Browser as Browser UI
    participant Plugin as Plugin Host
    participant SpeechSTT as Azure Speech STT
    participant ServicesProxy as services-proxy-server
    participant AgentProxy as agent-proxy-server
    participant Credential as DefaultAzureCredential
    participant AgentService as Azure AI Foundry Agent
    participant SpeechTTS as Azure Speech TTS/Avatar

    User->>Browser: Initialize session
    Browser->>ServicesProxy: GET /config
    ServicesProxy-->>Browser: Runtime config (sanitized)
    Browser->>ServicesProxy: POST /speech/token
    ServicesProxy->>SpeechTTS: Request relay + STS tokens
    SpeechTTS-->>ServicesProxy: Tokens (expires in 540s)
    ServicesProxy-->>Browser: Speech tokens
    Browser->>SpeechTTS: Establish WebRTC session
    
    Browser->>Plugin: Initialize active plugin
    Plugin-->>Browser: Plugin ready
    
    User->>Browser: Speak or type input
    opt Speech capture
        Browser->>SpeechSTT: Stream audio via WebRTC
        SpeechSTT-->>Browser: Recognized text payload
    end
    opt Text entry
        Browser-->>Browser: Capture typed message
    end
    
    Browser->>ServicesProxy: POST /agent/thread
    ServicesProxy->>AgentProxy: Forward thread creation
    AgentProxy->>Credential: Acquire Azure token
    Credential-->>AgentProxy: Access token
    AgentProxy->>AgentService: Create thread
    AgentService-->>AgentProxy: Thread ID
    AgentProxy-->>Browser: Thread ID
    
    Browser->>ServicesProxy: POST /agent/message
    ServicesProxy->>AgentProxy: Forward message (thread ID + content)
    AgentProxy->>AgentService: Create message in thread
    AgentService-->>AgentProxy: Message created
    AgentProxy-->>Browser: Message confirmed
    
    Browser->>ServicesProxy: GET /agent/run-stream
    ServicesProxy->>AgentProxy: Forward SSE request
    AgentProxy->>Credential: Refresh token if needed
    Credential-->>AgentProxy: Valid token
    AgentProxy->>AgentService: Create run & stream response
    
    loop Streaming response
        AgentService-->>AgentProxy: SSE: thread.message.delta
        AgentProxy-->>ServicesProxy: Forward SSE event
        ServicesProxy-->>Browser: SSE event stream
        Browser->>Plugin: onAgentContent(delta)
        Plugin-->>Browser: Process content (extract images, etc.)
    end
    
    AgentService-->>AgentProxy: SSE: thread.message.completed
    AgentProxy-->>Browser: Message completion payload
    Browser->>Plugin: onAgentContent(completed)
    Plugin->>Browser: setOverlayContent(HTML)
    Browser-->>Browser: Render overlay with custom content
    
    AgentService-->>AgentProxy: SSE: thread.run.completed
    AgentProxy-->>Browser: Run completed
    
    Browser->>SpeechTTS: Request avatar speech synthesis
    SpeechTTS-->>Browser: Voiced audio + viseme events
    Browser-->>User: Render avatar animation and play audio
```

## Copilot Studio (Direct Line) Flow

```mermaid
sequenceDiagram
    participant User
    participant Browser as Browser UI
    participant Plugin as Plugin Host
    participant ServicesProxy as services-proxy-server
    participant AgentProxy as agent-proxy-server
    participant DirectLine as Direct Line API
    participant CopilotBot as Copilot Studio Bot
    participant SpeechTTS as Azure Speech TTS/Avatar

    User->>Browser: Initialize session
    Browser->>ServicesProxy: GET /config
    ServicesProxy-->>Browser: Runtime config (with Direct Line settings)
    
    Browser->>ServicesProxy: POST /agent/thread
    ServicesProxy->>AgentProxy: Create conversation
    AgentProxy->>DirectLine: POST /v3/directline/tokens/generate
    Note over AgentProxy,DirectLine: Authorization: Bearer {directLineSecret}<br/>Payload: {user: {id}, bot: {id}, scope, region}
    DirectLine-->>AgentProxy: {conversationId, token, streamUrl, expires_in}
    AgentProxy->>DirectLine: POST /v3/directline/conversations
    Note over AgentProxy,DirectLine: Authorization: Bearer {token}
    DirectLine-->>AgentProxy: Conversation started
    AgentProxy-->>Browser: {id: conversationId}
    
    Browser->>Plugin: Initialize plugin for Copilot Studio
    Plugin-->>Browser: Plugin ready
    
    User->>Browser: Send message
    Browser->>ServicesProxy: POST /agent/message
    ServicesProxy->>AgentProxy: Forward message
    AgentProxy->>DirectLine: POST /v3/directline/conversations/{id}/activities
    Note over AgentProxy,DirectLine: Authorization: Bearer {token}<br/>Payload: {type: "message", text, attachments}
    DirectLine->>CopilotBot: Deliver message activity
    DirectLine-->>AgentProxy: Activity ID
    AgentProxy-->>Browser: Message sent
    
    Browser->>ServicesProxy: GET /agent/run-stream
    ServicesProxy->>AgentProxy: Start streaming
    
    loop Polling for response (every 1000ms)
        AgentProxy->>DirectLine: GET /v3/directline/conversations/{id}/activities?watermark={w}
        Note over AgentProxy,DirectLine: Authorization: Bearer {token}
        DirectLine->>CopilotBot: Check for new activities
        CopilotBot-->>DirectLine: Bot activities (if any)
        DirectLine-->>AgentProxy: {activities: [...], watermark}
        
        opt New bot activity received
            AgentProxy-->>Browser: SSE: thread.message.delta
            Browser->>Plugin: onAgentContent(delta text only)
            Note over Browser,Plugin: Filters structured content:<br/>only speakable text sent for TTS
            
            AgentProxy-->>Browser: SSE: thread.message.completed
            Browser->>Plugin: onAgentContent(full content)
            Note over Browser,Plugin: Includes all structured items:<br/>text, images, attachments
            Plugin->>Browser: setOverlayContent(carousel HTML)
            Browser-->>Browser: Render interactive overlay
            
            AgentProxy-->>Browser: SSE: thread.run.completed
            Note over AgentProxy,Browser: Exit polling loop
        end
        
        opt Token expiring
            AgentProxy->>DirectLine: POST /v3/directline/tokens/refresh
            DirectLine-->>AgentProxy: New token
        end
    end
    
    Browser->>SpeechTTS: Request TTS for filtered text
    SpeechTTS-->>Browser: Audio + viseme data
    Browser-->>User: Animate avatar with speech
```

## Key Differences: Azure AI Foundry vs Copilot Studio

### Azure AI Foundry
- **Authentication**: Uses DefaultAzureCredential (Managed Identity)
- **Protocol**: Native SSE streaming from Azure AI Foundry agents API
- **Message Flow**: Thread → Message → Run (async streaming)
- **Response Format**: Structured delta events with content arrays

### Copilot Studio (Direct Line)
- **Authentication**: Uses Direct Line secret (from manifest or env vars)
- **Protocol**: Direct Line REST API with polling
- **Message Flow**: Conversation → Activity → Poll for responses
- **Response Format**: Activities with text and attachments, transformed to delta events
- **Content Handling**: 
  - Delta events receive only speakable text (type: "text") for TTS
  - Completion events receive full structured content including images
  - Supports JSON-based structured content arrays

## Plugin Content Processing

Both flows support plugin-based custom rendering:

1. **Content Reception**: Plugin receives `onAgentContent(content, context)` callback
2. **Content Parsing**: Plugin extracts structured items (text, images, attachments)
3. **HTML Generation**: Plugin builds custom HTML (e.g., carousels) as strings
4. **Overlay Rendering**: Plugin calls `setOverlayContent({type: 'custom_html', html, visible})`
5. **DOM Attachment**: After rendering, plugin uses `requestAnimationFrame()` to query DOM and attach event handlers
6. **Interactive Features**: Navigation buttons, indicators, keyboard controls

## Voice Settings Priority

Voice settings follow a cascading priority system:

1. **Plugin manifest** `speech.ttsVoice` (HIGHEST)
2. **Plugin manifest** `avatar.voice`
3. **Host app** speech settings (FALLBACK)

STT locales follow the same pattern via `speech.sttLocales`.
