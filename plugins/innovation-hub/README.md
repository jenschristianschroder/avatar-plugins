# Innovation Hub Agent Plugin

This plugin provides an interface for the Innovation Hub Agent, showcasing innovative solutions and insights.

## Features

- **Interactive Image Carousel**: Display innovation showcases with navigation controls
- **JSON Response Support**: Agent responds with structured JSON containing images
- **Voice-Enabled**: Uses Andrew Multilingual Neural voice for natural conversations
- **Responsive Design**: Adapts to different screen sizes

## Configuration

### Azure AI Foundry Connection
- **Endpoint**: https://avatar-ai-foundry.services.ai.azure.com
- **Project**: avatar-agents
- **Agent ID**: asst_wJ39Cxo30HKzauQTGzDFxMMp

### Branding
- **Primary Color**: #cecece (light gray)
- **Background Image**: background.png
- **Voice**: en-US-AndrewMultilingualNeural

## Agent Response Format

The agent should respond with JSON structured as follows:

```json
[
  {
    "type": "text",
    "text": "Description or introduction text"
  },
  {
    "type": "image",
    "url": "/plugins/innovation-hub/image/innovation1.png",
    "title": "Innovation Title",
    "description": "Brief description of the innovation"
  }
]
```

## Image Assets Required

Place the following images in `/plugins/innovation-hub/image/`:

1. **background.png** - Background image for the agent interface
2. **logo.png** - Agent logo
3. **thumbnail.png** - Plugin thumbnail for selection menu
4. **[innovation images]** - Any images the agent will display in the carousel

## Testing

1. Ensure all image assets are in place
2. Configure the agent in Azure AI Foundry to respond with the JSON structure
3. Test voice interactions in English
4. Verify carousel navigation with multiple images
5. Test responsive behavior on different screen sizes

## Files

- `manifest.json` - Plugin configuration and metadata
- `plugin.js` - Plugin logic and carousel implementation
- `styles.css` - Scoped styles for the carousel overlay
- `README.md` - This file
- `INSTRUCTIONS.md` - Agent system prompt (to be created)
