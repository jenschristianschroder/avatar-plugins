# Plugin Development Guide

This guide explains how to create plugins for the Azure TTS Talking Avatar application.

## Overview

Plugins extend the avatar application with custom functionality, including:
- Custom AI agents with specific personas and capabilities
- Interactive UI components displayed in the attachment overlay
- Integration with external services and APIs
- Custom voice and speech recognition settings

## Plugin Structure

```
/plugins/
  /your-plugin-name/
    manifest.json          # Plugin configuration (required)
    plugin.js             # Plugin implementation (required)
    styles.css            # Scoped styles for overlay content (optional)
    INSTRUCTIONS.md       # AI agent system prompt (optional, for AI agents)
    README.md             # Plugin documentation (optional)
    image/                # Plugin assets (optional)
```

## Creating a Plugin

### Step 1: Create Plugin Folder

Create a new folder under `/plugins/` with a descriptive name (e.g., `my-agent`).

### Step 2: Create manifest.json

The manifest defines your plugin's configuration:

```json
{
  "id": "my-agent",
  "label": "My Agent",
  "description": "Description of what your agent does",
  "entry": "plugin.js",
  "style": "styles.css",
  "speech": {
    "ttsVoice": "en-US-AndrewMultilingualNeural",
    "customVoiceEndpointId": "",
    "sttLocales": "en-US,es-ES"
  },
  "avatar": {
    "character": "lisa",
    "style": "casual",
    "voice": "en-US-AndrewMultilingualNeural"
  },
  "config": {
    "customSetting": "value"
  }
}
```

**Key Properties:**
- `id` (required): Unique identifier for your plugin
- `label` (required): Display name shown in the UI
- `entry` (required): JavaScript file containing your plugin class
- `speech`: Override default voice and STT settings
- `avatar`: Avatar character and style configuration
- `config`: Custom configuration accessible via `this.config`

### Step 3: Create plugin.js

Your plugin must extend `PluginBase` from `shared/pluginBase.js`:

```javascript
import { PluginBase } from '../../shared/pluginBase.js';

export class MyPlugin extends PluginBase {
  constructor(manifest, api) {
    super(manifest, api);
    // Initialize state
    this.lastContentHash = null;
    this._detachNavigation = null;
  }

  init() {
    this.api.log(`${this.manifest.label} initialized`);
    // Setup logic here
  }

  destroy() {
    // Cleanup event handlers
    if (typeof this._detachNavigation === 'function') {
      this._detachNavigation();
      this._detachNavigation = null;
    }
    super.destroy();
  }

  onAgentContent(content, context) {
    // React to agent responses
    // See "Handling Agent Content" section below
  }

  _escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
}

export default MyPlugin;
```

**Note:** The default export is required for the plugin loader.

## Displaying Interactive Content

Plugins can display content in an attachment overlay using the two-phase rendering pattern.

### Phase 1: Build HTML and Render

```javascript
onAgentContent(content, context) {
  const structuredContent = this._parseStructuredContent(content);
  
  // Detect duplicate content
  const contentHash = JSON.stringify(structuredContent);
  if (contentHash === this.lastContentHash) {
    return;
  }
  this.lastContentHash = contentHash;
  
  const hasImages = structuredContent.some(item => item.type === 'image');
  
  if (hasImages) {
    const html = this._buildCarousel(structuredContent);
    
    this.setOverlayContent({
      type: 'custom_html',
      html: html,
      visible: true
    });
    
    // Phase 2: Attach interactivity (see next section)
    this._attachInteractivity();
  } else {
    this.setOverlayContent({
      type: 'custom_html',
      html: '',
      visible: false
    });
  }
}
```

### Phase 2: Attach Event Handlers

After `setOverlayContent()`, the host app renders the HTML. Use `requestAnimationFrame` to query the DOM and attach handlers:

```javascript
_attachInteractivity() {
  // Clean up previous handlers
  if (typeof this._detachNavigation === 'function') {
    this._detachNavigation();
    this._detachNavigation = null;
  }

  requestAnimationFrame(() => {
    const overlayContainers = document.querySelectorAll('.assistant-attachment-overlay__custom-html');
    for (const container of overlayContainers) {
      const carousel = container?.querySelector('.my-plugin-carousel');
      if (carousel) {
        const detach = this._attachCarouselNavigation(carousel);
        if (typeof detach === 'function') {
          this._detachNavigation = detach;
        }
        break;
      }
    }
  });
}

_attachCarouselNavigation(element) {
  const prevButton = element.querySelector('[data-carousel-prev]');
  const nextButton = element.querySelector('[data-carousel-next]');
  
  if (!prevButton || !nextButton) {
    return null;
  }

  const handlePrev = (e) => {
    e.preventDefault();
    // Navigation logic
  };

  const handleNext = (e) => {
    e.preventDefault();
    // Navigation logic
  };

  prevButton.addEventListener('click', handlePrev);
  nextButton.addEventListener('click', handleNext);

  // Return cleanup function
  return () => {
    prevButton.removeEventListener('click', handlePrev);
    nextButton.removeEventListener('click', handleNext);
  };
}
```

## CSS Styling Best Practices

All styles MUST be scoped to prevent conflicts with other plugins and the host app.

### Scoping Pattern

```css
/* ALWAYS prefix with plugin-specific scope */
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .your-class {
  /* styles here */
}
```

### Critical Image Styling

```css
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__image {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: contain;  /* CRITICAL: Prevents cropping */
  border-radius: 0;
}
```

### Interactive Elements

```css
/* Container blocks click events */
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__controls {
  pointer-events: none;
}

/* Interactive elements receive clicks */
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__arrow {
  pointer-events: auto;
  cursor: pointer;
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__arrow:disabled {
  opacity: 0.35;
  cursor: default;
}
```

### CSS State Management

Use hidden radio inputs for CSS-only state control:

```css
/* Hide radio controls */
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle {
  position: absolute;
  opacity: 0;
  pointer-events: none;
}

/* Transform based on checked state */
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(1):checked ~ .carousel__viewport .carousel__track {
  transform: translateX(0);
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(2):checked ~ .carousel__viewport .carousel__track {
  transform: translateX(-100%);
}
```

## Voice Settings

Plugins can override the host app's voice and speech recognition settings:

### Priority Order

1. **Plugin manifest `speech.ttsVoice`** (highest priority)
2. **Plugin manifest `avatar.voice`**
3. **Host app settings** (fallback)

### Example Configuration

```json
{
  "id": "multilingual-agent",
  "speech": {
    "ttsVoice": "en-US-AndrewMultilingualNeural",
    "customVoiceEndpointId": "optional-custom-voice-id",
    "sttLocales": "en-US,es-ES,fr-FR"
  }
}
```

## Handling Agent Content

The `onAgentContent` lifecycle method receives agent responses:

```javascript
onAgentContent(content, context) {
  // content: The actual content from the agent
  // context: {
  //   pluginId: string,
  //   manifest: object,
  //   rawPayload: object  // Full response from agent
  // }
  
  // Parse and structure the content
  const items = this._parseContent(content);
  
  // Render in overlay
  const html = this._buildUI(items);
  this.setOverlayContent({
    type: 'custom_html',
    html: html,
    visible: true
  });
}
```

## Sending User Input

Plugins can send text input back to the agent:

```javascript
handleUserAction(message) {
  const success = this.sendTextInput(message, {
    action: 'user_clicked',
    timestamp: Date.now()
  });
  
  if (!success) {
    this.api.log('Failed to send user input');
  }
}
```

## Error Handling

Always validate content and handle errors gracefully:

```javascript
_parseContent(content) {
  try {
    if (Array.isArray(content)) {
      return content;
    }
    
    if (typeof content === 'string') {
      const parsed = JSON.parse(content);
      return Array.isArray(parsed) ? parsed : [parsed];
    }
    
    return [content];
  } catch (error) {
    this.api.log('Content parsing failed', error);
    return [];
  }
}

_buildCarousel(items) {
  if (!Array.isArray(items) || items.length === 0) {
    return '<div class="empty-state">No content available</div>';
  }
  
  // Build HTML safely
}
```

## Testing Checklist

- [ ] Test with 1, 2, and 5+ items
- [ ] Verify navigation at boundaries (first/last slide)
- [ ] Check duplicate content handling
- [ ] Test with different image sizes and aspect ratios
- [ ] Verify cleanup when switching to another agent
- [ ] Test keyboard navigation (if applicable)
- [ ] Validate responsive behavior on mobile
- [ ] Test error states (missing data, network errors)
- [ ] Verify ARIA labels for accessibility
- [ ] Check console for errors and warnings

## Example Plugins

### Production Examples

- **retail-banking**: Danish banking agent with image carousel
- **trade**: Financial trading agent with interactive charts
- **tic-tac-toe**: Interactive game board with move handling

### Minimal Example

See the complete implementation in `/plugins/support/` for a simple agent without custom UI.

## Available APIs

### PluginBase Properties

- `this.manifest` - Plugin configuration (read-only)
- `this.id` - Plugin ID from manifest
- `this.api` - Host API methods
- `this.config` - Custom configuration from manifest
- `this.isDestroyed` - Destruction state flag

### PluginBase Methods

- `setOverlayContent(descriptor)` - Display content in overlay
- `clearOverlayContent()` - Hide overlay content
- `sendTextInput(text, metadata)` - Send user input to agent
- `onAgentContent(content, context)` - Override to handle agent responses
- `init()` - Override for initialization logic (required)
- `destroy()` - Override for cleanup logic

### Host API Methods (via `this.api`)

- `log(message, data)` - Log messages
- `onAgentContent(handler, context)` - Subscribe to agent responses
- `sendTextInput(text, context)` - Send text to agent
- `setOverlayContent(content, payload)` - Update overlay (called by plugin)

## Performance Optimization

1. **Use CSS Transforms**: Prefer `transform` over `left`/`top` for animations
2. **Lazy Load Images**: Add `loading="lazy"` attribute to images
3. **Debounce Events**: Throttle rapid user interactions
4. **Clean Up**: Always remove event listeners in `destroy()`
5. **Cache DOM Queries**: Store element references instead of re-querying
6. **Minimize Reflows**: Batch DOM updates when possible

## Security Considerations

1. **Always Escape User Content**: Use `_escapeHtml()` for all user-provided strings
2. **Validate Input**: Check content structure before processing
3. **Scope Styles**: Always use plugin-specific CSS selectors
4. **No Inline Scripts**: Never use `onclick` or similar inline handlers
5. **Sanitize URLs**: Validate image URLs before rendering

## Troubleshooting

### Overlay Not Showing

- Check that `visible: true` is set in `setOverlayContent()`
- Verify HTML string is not empty
- Check browser console for errors
- Ensure CSS is properly scoped

### Navigation Not Working

- Verify `requestAnimationFrame` is used before querying DOM
- Check that event handlers are attached correctly
- Ensure cleanup function is stored and called
- Verify element selectors match rendered HTML

### Styles Not Applied

- Confirm CSS scope: `.assistant-attachment-overlay[data-plugin-id="your-id"]`
- Check that `styles.css` is referenced in `manifest.json`
- Verify CSS file is in the plugin folder
- Inspect element in browser DevTools to check applied styles

### Duplicate Content

- Implement content hash tracking
- Compare new hash with `lastContentHash` before rendering
- Update hash AFTER comparison, not before

## Additional Resources

- See `.github/copilot-instructions.md` for detailed implementation patterns
- Review `shared/pluginBase.js` for base class implementation
- Study `/plugins/retail-banking/` and `/plugins/trade/` for complete examples
- Check Azure TTS documentation: https://learn.microsoft.com/azure/ai-services/speech-service/

## Support

For questions or issues:
1. Check existing plugin implementations for examples
2. Review the copilot instructions for patterns
3. Test in browser DevTools to inspect rendered output
4. Check console for error messages and warnings
