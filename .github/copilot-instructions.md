
// Plugin Architecture Instructions (No DOM Access)

// Folder Structure:
// /plugins/
//   /plugin-name/
//     manifest.json
//     plugin.js
//     styles.css

// manifest.json Example:
{
  "name": "MyPlugin",
  "entry": "plugin.js",
  "style": "styles.css",
  "uiMountPoint": "#plugin-container"
}

// ============================================
// VOICE SETTINGS PRIORITY
// ============================================

// Voice settings follow a priority system where plugin manifest settings 
// take precedence over host app settings:
//
// 1. Plugin manifest speech.ttsVoice (HIGHEST PRIORITY)
// 2. Plugin manifest avatar.voice
// 3. Host app speech.ttsVoice settings (FALLBACK)
//
// Speech-to-Text (STT) locales follow the same priority:
// 1. Plugin manifest speech.sttLocales (HIGHEST PRIORITY)
// 2. Host app speech.sttLocales settings (FALLBACK)
//
// Example plugin manifest with voice settings:
{
  "id": "my-agent",
  "label": "My Agent",
  "speech": {
    "ttsVoice": "en-US-AndrewMultilingualNeural",
    "customVoiceEndpointId": "optional-endpoint-id",
    "sttLocales": "en-US,es-ES"
  },
  "avatar": {
    "character": "harry",
    "style": "business",
    "voice": "en-US-AndrewMultilingualNeural"  // Alternative location for voice
  }
}

// The plugin's speech settings will override the host app's speech settings
// when the plugin is selected. This allows each agent to have its own voice
// and speech recognition locales.

// PluginBase.js
export class PluginBase {
  constructor(manifest, api) {
    this.manifest = manifest;
    this.api = api;

    // Enforce DOM access restriction
    if (typeof document !== 'undefined') {
      Object.defineProperty(window, 'document', {
        get: () => {
          throw new Error("Access to DOM is forbidden in plugins.");
        }
      });
    }
  }

  init() {
    throw new Error("init() must be implemented by the plugin.");
  }

  destroy() {
    // Optional cleanup logic
  }
}

// ============================================
// AGENT RESPONSE FORMAT WITH COMMANDS
// ============================================

// Plugins can process structured JSON responses from agents that include both
// text for TTS and explicit commands for plugin actions. This is more reliable
// than text pattern matching.

// Response Format:
[
  {"type": "text", "text": "I'm starting the camera now"},
  {"type": "command", "command": "start_camera"}
]

// Plugin Processing Example:
onAgentContent(content, context) {
  // Try to parse structured commands first
  const commandHandled = this._tryParseCommand(content);
  if (commandHandled) {
    return;
  }
  
  // Fallback to text pattern matching for backward compatibility
  const text = this._extractTextContent(content);
  // ... pattern matching logic
}

_tryParseCommand(content) {
  try {
    let parsed = content;
    
    if (typeof content === 'string') {
      const trimmed = content.trim();
      if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
        parsed = JSON.parse(trimmed);
      } else {
        return false;
      }
    }

    if (Array.isArray(parsed)) {
      for (const item of parsed) {
        if (item && item.type === 'command' && item.command) {
          this._handleCommand(item.command);
          return true;
        }
      }
    }

    if (parsed && parsed.type === 'command' && parsed.command) {
      this._handleCommand(parsed.command);
      return true;
    }

    return false;
  } catch (e) {
    return false;
  }
}

_handleCommand(command) {
  const cmd = String(command).toLowerCase().trim();
  
  switch (cmd) {
    case 'start_camera':
      this._startCamera();
      break;
    case 'take_photo':
      this._takePhoto();
      break;
    // ... other commands
  }
}

// Agent Instructions Example:
// In your INSTRUCTIONS.md, specify the exact JSON format:
//
// **CRITICAL**: All responses MUST be valid JSON arrays with this structure:
// 
// ```json
// [
//   {"type": "text", "text": "Your spoken message here"},
//   {"type": "command", "command": "start_camera"}
// ]
// ```
//
// Available Commands:
// - `start_camera` - Opens the camera
// - `take_photo` - Triggers photo capture
//
// This approach is more reliable than pattern matching because:
// 1. No ambiguity - explicit commands vs trying to detect phrases
// 2. Language independent - commands work regardless of text language
// 3. Better for LLMs - clear structure easier to follow than exact phrases

// Example plugin.js
import { PluginBase } from '../../core/PluginBase.js';

export class MyPlugin extends PluginBase {
  init() {
    this.api.log(`Plugin ${this.manifest.name} initialized.`);
    // No DOM manipulation allowed
  }
}

// Plugin Loader
async function loadPlugin(pluginPath) {
  const manifest = await fetch(`${pluginPath}/manifest.json`).then(res => res.json());
  const { MyPlugin } = await import(`${pluginPath}/${manifest.entry}`);

  const pluginInstance = new MyPlugin(manifest, {
    log: console.log,
    config: manifest.config || {}
  });

  pluginInstance.init();
} 

// Best Practices:
// - Do not pass document, window, or DOM APIs to plugins
// - Use static analysis or runtime guards to detect forbidden DOM usage
// - Keep plugin logic focused on computation, configuration, or isolated rendering

// ============================================
// OVERLAY INTEGRATION FOR PLUGINS
// ============================================

// Plugins display interactive content in an attachment overlay using the 
// setOverlayContent method provided by the plugin API. After calling this method,
// the host app renders the HTML into the DOM, and the plugin can then access it
// to attach event handlers for interactive features.

// IMPORTANT: Two-Phase Rendering Pattern
// 1. Build HTML as strings and call setOverlayContent()
// 2. Use requestAnimationFrame to query the rendered DOM and attach interactivity

// Overlay Content Structure:
this.setOverlayContent({
  type: 'custom_html',     // Type of content (always 'custom_html' for HTML content)
  html: htmlString,        // Your HTML content as a string
  visible: true            // Controls overlay visibility
  // Note: Do NOT include 'title' property unless you want a title bar
});

// ============================================
// INTERACTIVE CAROUSEL IMPLEMENTATION
// ============================================

// Example: Building an Interactive Image Carousel
// This demonstrates best practices from the retail-banking and trade plugins

// Step 1: Build HTML Structure with Hidden Radio Inputs
_buildImageCarousel(items) {
  const imageItems = items.filter(item => item.type === 'image');
  
  if (imageItems.length === 0) {
    return '';
  }

  const carouselId = `carousel-${Math.random().toString(36).slice(2, 8)}`;
  
  // Hidden radio inputs control carousel state via CSS
  const togglesHtml = imageItems.map((_, index) => {
    const isChecked = index === 0 ? 'checked' : '';
    return `<input type="radio" class="carousel__toggle" name="${carouselId}" id="${carouselId}-slide-${index}" ${isChecked} hidden>`;
  }).join('');
  
  // Build slides with proper structure
  const slidesHtml = imageItems.map((item, index) => {
    const imageUrl = this._escapeHtml(item.url);
    const title = item.title ? this._escapeHtml(item.title) : '';
    const description = item.description ? this._escapeHtml(item.description) : '';
    
    return `
      <div class="carousel__slide">
        <div class="carousel__slide-inner">
          <figure class="carousel__figure">
            <img src="${imageUrl}" alt="${title || `Image ${index + 1}`}" class="carousel__image" loading="lazy" decoding="async" />
            ${title ? `<figcaption class="carousel__caption">${title}${description ? `<br/><span>${description}</span>` : ''}</figcaption>` : ''}
          </figure>
        </div>
      </div>
    `;
  }).join('');
  
  // Navigation controls (arrows + indicator dots)
  const navHtml = imageItems.length > 1 ? `
    <div class="carousel__controls">
      <button type="button" class="carousel__arrow carousel__arrow--prev" data-carousel-prev aria-label="Previous">
        <span aria-hidden="true">&#x2039;</span>
      </button>
      <div class="carousel__indicators">
        ${imageItems.map((_, index) => 
          `<label for="${carouselId}-slide-${index}" class="carousel__indicator" aria-label="Show image ${index + 1}"></label>`
        ).join('')}
      </div>
      <button type="button" class="carousel__arrow carousel__arrow--next" data-carousel-next aria-label="Next">
        <span aria-hidden="true">&#x203A;</span>
      </button>
    </div>
  ` : '';
  
  return `
    <div class="my-plugin-carousel" data-carousel-id="${carouselId}">
      ${togglesHtml}
      <div class="carousel__viewport">
        <div class="carousel__track">
          ${slidesHtml}
        </div>
      </div>
      ${navHtml}
    </div>
  `;
}

// Step 2: Attach Interactive Event Handlers After Rendering
// Store detach functions for cleanup
constructor(manifest, api) {
  super(manifest, api);
  this.lastContentHash = null;
  this._detachNavigation = null;  // Store cleanup function
}

onAgentContent(content, context) {
  const structuredContent = this._parseStructuredContent(content);
  
  // Generate hash to detect duplicate content
  const contentHash = JSON.stringify(structuredContent);
  const isDuplicateContent = contentHash === this.lastContentHash;
  
  if (isDuplicateContent) {
    return;  // Skip duplicate content
  }
  
  // Update hash for new content
  this.lastContentHash = contentHash;
  
  const hasImages = structuredContent.some(item => item.type === 'image');
  
  if (hasImages) {
    const html = this._buildImageCarousel(structuredContent);
    
    // Display in overlay WITHOUT a title bar
    this.setOverlayContent({
      type: 'custom_html',
      html: html,
      visible: true
    });

    // CRITICAL: Attach navigation AFTER overlay is rendered
    // Clean up previous handlers first
    if (typeof this._detachNavigation === 'function') {
      this._detachNavigation();
      this._detachNavigation = null;
    }

    // Use requestAnimationFrame to ensure DOM is ready
    requestAnimationFrame(() => {
      const overlayContainers = document.querySelectorAll('.assistant-attachment-overlay__custom-html');
      for (const container of overlayContainers) {
        const carousel = container?.querySelector('.my-plugin-carousel');
        if (carousel) {
          const detachNavigation = this._attachCarouselNavigation(carousel);
          if (typeof detachNavigation === 'function') {
            this._detachNavigation = detachNavigation;
          }
          break;
        }
      }
    });
  } else {
    // Clear overlay when new content arrives without images
    this.setOverlayContent({
      type: 'custom_html',
      html: '',
      visible: false
    });
  }
}

// Step 3: Implement Navigation Event Handlers
_attachCarouselNavigation(carouselElement) {
  if (!carouselElement || typeof window === 'undefined') {
    return null;
  }

  const prevButton = carouselElement.querySelector('[data-carousel-prev]');
  const nextButton = carouselElement.querySelector('[data-carousel-next]');
  if (!prevButton || !nextButton) {
    return null;
  }

  const toggles = Array.from(carouselElement.querySelectorAll('.carousel__toggle'));
  if (toggles.length <= 1) {
    prevButton.disabled = true;
    nextButton.disabled = true;
    return () => {};
  }

  const getActiveIndex = () => {
    const index = toggles.findIndex((toggle) => toggle.checked);
    return index >= 0 ? index : 0;
  };

  const activateIndex = (index) => {
    const clamped = Math.max(0, Math.min(index, toggles.length - 1));
    const target = toggles[clamped];
    if (!target || target.checked) {
      return;
    }
    target.checked = true;
    target.dispatchEvent(new Event('change', { bubbles: true }));
  };

  const updateArrowState = () => {
    const index = getActiveIndex();
    const isAtStart = index <= 0;
    const isAtEnd = index >= toggles.length - 1;
    prevButton.disabled = isAtStart;
    prevButton.setAttribute('aria-disabled', isAtStart ? 'true' : 'false');
    nextButton.disabled = isAtEnd;
    nextButton.setAttribute('aria-disabled', isAtEnd ? 'true' : 'false');
  };

  const handlePrevClick = (event) => {
    event.preventDefault();
    const index = getActiveIndex();
    if (index > 0) {
      activateIndex(index - 1);
    }
    updateArrowState();
  };

  const handleNextClick = (event) => {
    event.preventDefault();
    const index = getActiveIndex();
    if (index < toggles.length - 1) {
      activateIndex(index + 1);
    }
    updateArrowState();
  };

  prevButton.addEventListener('click', handlePrevClick);
  nextButton.addEventListener('click', handleNextClick);
  toggles.forEach((toggle) => toggle.addEventListener('change', updateArrowState));

  updateArrowState();

  // Return cleanup function
  return () => {
    prevButton.removeEventListener('click', handlePrevClick);
    nextButton.removeEventListener('click', handleNextClick);
    toggles.forEach((toggle) => toggle.removeEventListener('change', updateArrowState));
  };
}

// Step 4: Clean Up in destroy()
destroy() {
  if (typeof this._detachNavigation === 'function') {
    this._detachNavigation();
    this._detachNavigation = null;
  }
  super.destroy();
}

// ============================================
// CSS STYLING BEST PRACTICES
// ============================================

// All plugin styles MUST be scoped to prevent conflicts
// Use: .assistant-attachment-overlay[data-plugin-id="your-plugin-id"] prefix

// Carousel Base Styles
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .my-plugin-carousel {
  position: relative;
  width: 100%;
  height: 100%;
  margin: 0;
  border-radius: 16px;
  background: #ffffff;
  box-shadow: 0 14px 28px rgba(0, 1, 159, 0.18);
  overflow: hidden;
}

// Hide radio button controls (state management)
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle {
  position: absolute;
  opacity: 0;
  pointer-events: none;
}

// Viewport and Track (horizontal scrolling container)
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__viewport {
  overflow: hidden;
  -ms-overflow-style: none;
  scrollbar-width: none;
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__track {
  display: flex;
  width: 100%;
  height: 100%;
  transition: transform 0.35s ease;
}

// Individual Slides
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__slide {
  flex: 0 0 100%;
  max-width: 100%;
  height: 100%;
  overflow-y: auto;
  background: #f1f5f9;
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__slide-inner {
  min-height: 100%;
  min-width: 100%;
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 0;
  box-sizing: border-box;
}

// Images - CRITICAL: Use object-fit: contain
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__image {
  display: block;
  width: 100%;
  height: 100%;
  border-radius: 0;
  object-fit: contain;  // IMPORTANT: Prevents cropping
}

// Navigation Controls
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__controls {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 16px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 0 20px;
  pointer-events: none;  // Container doesn't intercept clicks
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__arrow {
  width: 44px;
  height: 44px;
  border-radius: 50%;
  border: none;
  background: rgba(0, 1, 159, 0.75);
  color: #f8fafc;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 2rem;
  line-height: 1;
  cursor: pointer;
  transition: background 0.2s ease, transform 0.2s ease, opacity 0.2s ease;
  pointer-events: auto;  // Buttons ARE clickable
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__arrow:hover {
  background: rgba(0, 1, 159, 0.9);
  transform: translateY(-2px);
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__arrow:disabled {
  opacity: 0.35;
  cursor: default;
  transform: none;
}

// Indicator Dots
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__indicators {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 10px 14px;
  border-radius: 999px;
  background: rgba(0, 1, 159, 0.12);
  pointer-events: auto;
  flex: 1 1 auto;
  max-width: 200px;
}

.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__indicator {
  width: 10px;
  height: 10px;
  border-radius: 50%;
  background: rgba(0, 1, 159, 0.3);
  cursor: pointer;
  transition: transform 0.2s ease, background 0.2s ease;
}

// CSS State Management - Radio buttons control which slide is visible
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(1):checked ~ .carousel__viewport .carousel__track { 
  transform: translateX(0); 
}
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(2):checked ~ .carousel__viewport .carousel__track { 
  transform: translateX(-100%); 
}
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(3):checked ~ .carousel__viewport .carousel__track { 
  transform: translateX(-200%); 
}

// Active indicator highlighting
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(1):checked ~ .carousel__controls .carousel__indicators label:nth-of-type(1),
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(2):checked ~ .carousel__controls .carousel__indicators label:nth-of-type(2),
.assistant-attachment-overlay[data-plugin-id="your-plugin-id"] .carousel__toggle:nth-of-type(3):checked ~ .carousel__controls .carousel__indicators label:nth-of-type(3) {
  background: #00019f;
  transform: scale(1.15);
}

// ============================================
// KEY GUIDELINES FOR PLUGIN DEVELOPMENT
// ============================================

// 1. HTML Generation
//    - Build HTML as strings, not DOM elements
//    - Always escape user-provided content to prevent XSS attacks
//    - Use semantic HTML5 elements (figure, figcaption, etc.)
//    - Include ARIA labels for accessibility

// 2. DOM Access Pattern
//    - Call setOverlayContent() first to render HTML
//    - Use requestAnimationFrame() before querying DOM
//    - Query for elements using document.querySelectorAll()
//    - Attach event handlers after elements are found
//    - Store cleanup functions for event listener removal

// 3. CSS Styling
//    - ALWAYS scope styles: .assistant-attachment-overlay[data-plugin-id="your-plugin-id"]
//    - Use object-fit: contain for images (prevents cropping)
//    - Set pointer-events: none on containers, auto on interactive elements
//    - Use CSS transforms for animations (better performance)
//    - Support responsive design with @media queries

// 4. State Management
//    - Use hidden radio inputs for CSS-only state control
//    - Track content hashes to detect duplicate updates
//    - Store detach functions for cleanup in destroy()
//    - Manage loading/error states explicitly

// 5. Interactive Features
//    - Implement proper keyboard navigation (arrows, tab, enter)
//    - Disable/enable buttons based on state
//    - Update ARIA attributes dynamically (aria-disabled, aria-label)
//    - Provide visual feedback for interactions (hover, focus, active)

// 6. Performance
//    - Use CSS transitions over JavaScript animations
//    - Lazy load images with loading="lazy" attribute
//    - Debounce rapid state changes
//    - Clean up event listeners in destroy()

// 7. Error Handling
//    - Validate content structure before rendering
//    - Handle missing or malformed data gracefully
//    - Log errors without breaking the UI
//    - Provide fallback content when data is unavailable

// 8. Testing Checklist
//    - Test with 1, 2, and 5+ items
//    - Verify navigation at boundaries (first/last)
//    - Check duplicate content handling
//    - Test with different image sizes and aspect ratios
//    - Verify cleanup when switching agents
//    - Test keyboard navigation
//    - Validate on mobile viewports

// ============================================
// COMPLETE PLUGIN EXAMPLE
// ============================================

// See /plugins/retail-banking/ and /plugins/trade/ for production implementations
// Key files:
//   - manifest.json: Plugin configuration
//   - plugin.js: Main plugin logic with carousel
//   - styles.css: Scoped styles for overlay content
//   - INSTRUCTIONS.md: Agent system prompt (if using AI agent)

