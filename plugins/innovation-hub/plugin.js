import { PluginBase } from '../../shared/pluginBase.js';

export class InnovationHubPlugin extends PluginBase {
  constructor(manifest, api) {
    super(manifest, api);
    this.lastContentHash = null;
    this._detachNavigation = null;
  }

  init() {
    this.api.log(`Plugin ${this.manifest.label} initialized.`);
  }

  destroy() {
    if (typeof this._detachNavigation === 'function') {
      this._detachNavigation();
      this._detachNavigation = null;
    }
    super.destroy();
  }

  /**
   * Called when agent produces content. Looks for image content and displays
   * it in an overlay carousel following the no-DOM-access pattern.
   */
  onAgentContent(content, context) {
    // Log the content for debugging
    if (this.api && typeof this.api.log === 'function') {
      this.api.log('[InnovationHub] Received agent content', {
        contentType: typeof content,
        isArray: Array.isArray(content),
        content: content
      });
    }
    
    const structuredContent = this._parseStructuredContent(content);
    
    // Generate hash to detect duplicate content
    const contentHash = JSON.stringify(structuredContent);
    const isDuplicateContent = contentHash === this.lastContentHash;
    
    if (isDuplicateContent) {
      return; // Skip duplicate content
    }
    
    // Update hash for new content
    this.lastContentHash = contentHash;
    
    const hasImages = structuredContent.some(item => item.type === 'image');
    
    if (hasImages) {
      const html = this._buildImageGallery(structuredContent);
      
      // Display in overlay WITHOUT a title bar
      this.setOverlayContent({
        type: 'custom_html',
        html: html,
        visible: true
      });

      // Attach navigation after overlay is rendered
      if (typeof this._detachNavigation === 'function') {
        this._detachNavigation();
        this._detachNavigation = null;
      }

      requestAnimationFrame(() => {
        const overlayContainers = document.querySelectorAll('.assistant-attachment-overlay__custom-html');
        for (const container of overlayContainers) {
          const carousel = container?.querySelector('.innovation-hub-carousel');
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

  /**
   * Parses content looking for structured data including images.
   * Supports various formats: markdown, JSON objects, or plain text with URLs.
   */
  _parseStructuredContent(content) {
    const items = [];
    
    if (!content) {
      return items;
    }

    // If content is already structured (array)
    if (Array.isArray(content)) {
      // Handle array of OpenAI message content objects
      for (const item of content) {
        if (typeof item === 'object' && item.type === 'text' && item.text) {
          // Extract text value from nested structure
          const textValue = typeof item.text === 'object' ? item.text.value : item.text;
          const parsed = this._parseStructuredContent(textValue);
          items.push(...parsed);
        } else if (typeof item === 'object' && item.type === 'image') {
          items.push(item);
        } else {
          items.push({ type: 'text', content: String(item) });
        }
      }
      return items.length > 0 ? items : [{ type: 'text', content: String(content) }];
    }

    // If content is an object with images
    if (typeof content === 'object' && !Array.isArray(content)) {
      if (content.images && Array.isArray(content.images)) {
        return content.images.map(img => ({
          type: 'image',
          url: img.url || img.src || img,
          title: img.title || '',
          description: img.description || '',
          caption: img.caption || img.alt || ''
        }));
      }
      
      // Single image object
      if (content.type === 'image' || content.url || content.src) {
        return [{
          type: 'image',
          url: content.url || content.src,
          title: content.title || '',
          description: content.description || '',
          caption: content.caption || content.alt || ''
        }];
      }
    }

    // Parse string content
    const text = String(content);
    
    // Try parsing as JSON first
    try {
      const parsed = JSON.parse(text);
      
      // Handle array of mixed content items (text + images)
      if (Array.isArray(parsed)) {
        return parsed.map(item => {
          if (item.type === 'image') {
            return {
              type: 'image',
              url: item.url,
              title: item.title || '',
              description: item.description || '',
              caption: item.caption || ''
            };
          }
          return { type: 'text', content: item.text || String(item) };
        });
      }
      
      // Handle object with images array property
      if (parsed.images && Array.isArray(parsed.images)) {
        return parsed.images.map(img => ({
          type: 'image',
          url: img.url || img.src || img,
          title: img.title || '',
          description: img.description || '',
          caption: img.caption || img.alt || ''
        }));
      }
    } catch (e) {
      // Not JSON, continue with text parsing
    }

    // Look for markdown images: ![alt](url)
    const markdownImageRegex = /!\[([^\]]*)\]\(([^)]+)\)/g;
    let match;
    
    while ((match = markdownImageRegex.exec(text)) !== null) {
      items.push({
        type: 'image',
        url: match[2],
        caption: match[1] || ''
      });
    }

    // Look for plain URLs that might be images
    const urlRegex = /(https?:\/\/[^\s]+\.(?:jpg|jpeg|png|gif|webp|svg))/gi;
    const urlMatches = text.matchAll(urlRegex);
    
    for (const urlMatch of urlMatches) {
      // Avoid duplicates from markdown parsing
      if (!items.some(item => item.url === urlMatch[1])) {
        items.push({
          type: 'image',
          url: urlMatch[1],
          caption: ''
        });
      }
    }

    return items.length > 0 ? items : [{ type: 'text', content: text }];
  }

  /**
   * Builds HTML for image carousel display in overlay.
   * Simplified carousel with button-based navigation.
   */
  _buildImageGallery(items) {
    const imageItems = items.filter(item => item.type === 'image');
    
    if (imageItems.length === 0) {
      return '';
    }

    const carouselId = `innovation-carousel-${Math.random().toString(36).slice(2, 8)}`;
    
    // Build hidden radio inputs
    const togglesHtml = imageItems.map((_, index) => {
      const isChecked = index === 0 ? 'checked' : '';
      return `<input type="radio" class="carousel__toggle" name="${carouselId}" id="${carouselId}-slide-${index}" ${isChecked} hidden>`;
    }).join('');
    
    // Build slides
    const slidesHtml = imageItems.map((item, index) => {
      const imageUrl = this._escapeHtml(item.url);
      const title = item.title ? this._escapeHtml(item.title) : '';
      const description = item.description ? this._escapeHtml(item.description) : '';
      
      return `
        <div class="carousel__slide">
          <div class="carousel__slide-inner">
            <figure class="carousel__figure">
              <img src="${imageUrl}" alt="${title || `Innovation ${index + 1}`}" class="carousel__image" loading="lazy" decoding="async" />
              ${title ? `<figcaption class="carousel__caption">${title}${description ? `<br/><span class="carousel__description">${description}</span>` : ''}</figcaption>` : ''}
            </figure>
          </div>
        </div>
      `;
    }).join('');
    
    // Build navigation controls (only if multiple images)
    const navHtml = imageItems.length > 1 ? `
      <div class="carousel__controls">
        <button type="button" class="carousel__arrow carousel__arrow--prev" data-carousel-prev aria-label="Previous innovation">
          <span aria-hidden="true">&#x2039;</span>
        </button>
        <div class="carousel__indicators">
          ${imageItems.map((_, index) => 
            `<label for="${carouselId}-slide-${index}" class="carousel__indicator" aria-label="Show innovation ${index + 1}"></label>`
          ).join('')}
        </div>
        <button type="button" class="carousel__arrow carousel__arrow--next" data-carousel-next aria-label="Next innovation">
          <span aria-hidden="true">&#x203A;</span>
        </button>
      </div>
    ` : '';
    
    return `
      <div class="innovation-hub-carousel" data-carousel-id="${carouselId}">
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

  /**
   * Attaches event handlers for carousel navigation.
   */
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

    return () => {
      prevButton.removeEventListener('click', handlePrevClick);
      nextButton.removeEventListener('click', handleNextClick);
      toggles.forEach((toggle) => toggle.removeEventListener('change', updateArrowState));
    };
  }

  /**
   * Escapes HTML to prevent XSS attacks.
   */
  _escapeHtml(text) {
    if (!text) {
      return '';
    }
    
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    
    return String(text).replace(/[&<>"']/g, (char) => map[char]);
  }

  destroy() {
    this._log('info', `Plugin ${this.manifest.label || this.id} destroyed`);
    super.destroy();
  }
}
