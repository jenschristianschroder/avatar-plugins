// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.

// Global objects
var speechRecognizer
var avatarSynthesizer
var peerConnection
var peerConnectionDataChannel
var messages = []
var messageInitiated = false
var dataSources = []
var sentenceLevelPunctuations = [ '.', '?', '!', ':', ';', '。', '？', '！', '：', '；' ]
var enableDisplayTextAlignmentWithSpeech = true
var enableQuickReply = false
var defaultQuickReplies = [ 'Let me take a look.', 'Let me check.', 'One moment, please.' ]
var quickReplies = defaultQuickReplies.slice()
var byodDocRegex = new RegExp(/\[doc(\d+)\]/g)
var isSpeaking = false
var isReconnecting = false
var speakingText = ""
var spokenTextQueue = []
var repeatSpeakingSentenceAfterReconnection = true
var sessionActive = false
var userClosedSession = false
var lastInteractionTime = new Date()
var lastSpeakTime
var imgUrl = ""
var currentThreadId = null
var spokenAssistantMessageIds = new Set()
var currentRunEventSource = null
var assistantMessageBuffers = new Map()
var lastAssistantMessageKey = null
var runInProgress = false
var runtimeConfig = null
var runtimeConfigDefaults = null
var servicesProxyBaseUrl = ''
var startSessionButtonOriginalLabel = null
var pendingAvatarConnection = null
var sessionShutdownPromise = null
var activeBrandingConfig = null
var runtimeConfigAssetBaseUrl = ''

const AGENT_SELECTOR_STORAGE_KEY = 'avatar.selectedAgent'

var agentSelectorState = {
    options: [],
    selectedKey: null,
    selectedPluginId: null,
    panelOpen: false,
    initialized: false,
    locked: false,
    overrideQuery: '',
    tagFilters: [],
    baseBranding: null,
    pluginMap: {}
}

var agentSelectorElements = {
    container: null,
    carouselTrack: null,
    prevBtn: null,
    nextBtn: null,
    indicators: null
}

var carouselState = {
    currentIndex: 0,
    slideCount: 0,
    isTransitioning: false
}

const attachmentOverlayState = {
    attachments: [],
    pluginDescriptor: null
}

const pluginHostState = {
    activePluginId: null,
    activeInstance: null,
    pendingActivationToken: null,
    agentContentListeners: new Set(),
    moduleCache: new Map(),
    lastManifest: null,
    activeStylesheetElement: null,
    activeStylesheetUrl: null
}

let pluginOverlayEventsBound = false

const pluginDebugState = {
    history: [],
    limit: 200,
    lastIndex: 0,
    flushPending: false
}

function recordPluginDebugEntry(args) {
    if (!pluginDebugState.history) {
        pluginDebugState.history = []
    }
    const timestamp = Date.now()
    pluginDebugState.history.push({ timestamp, args })
    if (pluginDebugState.history.length > pluginDebugState.limit) {
        pluginDebugState.history.shift()
        if (pluginDebugState.lastIndex > 0) {
            pluginDebugState.lastIndex -= 1
        }
    }
    if (pluginDebugState.flushPending && isPluginDebugEnabled()) {
        flushPluginDebugHistory()
    }
}

function flushPluginDebugHistory(includeAll = false) {
    const entries = pluginDebugState.history || []
    if (!entries.length) {
        pluginDebugState.lastIndex = 0
        pluginDebugState.flushPending = false
        return
    }

    if (typeof console === 'undefined' || typeof console.log !== 'function') {
        pluginDebugState.flushPending = true
        return
    }

    const startIndex = includeAll ? 0 : pluginDebugState.lastIndex
    if (startIndex >= entries.length) {
        pluginDebugState.flushPending = false
        return
    }

    const canGroup = typeof console.group === 'function' && typeof console.groupEnd === 'function'
    if (canGroup) {
        console.group('[PluginHost] Replaying buffered plugin logs')
    } else {
        console.log('[PluginHost] Replaying buffered plugin logs')
    }

    for (let index = startIndex; index < entries.length; index += 1) {
        const entry = entries[index]
        const timestamp = entry.timestamp
        const label = typeof timestamp === 'number' ? new Date(timestamp).toISOString() : ''
        if (label) {
            console.log(`[PluginHost][${label}]`, ...(entry.args || []))
        } else {
            console.log('[PluginHost]', ...(entry.args || []))
        }
    }

    if (canGroup) {
        console.groupEnd()
    }

    pluginDebugState.lastIndex = entries.length
    pluginDebugState.flushPending = false
}

function interpretDebugToggle(value) {
    if (value === true) {
        return true
    }
    if (typeof value === 'string') {
        const lowered = value.trim().toLowerCase()
        return lowered === 'true' || lowered === 'verbose' || lowered === 'debug'
    }
    return false
}

function installPluginDebugToggle() {
    if (typeof window === 'undefined') {
        return
    }
    if (window.__PLUGIN_DEBUG_TOGGLE_INSTALLED__) {
        return
    }

    let enabled = interpretDebugToggle(window.DEBUG_PLUGINS)
    window.__PLUGIN_DEBUG_ENABLED__ = enabled

    try {
        Object.defineProperty(window, 'DEBUG_PLUGINS', {
            configurable: true,
            enumerable: true,
            get() {
                return window.__PLUGIN_DEBUG_ENABLED__ === true
            },
            set(value) {
                const next = interpretDebugToggle(value)
                if (window.__PLUGIN_DEBUG_ENABLED__ !== next) {
                    window.__PLUGIN_DEBUG_ENABLED__ = next
                    const status = next ? 'enabled' : 'disabled'
                    recordPluginDebugEntry([`Plugin debug logging ${status}`])
                    if (typeof console !== 'undefined' && typeof console.info === 'function') {
                        console.info(`[PluginHost] Plugin debug logging ${status}`)
                    }
                    if (next) {
                        flushPluginDebugHistory()
                    }
                }
            }
        })
    } catch (err) {
        // Fall back gracefully if the property cannot be redefined (for example, due to CSP or readonly definitions).
        if (typeof console !== 'undefined' && typeof console.debug === 'function') {
            console.debug('[PluginHost] Unable to install DEBUG_PLUGINS setter', err)
        }
    }

    window.__PLUGIN_DEBUG_TOGGLE_INSTALLED__ = true
    if (typeof window.dumpPluginDebugHistory !== 'function') {
        window.dumpPluginDebugHistory = (options) => {
            const includeAll = Boolean(options && options.all)
            flushPluginDebugHistory(includeAll)
        }
    }
}

function isPluginDebugEnabled() {
    if (typeof window === 'undefined') {
        return false
    }
    if (typeof window.__PLUGIN_DEBUG_ENABLED__ === 'boolean') {
        return window.__PLUGIN_DEBUG_ENABLED__
    }
    const enabled = interpretDebugToggle(window.DEBUG_PLUGINS)
    window.__PLUGIN_DEBUG_ENABLED__ = enabled
    return enabled
}

installPluginDebugToggle()

function pluginDebugLog(...args) {
    recordPluginDebugEntry(args)
    if (!isPluginDebugEnabled()) {
        return
    }
    if (typeof console === 'undefined' || typeof console.log !== 'function') {
        return
    }
    try {
        console.log('[PluginHost]', ...args)
    } catch (_) {
        // Swallow logging errors to avoid impacting runtime flow when debugging is enabled.
    }
}

function updateButtonLabel(button, label) {
    if (!button) {
        return
    }

    const srLabel = button.querySelector('.sr-only')
    if (button.dataset.defaultLabel === undefined) {
        const inferred = typeof button.getAttribute('aria-label') === 'string' && button.getAttribute('aria-label')?.trim()
            ? button.getAttribute('aria-label')?.trim()
            : (typeof button.getAttribute('title') === 'string' && button.getAttribute('title')?.trim()
                ? button.getAttribute('title')?.trim()
                : (srLabel && srLabel.textContent ? srLabel.textContent.trim() : ''))
        if (inferred) {
            button.dataset.defaultLabel = inferred
        }
    }

    const normalized = typeof label === 'string' ? label.trim() : ''
    const effectiveLabel = normalized || button.dataset.defaultLabel || ''

    if (effectiveLabel) {
        button.setAttribute('aria-label', effectiveLabel)
        button.setAttribute('title', effectiveLabel)
    } else {
        button.removeAttribute('aria-label')
        button.removeAttribute('title')
    }

    if (srLabel) {
        srLabel.textContent = effectiveLabel
    }
}

function scrollChatHistoryToBottom(force = false) {
    const chatHistory = document.getElementById('chatHistory')
    if (!chatHistory) {
        return
    }

    const target = chatHistory

    const distanceFromBottom = target.scrollHeight - target.clientHeight - target.scrollTop
    if (!force && distanceFromBottom <= 8) {
        return
    }

    requestAnimationFrame(() => {
        target.scrollTop = target.scrollHeight
    })
}

const inlineCitationRegex = /\s*\u3010\d+:\d+\u2020[^\u3011]*\u3011/g

function stripInlineCitations(value) {
    if (typeof value !== 'string') {
        return value
    }

    const withoutCitations = value.replace(inlineCitationRegex, ' ')
    const normalizedWhitespace = withoutCitations
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n[ \t]+/g, '\n')
        .replace(/[ \t]{2,}/g, ' ')
        .replace(/\s+([\.,;:!?])/g, '$1')

    return normalizedWhitespace.trim()
}

function ensurePluginOverlayEventBindings() {
    const { overlay } = getAssistantAttachmentOverlayElements()
    if (!overlay || pluginOverlayEventsBound) {
        return
    }
    overlay.addEventListener('click', handlePluginOverlayActivate)
    overlay.addEventListener('keydown', handlePluginOverlayKeydown)
    pluginOverlayEventsBound = true
}

function handlePluginOverlayActivate(event) {
    const target = event.target instanceof Element ? event.target.closest('[data-plugin-action]') : null
    if (!target) {
        return
    }

    const action = target.getAttribute('data-plugin-action')
    if (!action) {
        return
    }

    if (target.getAttribute('data-disabled') === 'true' || target.dataset.disabled === 'true' || target.hasAttribute('disabled') || target.getAttribute('aria-disabled') === 'true') {
        event.preventDefault()
        event.stopPropagation()
        return
    }

    const detail = {
        action,
        row: target.dataset.row !== undefined ? Number(target.dataset.row) : undefined,
        column: target.dataset.column !== undefined ? Number(target.dataset.column) : undefined,
        gameId: target.dataset.gameId
    }

    if (!detail.gameId) {
        const container = target.closest('[data-game-id]')
        if (container && container.dataset.gameId) {
            detail.gameId = container.dataset.gameId
        }
    }

    invokePluginOverlayAction(action, detail)
    event.preventDefault()
    event.stopPropagation()
}

function handlePluginOverlayKeydown(event) {
    if (event.key !== 'Enter' && event.key !== ' ') {
        return
    }
    handlePluginOverlayActivate(event)
}

function getAssistantAttachmentOverlayElements() {
    const overlay = document.getElementById('assistantAttachmentOverlay')
    if (!overlay) {
        return { overlay: null, inner: null }
    }
    const inner = overlay.querySelector('.assistant-attachment-overlay__inner') || overlay
    return { overlay, inner }
}

function renderOverlayFromAttachments(attachments) {
    const { overlay, inner } = getAssistantAttachmentOverlayElements()
    if (!overlay || !inner) {
        return
    }

    const images = Array.isArray(attachments)
        ? attachments.filter((attachment) => attachment && attachment.type === 'image' && attachment.url)
        : []
    if (!images.length) {
        if (typeof inner.replaceChildren === 'function') {
            inner.replaceChildren()
        } else {
            inner.innerHTML = ''
        }
        overlay.hidden = true
        overlay.dataset.mode = 'hidden'
        delete overlay.dataset.pluginId
        if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
            console.log('[AgentAttachments] Overlay hidden (no images)')
        }
        return
    }

    let overlayContent

    if (images.length > 1) {
        overlayContent = createImageAttachmentCarousel(images)
    } else {
        const card = createImageAttachmentCard(images[0])
        card.classList.add('attachment-card--overlay')
        overlayContent = card
    }

    if (typeof inner.replaceChildren === 'function') {
        inner.replaceChildren(overlayContent)
    } else {
        inner.innerHTML = ''
        inner.appendChild(overlayContent)
    }

    overlay.hidden = false
    overlay.dataset.mode = 'attachments'
    delete overlay.dataset.pluginId

    if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
        console.log('[AgentAttachments] Overlay rendered with image:', images[0])
    }
}

function renderOverlayFromPlugin(descriptor) {
    const { overlay, inner } = getAssistantAttachmentOverlayElements()
    if (!overlay || !inner) {
        return
    }

    pluginDebugLog('Rendering plugin overlay', descriptor ? {
        title: descriptor.title || descriptor?.heading || descriptor?.label || '',
        hasHtml: typeof descriptor.html === 'string',
        hasAttachments: Array.isArray(descriptor.attachments),
        textPreview: typeof descriptor.text === 'string' ? descriptor.text.slice(0, 80) : undefined
    } : null)

    const requestedHidden = descriptor === null
        || descriptor === undefined
        || (typeof descriptor === 'object' && descriptor !== null && descriptor.visible === false)

    if (requestedHidden) {
        attachmentOverlayState.pluginDescriptor = null
        if (typeof inner.replaceChildren === 'function') {
            inner.replaceChildren()
        } else {
            inner.innerHTML = ''
        }
        overlay.hidden = true
        overlay.dataset.mode = 'hidden'
        return
    }

    if (descriptor && typeof descriptor === 'object' && Array.isArray(descriptor.attachments)) {
        const attachments = descriptor.attachments.filter(Boolean)
        attachmentOverlayState.attachments = attachments
        if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
            console.log('[AgentAttachments] Plugin supplied attachments override:', attachments)
        }
        renderOverlayFromAttachments(attachments)
        overlay.dataset.mode = 'plugin'
        overlay.dataset.pluginId = pluginHostState.activePluginId || ''
        return
    }

    const fragment = document.createDocumentFragment()

    const titleText = typeof descriptor === 'object' && descriptor !== null
        ? coerceDisplayString(descriptor.title || descriptor.heading || descriptor.label)
        : ''
    let titleAppended = false
    if (titleText) {
        const heading = document.createElement('div')
        heading.className = 'assistant-attachment-overlay__title'
        heading.textContent = titleText
        fragment.appendChild(heading)
        titleAppended = true
    }

    if (descriptor && typeof descriptor === 'object' && typeof descriptor.html === 'string' && descriptor.html.trim() !== '') {
        const htmlWrapper = document.createElement('div')
        htmlWrapper.className = 'assistant-attachment-overlay__custom-html'
        htmlWrapper.innerHTML = descriptor.html
        fragment.appendChild(htmlWrapper)
    } else {
        let textContent = ''
        if (descriptor && typeof descriptor === 'object' && descriptor !== null) {
            textContent = coerceDisplayString(descriptor.text || descriptor.body)
        } else if (descriptor !== null && descriptor !== undefined) {
            textContent = coerceDisplayString(descriptor)
        }

        if (textContent) {
            const paragraph = document.createElement('p')
            paragraph.className = 'assistant-attachment-overlay__text'
            paragraph.textContent = textContent
            fragment.appendChild(paragraph)
        }

        if (descriptor && typeof descriptor === 'object' && Array.isArray(descriptor.items) && descriptor.items.length > 0) {
            const list = document.createElement('ul')
            list.className = 'assistant-attachment-overlay__list'
            descriptor.items.forEach((item) => {
                const normalized = coerceDisplayString(typeof item === 'string' ? item : (item && item.label) || (item && item.title) || '')
                if (!normalized) {
                    return
                }
                const listItem = document.createElement('li')
                listItem.textContent = normalized
                list.appendChild(listItem)
            })
            if (list.childNodes.length > 0) {
                fragment.appendChild(list)
            }
        }
    }

    if (fragment.childNodes.length === 0) {
        const fallback = document.createElement('pre')
        fallback.className = 'assistant-attachment-overlay__fallback'
        const serialized = typeof descriptor === 'object' ? JSON.stringify(descriptor, null, 2) : coerceDisplayString(descriptor)
        fallback.textContent = serialized || 'Content unavailable'
        fragment.appendChild(fallback)
    } else if (titleAppended && fragment.childNodes.length === 1) {
        fragment.removeChild(fragment.firstChild)
    }

    if (typeof inner.replaceChildren === 'function') {
        inner.replaceChildren(fragment)
    } else {
        inner.innerHTML = ''
        inner.appendChild(fragment)
    }

    overlay.hidden = false
    overlay.dataset.mode = 'plugin'
    overlay.dataset.pluginId = pluginHostState.activePluginId || ''

    ensurePluginOverlayEventBindings()
}

function setPluginOverlayDescriptor(descriptor) {
    pluginDebugLog('Overlay descriptor update', descriptor ? {
        hasHtml: typeof descriptor.html === 'string',
        title: descriptor.title || descriptor?.name || '',
        textPreview: typeof descriptor.text === 'string' ? descriptor.text.slice(0, 80) : undefined
    } : null)
    attachmentOverlayState.pluginDescriptor = descriptor ?? null
    renderOverlayFromPlugin(attachmentOverlayState.pluginDescriptor)
    if (!attachmentOverlayState.pluginDescriptor) {
        renderOverlayFromAttachments(attachmentOverlayState.attachments)
    }
}

function registerAgentContentListener(handler) {
    if (typeof handler !== 'function') {
        return () => {}
    }
    pluginHostState.agentContentListeners.add(handler)
    pluginDebugLog('Registered agent content listener', {
        listenerCount: pluginHostState.agentContentListeners.size
    })
    return () => {
        pluginHostState.agentContentListeners.delete(handler)
        pluginDebugLog('Agent content listener removed', {
            listenerCount: pluginHostState.agentContentListeners.size
        })
    }
}

function clearAgentContentListeners() {
    if (pluginHostState.agentContentListeners.size > 0) {
        pluginDebugLog('Clearing agent content listeners', {
            count: pluginHostState.agentContentListeners.size
        })
    }
    pluginHostState.agentContentListeners.clear()
}

function notifyAgentContentObservers(payload) {
    if (!pluginHostState.agentContentListeners.size) {
        return
    }
    const attachmentCount = Array.isArray(payload?.__attachments) ? payload.__attachments.length : undefined
    const identifier = payload?.message?.id
        || payload?.data?.message?.id
        || payload?.message_id
        || payload?.id
        || payload?.type
        || 'unknown'
    pluginDebugLog('Dispatching agent content payload', {
        listenerCount: pluginHostState.agentContentListeners.size,
        identifier,
        attachmentCount
    })
    pluginHostState.agentContentListeners.forEach((listener) => {
        try {
            listener(payload)
        } catch (err) {
            console.warn('[PluginHost] Agent content listener failed', err)
        }
    })
}

function pluginHostSendTextInput(text, options = {}) {
    const normalized = typeof text === 'string' ? text.trim() : ''
    if (!normalized) {
        return Promise.resolve(false)
    }
    const html = typeof options.html === 'string' && options.html.trim() !== '' ? options.html : normalized
    const imageUrl = typeof options.imageUrl === 'string' ? options.imageUrl : ''
    pluginDebugLog('Plugin emitting user text', {
        textPreview: normalized.slice(0, 80),
        hasHtml: html !== normalized,
        imageUrl: imageUrl ? '[provided]' : undefined
    })
    try {
        const result = handleUserQuery(normalized, html, imageUrl)
        return result === undefined ? Promise.resolve(true) : Promise.resolve(result)
    } catch (err) {
        console.error('[PluginHost] sendTextInput failed', err)
        pluginDebugLog('Plugin sendTextInput error', err?.message || err)
        return Promise.reject(err)
    }
}

function resolvePluginEntryUrl(pluginRecord) {
    if (!pluginRecord || typeof pluginRecord !== 'object') {
        return null
    }
    const assets = pluginRecord.assets && typeof pluginRecord.assets === 'object' ? pluginRecord.assets : {}
    const candidates = []

    const pushCandidate = (value) => {
        if (typeof value !== 'string') {
            return
        }
        const trimmed = value.trim()
        if (trimmed) {
            candidates.push(trimmed)
        }
    }

    pushCandidate(assets.clientEntry)
    pushCandidate(assets.clientModule)
    pushCandidate(assets.entry)
    pushCandidate(assets.module)
    pushCandidate(assets.url)

    if (assets.client && typeof assets.client === 'object') {
        pushCandidate(assets.client.entry)
        pushCandidate(assets.client.module)
        if (typeof assets.client.path === 'string' && typeof assets.client.entry === 'string') {
            const basePath = assets.client.path.replace(/\/$/, '')
            const entry = assets.client.entry.replace(/^\.\//, '')
            pushCandidate(`${basePath}/${entry}`)
        }
    }

    if (typeof assets.path === 'string' && typeof assets.entry === 'string') {
        const basePath = assets.path.replace(/\/$/, '')
        const entry = assets.entry.replace(/^\.\//, '')
        pushCandidate(`${basePath}/${entry}`)
    }

    if (typeof assets.baseUrl === 'string' && typeof assets.entry === 'string') {
        const base = assets.baseUrl.replace(/\/$/, '')
        const entry = assets.entry.replace(/^\.\//, '')
        pushCandidate(`${base}/${entry}`)
    }

    return candidates.find((candidate) => candidate.length > 0) || null
}

function resolvePluginStyleUrl(pluginRecord) {
    if (!pluginRecord || typeof pluginRecord !== 'object') {
        return null
    }

    const assets = pluginRecord.assets && typeof pluginRecord.assets === 'object' ? pluginRecord.assets : {}
    const candidates = []

    const pushCandidate = (value) => {
        if (typeof value !== 'string') {
            return
        }
        const trimmed = value.trim()
        if (trimmed) {
            candidates.push(trimmed)
        }
    }

    pushCandidate(pluginRecord.style)
    pushCandidate(pluginRecord.stylesheet)
    pushCandidate(assets.clientStyle)
    pushCandidate(assets.style)
    pushCandidate(assets.stylesheet)

    if (assets.client && typeof assets.client === 'object') {
        pushCandidate(assets.client.style)
        pushCandidate(assets.client.stylesheet)
        if (typeof assets.client.path === 'string' && typeof assets.client.style === 'string') {
            const basePath = assets.client.path.replace(/\/$/, '')
            const style = assets.client.style.replace(/^\.\//, '')
            pushCandidate(`${basePath}/${style}`)
        }
    }

    if (typeof assets.path === 'string' && typeof assets.style === 'string') {
        const basePath = assets.path.replace(/\/$/, '')
        const style = assets.style.replace(/^\.\//, '')
        pushCandidate(`${basePath}/${style}`)
    }

    if (typeof assets.baseUrl === 'string' && typeof assets.style === 'string') {
        const base = assets.baseUrl.replace(/\/$/, '')
        const style = assets.style.replace(/^\.\//, '')
        pushCandidate(`${base}/${style}`)
    }

    return candidates.find((candidate) => candidate.length > 0) || null
}

function loadPluginStylesheet(styleUrl) {
    if (typeof document === 'undefined' || !styleUrl) {
        return null
    }

    const head = document.head || document.getElementsByTagName('head')[0]
    if (!head) {
        return null
    }

    const existing = head.querySelector(`link[data-plugin-style="${styleUrl}"]`)
    if (existing) {
        return existing
    }

    const link = document.createElement('link')
    link.rel = 'stylesheet'
    link.type = 'text/css'
    link.href = styleUrl
    link.dataset.pluginStyle = styleUrl

    head.appendChild(link)
    return link
}

function unloadActivePluginStylesheet() {
    const element = pluginHostState.activeStylesheetElement
    if (!element) {
        pluginHostState.activeStylesheetUrl = null
        return
    }

    try {
        if (element.parentNode) {
            element.parentNode.removeChild(element)
        }
    } catch (err) {
        console.warn('[PluginHost] Failed to remove plugin stylesheet', err)
    }

    pluginHostState.activeStylesheetElement = null
    pluginHostState.activeStylesheetUrl = null
}

async function loadPluginModule(entryUrl) {
    if (pluginHostState.moduleCache.has(entryUrl)) {
        return pluginHostState.moduleCache.get(entryUrl)
    }
    const loader = import(/* webpackIgnore: true */ entryUrl)
    pluginHostState.moduleCache.set(entryUrl, loader)
    return loader
}

function resolvePluginConstructor(moduleExports, pluginRecord) {
    if (!moduleExports || typeof moduleExports !== 'object') {
        return null
    }

    if (typeof moduleExports.default === 'function') {
        return moduleExports.default
    }

    const assets = pluginRecord && typeof pluginRecord.assets === 'object' ? pluginRecord.assets : {}
    const clientConfig = assets && typeof assets.client === 'object' ? assets.client : {}
    const exportName = clientConfig.export || clientConfig.class || assets.export || assets.class || pluginRecord?.className || pluginRecord?.export

    if (typeof exportName === 'string' && typeof moduleExports[exportName] === 'function') {
        return moduleExports[exportName]
    }

    const firstFunction = Object.values(moduleExports).find((value) => typeof value === 'function')
    return typeof firstFunction === 'function' ? firstFunction : null
}

function destroyActivePlugin(options = {}) {
    const preservePendingToken = options && options.preservePendingToken === true
    if (pluginHostState.activePluginId) {
        pluginDebugLog('Destroying active plugin', {
            pluginId: pluginHostState.activePluginId,
            hasInstance: !!pluginHostState.activeInstance
        })
    }
    const instance = pluginHostState.activeInstance
    if (instance && typeof instance.destroy === 'function') {
        try {
            instance.destroy()
        } catch (err) {
            console.warn('[PluginHost] Plugin destroy failed', err)
            pluginDebugLog('Plugin destroy threw', err?.message || err)
        }
    }
    pluginHostState.activeInstance = null
    pluginHostState.activePluginId = null
    pluginHostState.lastManifest = null
    if (!preservePendingToken) {
        pluginHostState.pendingActivationToken = null
    }
    clearAgentContentListeners()
    setPluginOverlayDescriptor(null)
    unloadActivePluginStylesheet()
    pluginDebugLog('Plugin host reset complete')
}

function invokePluginOverlayAction(action, detail) {
    const instance = pluginHostState.activeInstance
    if (!instance) {
        return
    }

    const handler = typeof instance.handleOverlayAction === 'function'
        ? instance.handleOverlayAction.bind(instance)
        : typeof instance.onOverlayAction === 'function'
            ? instance.onOverlayAction.bind(instance)
            : null

    if (!handler) {
        return
    }

    pluginDebugLog('Overlay action invoked', {
        action,
        detail
    })

    try {
        handler(action, detail || {})
    } catch (err) {
        console.error('[PluginHost] Overlay action handler failed', err)
        pluginDebugLog('Overlay action handler error', err?.message || err)
    }
}

async function activateAgentPlugin(pluginRecord) {
    pluginDebugLog('Activation requested', {
        recordId: pluginRecord?.id || pluginRecord?.pluginId || pluginRecord?.key || null,
        hasAssets: !!(pluginRecord && typeof pluginRecord === 'object' && pluginRecord.assets)
    })
    const activationToken = {}
    pluginHostState.pendingActivationToken = activationToken

    if (!pluginRecord || typeof pluginRecord !== 'object') {
        pluginDebugLog('Activation aborted: invalid plugin record', pluginRecord)
        destroyActivePlugin()
        return
    }

    const pluginId = pluginRecord.id || pluginRecord.pluginId || pluginRecord.key
    if (!pluginId) {
        pluginDebugLog('Activation aborted: missing plugin identifier', pluginRecord)
        destroyActivePlugin()
        return
    }

    if (pluginHostState.activePluginId === pluginId && pluginHostState.activeInstance) {
        pluginDebugLog('Activation skipped: plugin already active', pluginId)
        pluginHostState.pendingActivationToken = null
        return
    }

    destroyActivePlugin({ preservePendingToken: true })

    const entryUrl = resolvePluginEntryUrl(pluginRecord)
    const styleUrl = resolvePluginStyleUrl(pluginRecord)
    pluginDebugLog('Resolved plugin entry URL', { pluginId, entryUrl })
    if (!entryUrl) {
        if (typeof window !== 'undefined' && window.DEBUG_PLUGINS === true) {
            console.warn('[PluginHost] No client entry specified for plugin', pluginId)
        }
        return
    }

    let styleElement = null
    if (styleUrl) {
        pluginDebugLog('Loading plugin stylesheet', { pluginId, styleUrl })
        styleElement = loadPluginStylesheet(styleUrl)
        if (!styleElement) {
            pluginDebugLog('Plugin stylesheet load failed or DOM unavailable', { pluginId, styleUrl })
        }
    }

    try {
        pluginDebugLog('Loading plugin module', { pluginId, entryUrl })
        const moduleExports = await loadPluginModule(entryUrl)
        pluginDebugLog('Module loaded', { pluginId, keys: Object.keys(moduleExports || {}) })
        if (pluginHostState.pendingActivationToken !== activationToken) {
            if (styleElement && styleElement.parentNode) {
                try {
                    styleElement.parentNode.removeChild(styleElement)
                } catch (removeErr) {
                    console.warn('[PluginHost] Failed to remove plugin stylesheet after module load cancellation', removeErr)
                }
            }
            pluginDebugLog('Activation cancelled before module resolution', pluginId)
            return
        }

        const PluginConstructor = resolvePluginConstructor(moduleExports, pluginRecord)
        if (typeof PluginConstructor !== 'function') {
            console.error('[PluginHost] Plugin module does not export a constructor', pluginId, moduleExports)
            pluginDebugLog('Activation failed: constructor not found', { pluginId, exports: Object.keys(moduleExports || {}) })
            if (styleElement && styleElement.parentNode) {
                try {
                    styleElement.parentNode.removeChild(styleElement)
                } catch (removeErr) {
                    console.warn('[PluginHost] Failed to remove plugin stylesheet after constructor resolution failure', removeErr)
                }
            }
            return
        }
        pluginDebugLog('Instantiating plugin', pluginId)

        const api = {
            log: (...args) => console.log(`[Plugin:${pluginId}]`, ...args),
            config: pluginRecord.config || {},
            sendTextInput: (text, meta) => pluginHostSendTextInput(text, meta),
            onAgentContent: (handler) => registerAgentContentListener(handler),
            updateAttachmentOverlay: (descriptor) => setPluginOverlayDescriptor(descriptor),
            setAttachmentOverlay: (descriptor) => setPluginOverlayDescriptor(descriptor),
            setOverlayContent: (descriptor) => setPluginOverlayDescriptor(descriptor),
            clearAttachmentOverlay: () => setPluginOverlayDescriptor(null),
            clearOverlayContent: () => setPluginOverlayDescriptor(null)
        }

        const instance = new PluginConstructor(pluginRecord, api)
        pluginDebugLog('Plugin constructed', {
            pluginId,
            hasOnAgentContent: typeof instance.onAgentContent === 'function',
            hasOverlayActions: typeof instance.handleOverlayAction === 'function' || typeof instance.onOverlayAction === 'function'
        })
        const initResult = typeof instance.init === 'function' ? instance.init() : undefined
        if (initResult && typeof initResult.then === 'function') {
            pluginDebugLog('Awaiting async init', pluginId)
            await initResult
        }

        if (pluginHostState.pendingActivationToken !== activationToken) {
            if (typeof instance.destroy === 'function') {
                instance.destroy()
            }
            if (styleElement && styleElement.parentNode) {
                try {
                    styleElement.parentNode.removeChild(styleElement)
                } catch (removeErr) {
                    console.warn('[PluginHost] Failed to remove plugin stylesheet after activation cancellation', removeErr)
                }
            }
            pluginDebugLog('Activation cancelled post-init', pluginId)
            return
        }

        pluginHostState.activePluginId = pluginId
        pluginHostState.activeInstance = instance
        pluginHostState.lastManifest = pluginRecord
        pluginHostState.activeStylesheetElement = styleElement || null
        pluginHostState.activeStylesheetUrl = styleElement ? styleUrl : null
        pluginDebugLog('Plugin activated', pluginId)
        
        // Apply plugin's avatar background image and pre-session video
        applyAvatarBackground(getAvatarBackgroundImage())
        applyPreSessionVideo(getPreSessionVideoUrl())
    } catch (err) {
        console.error('[PluginHost] Failed to activate plugin', pluginRecord?.id || pluginRecord, err)
        pluginDebugLog('Activation error', err?.message || err)
        if (styleElement && styleElement.parentNode) {
            try {
                styleElement.parentNode.removeChild(styleElement)
            } catch (removeErr) {
                console.warn('[PluginHost] Failed to remove plugin stylesheet after activation error', removeErr)
            }
        }
    } finally {
        if (pluginHostState.pendingActivationToken === activationToken) {
            pluginHostState.pendingActivationToken = null
        }
    }
}

function getQueryParameter(name) {
    if (typeof window === 'undefined' || typeof name !== 'string') {
        return ''
    }
    try {
        const params = new URLSearchParams(window.location.search)
        const value = params.get(name)
        return value ? value.trim() : ''
    } catch (err) {
        console.debug('Unable to read query parameter', name, err)
        return ''
    }
}

function initializeAgentOverrideFromQuery() {
    const override = getQueryParameter('agent')
    if (override) {
        agentSelectorState.overrideQuery = override
    }

    if (typeof window !== 'undefined') {
        try {
            const params = new URLSearchParams(window.location.search)
            const collected = params.getAll('tags')
            const tags = []
            collected.forEach((value) => {
                if (typeof value !== 'string') {
                    return
                }
                value.split(',').forEach((part) => {
                    const trimmed = part.trim().toLowerCase()
                    if (trimmed && !tags.includes(trimmed)) {
                        tags.push(trimmed)
                    }
                })
            })
            agentSelectorState.tagFilters = tags
        } catch (err) {
            console.debug('Unable to parse tags query parameter', err)
            agentSelectorState.tagFilters = []
        }
    }
}

const MAX_ASSISTANT_ATTACHMENTS = 6
const ATTACHMENT_SUMMARY_LIMIT = 800

function getStartSessionButton() {
    const button = document.getElementById('startSession')
    if (button && startSessionButtonOriginalLabel === null) {
        startSessionButtonOriginalLabel = button.textContent || 'Talk to Avatar'
    }
    return button
}

function showLocalIdleVideo() {
    console.log("show local video")
    if (!isLocalIdleVideoEnabled()) {
        return
    }
    const localVideo = document.getElementById('localVideo')
    const remoteVideoContainer = document.getElementById('remoteVideo')
    if (localVideo) {
        localVideo.hidden = false
        localVideo.style.width = '100%'
        localVideo.style.height = '100vh'
        localVideo.style.visibility = 'visible'
    }
    if (remoteVideoContainer) {
        remoteVideoContainer.hidden = true
    }
}

function hideLocalIdleVideo() {
    console.log("hide local video")
    const localVideo = document.getElementById('localVideo')
    const remoteVideoContainer = document.getElementById('remoteVideo')
    if (localVideo) {
        localVideo.hidden = true
        localVideo.style.width = '0px'
        localVideo.style.height = '0vh'
        localVideo.style.visibility = 'hidden'
    }
    if (remoteVideoContainer) {
        remoteVideoContainer.hidden = false
        remoteVideoContainer.style.width = ''
        remoteVideoContainer.style.visibility = 'visible'
    }
}

function getAvatarBackgroundImage() {
    // Priority 1: Active plugin's branding background image
    if (pluginHostState.lastManifest && typeof pluginHostState.lastManifest === 'object') {
        const pluginBranding = pluginHostState.lastManifest.branding
        if (pluginBranding && typeof pluginBranding === 'object' && typeof pluginBranding.backgroundImage === 'string') {
            const pluginTrimmed = pluginBranding.backgroundImage.trim()
            if (pluginTrimmed !== '') {
                return pluginTrimmed
            }
        }
    }
    
    // Priority 2: Runtime config avatar section
    const section = getRuntimeSection('avatar')
    const value = section.backgroundImage
    if (typeof value === 'string') {
        const trimmed = value.trim()
        if (trimmed !== '') {
            return trimmed
        }
    }
    
    // Priority 3: Active branding config background image
    if (activeBrandingConfig && typeof activeBrandingConfig.backgroundImage === 'string') {
        const activeTrimmed = activeBrandingConfig.backgroundImage.trim()
        if (activeTrimmed !== '') {
            return activeTrimmed
        }
    }
    
    // Priority 4: Default branding config background image
    const brandingBackground = runtimeConfig?.branding?.backgroundImage
    if (typeof brandingBackground === 'string') {
        const brandingTrimmed = brandingBackground.trim()
        if (brandingTrimmed !== '') {
            return brandingTrimmed
        }
    }
    return ''
}

function applyAvatarBackground(imageUrl) {
    const wrapper = document.querySelector('.video-wrapper')
    if (!wrapper) {
        return
    }
    if (isTransparentBackgroundEnabled()) {
        wrapper.style.backgroundColor = 'transparent'
    } else {
        wrapper.style.backgroundColor = ''
    }
    if (typeof imageUrl === 'string' && imageUrl.trim() !== '') {
        const resolvedUrl = resolveAssetUrl(imageUrl)
        if (resolvedUrl) {
            const sanitized = resolvedUrl.replace(/"/g, '\\"')
            wrapper.style.backgroundImage = `url("${sanitized}")`
            wrapper.style.backgroundSize = 'cover'
            wrapper.style.backgroundPosition = 'center'
            wrapper.style.backgroundRepeat = 'no-repeat'
        } else {
            wrapper.style.backgroundImage = ''
        }
    } else {
        wrapper.style.backgroundImage = ''
    }
}

function getPreSessionVideoUrl() {
    // Priority 1: Active plugin's local idle video or video
    if (pluginHostState.lastManifest && typeof pluginHostState.lastManifest === 'object') {
        const pluginAvatar = pluginHostState.lastManifest.avatar
        if (pluginAvatar && typeof pluginAvatar === 'object') {
            const localIdleVideo = typeof pluginAvatar.localIdleVideo === 'string' ? pluginAvatar.localIdleVideo.trim() : ''
            if (localIdleVideo) {
                return localIdleVideo
            }
            const video = typeof pluginAvatar.video === 'string' ? pluginAvatar.video.trim() : ''
            if (video) {
                return video
            }
        }
    }
    
    // Priority 2: Runtime config avatar section
    const section = getRuntimeSection('avatar')
    const localIdleVideo = typeof section.localIdleVideo === 'string' ? section.localIdleVideo.trim() : ''
    if (localIdleVideo) {
        return localIdleVideo
    }
    const video = typeof section.video === 'string' ? section.video.trim() : ''
    if (video) {
        return video
    }
    
    // Fallback: default video
    return 'video/meg-formal-idle.webm'
}

function applyPreSessionVideo(videoUrl) {
    const preSessionVideo = document.querySelector('.pre-session-video')
    if (!preSessionVideo) {
        return
    }
    
    if (typeof videoUrl === 'string' && videoUrl.trim() !== '') {
        const resolvedUrl = resolveAssetUrl(videoUrl)
        if (resolvedUrl && resolvedUrl !== preSessionVideo.src) {
            preSessionVideo.src = resolvedUrl
            // Reload the video to apply the new source
            preSessionVideo.load()
        }
    }
}

function isTransparentBackgroundEnabled() {
    return getRuntimeBoolean('avatar', 'transparentBackground', false)
}

function setStartSessionButtonState(options) {
    const button = getStartSessionButton()
    if (!button) {
        return
    }

    const { disabled, label, busy } = options || {}

    if (typeof disabled === 'boolean') {
        button.disabled = disabled
    }

    if (typeof busy === 'boolean') {
        if (busy) {
            button.setAttribute('aria-busy', 'true')
        } else {
            button.removeAttribute('aria-busy')
        }
    }

    if (label !== undefined) {
        if (label === null) {
            button.textContent = startSessionButtonOriginalLabel || button.textContent
        } else {
            button.textContent = label
        }
    } else if (button.disabled === false && startSessionButtonOriginalLabel !== null) {
        button.textContent = startSessionButtonOriginalLabel
    }
}

function getStoredAgentSelectionKey() {
    try {
        const stored = window.localStorage?.getItem?.(AGENT_SELECTOR_STORAGE_KEY)
        if (typeof stored === 'string' && stored.trim() !== '') {
            return stored.trim()
        }
    } catch (err) {
        console.debug('Agent selector storage unavailable', err)
    }
    return ''
}

function setStoredAgentSelectionKey(key) {
    try {
        if (!key) {
            window.localStorage?.removeItem?.(AGENT_SELECTOR_STORAGE_KEY)
            return
        }
        window.localStorage?.setItem?.(AGENT_SELECTOR_STORAGE_KEY, key)
    } catch (err) {
        console.debug('Unable to persist agent selector choice', err)
    }
}

function setRuntimeConfigAssetBase(candidateUrl) {
    if (typeof candidateUrl !== 'string') {
        return
    }
    const trimmed = candidateUrl.trim()
    if (!trimmed) {
        return
    }

    const baseHref = (typeof window !== 'undefined' && window.location) ? window.location.href : undefined
    try {
        const resolved = baseHref ? new URL(trimmed, baseHref) : new URL(trimmed)
        runtimeConfigAssetBaseUrl = new URL('.', resolved).href
        return
    } catch (err) {
        // ignore and fall back to origin lookup
    }

    try {
        if (typeof window !== 'undefined' && window.location?.origin) {
            const resolved = new URL(trimmed, window.location.origin)
            runtimeConfigAssetBaseUrl = new URL('.', resolved).href
        }
    } catch (err) {
        // ignore
    }
}

function resolveAssetUrl(value) {
    if (typeof value !== 'string') {
        return ''
    }
    const trimmed = value.trim()
    if (!trimmed) {
        return ''
    }

    if (/^(?:[a-z][a-z0-9+.-]*:|\/{2})/i.test(trimmed)) {
        return trimmed
    }

    if (trimmed.startsWith('/')) {
        if (typeof window !== 'undefined' && window.location?.origin) {
            return `${window.location.origin}${trimmed}`
        }
        return trimmed
    }

    const base = runtimeConfigAssetBaseUrl || ((typeof window !== 'undefined' && window.location) ? window.location.href : '')
    if (base) {
        try {
            return new URL(trimmed, base).href
        } catch (err) {
            // ignore and fall back
        }
    }

    if (typeof window !== 'undefined' && window.location?.origin) {
        try {
            return new URL(trimmed, window.location.origin).href
        } catch (err) {
            // ignore and fall back
        }
    }

    return trimmed
}

function sanitizeBrandingConfig(rawBranding) {
    if (!rawBranding || typeof rawBranding !== 'object') {
        return null
    }

    const branding = {}
    const assignIfString = (key, transform) => {
        const value = rawBranding[key]
        if (typeof value === 'string') {
            const trimmed = value.trim()
            if (trimmed !== '') {
                const resolved = transform ? transform(trimmed) : trimmed
                if (typeof resolved === 'string' && resolved.trim() !== '') {
                    branding[key] = resolved.trim()
                }
            }
        }
    }

    assignIfString('primaryColor')
    assignIfString('backgroundImage', resolveAssetUrl)
    assignIfString('logoUrl', resolveAssetUrl)
    assignIfString('logoAlt')

    return Object.keys(branding).length > 0 ? branding : null
}

function cloneBrandingConfig(branding) {
    if (!branding || typeof branding !== 'object') {
        return null
    }
    const clone = {}
    const keys = ['primaryColor', 'backgroundImage', 'logoUrl', 'logoAlt']
    keys.forEach((key) => {
        if (typeof branding[key] === 'string') {
            const trimmed = branding[key].trim()
            if (trimmed !== '') {
                clone[key] = trimmed
            }
        }
    })
    return Object.keys(clone).length > 0 ? clone : null
}

function mergeBranding(baseBranding, overrideBranding) {
    const result = {}
    const applySource = (source) => {
        if (!source || typeof source !== 'object') {
            return
        }
        const keys = ['primaryColor', 'backgroundImage', 'logoUrl', 'logoAlt']
        keys.forEach((key) => {
            const value = source[key]
            if (typeof value === 'string') {
                const trimmed = value.trim()
                if (trimmed !== '') {
                    if (key === 'backgroundImage' || key === 'logoUrl') {
                        const resolved = resolveAssetUrl(trimmed)
                        if (resolved) {
                            result[key] = resolved
                        }
                    } else {
                        result[key] = trimmed
                    }
                }
            }
        })
    }

    applySource(baseBranding)
    applySource(overrideBranding)

    if (Object.keys(result).length === 0) {
        return null
    }
    return result
}

function cloneDeep(value) {
    if (value === undefined) {
        return undefined
    }
    try {
        return JSON.parse(JSON.stringify(value))
    } catch (err) {
        console.debug('Unable to clone value deeply', err)
        return undefined
    }
}

function deepMergeInto(target, source) {
    if (!target || typeof target !== 'object') {
        return target
    }
    if (!source || typeof source !== 'object') {
        return target
    }
    Object.entries(source).forEach(([key, value]) => {
        if (value === undefined) {
            return
        }
        if (Array.isArray(value)) {
            target[key] = value.slice()
            return
        }
        if (value && typeof value === 'object') {
            if (!target[key] || typeof target[key] !== 'object') {
                target[key] = {}
            }
            deepMergeInto(target[key], value)
            return
        }
        target[key] = value
    })
    return target
}

function captureRuntimeConfigDefaults(config, baseBranding) {
    runtimeConfigDefaults = {
        agent: cloneDeep(config?.agent || {}) || {},
        avatar: cloneDeep(config?.avatar || {}) || {},
        speech: cloneDeep(config?.speech || {}) || {},
        branding: cloneBrandingConfig(baseBranding) || cloneBrandingConfig(config?.branding) || null,
        features: cloneDeep(config?.features || {}) || {},
        ui: cloneDeep(config?.ui || {}) || {},
        conversation: cloneDeep(config?.conversation || {}) || {}
    }
}

function applyFeatureSettings() {
    const fallbackEnabled = runtimeConfigDefaults?.features?.quickReplyEnabled ?? false
    const enabled = getRuntimeBoolean('features', 'quickReplyEnabled', fallbackEnabled)
    enableQuickReply = Boolean(enabled)

    const fallbackOptions = Array.isArray(runtimeConfigDefaults?.features?.quickReplyOptions)
        ? runtimeConfigDefaults.features.quickReplyOptions
        : defaultQuickReplies
    const options = getRuntimeArray('features', 'quickReplyOptions', fallbackOptions)
    if (Array.isArray(options) && options.length > 0) {
        quickReplies = options.slice()
    } else {
        quickReplies = defaultQuickReplies.slice()
    }
}

function coerceCssUrl(value) {
    if (typeof value !== 'string') {
        return ''
    }
    const trimmed = value.trim()
    if (!trimmed) {
        return ''
    }
    if (/^url\(/i.test(trimmed)) {
        return trimmed
    }
    const resolved = resolveAssetUrl(trimmed)
    if (!resolved) {
        return ''
    }
    const sanitized = resolved.replace(/"/g, '\\"')
    return `url("${sanitized}")`
}

function applyBrandingSettings(brandingConfig) {
    const root = document.documentElement
    if (!root) {
        return
    }

    const branding = brandingConfig && typeof brandingConfig === 'object' ? { ...brandingConfig } : {}

    const primaryColor = typeof branding.primaryColor === 'string' ? branding.primaryColor.trim() : ''
    if (primaryColor) {
        root.style.setProperty('--branding-primary', primaryColor)
    } else {
        root.style.removeProperty('--branding-primary')
    }

    if (typeof branding.backgroundImage === 'string' && branding.backgroundImage.trim() !== '') {
        const resolvedBackground = resolveAssetUrl(branding.backgroundImage)
        if (resolvedBackground) {
            branding.backgroundImage = resolvedBackground
        }
    }

    const backgroundValue = coerceCssUrl(branding.backgroundImage)
    if (backgroundValue) {
        root.style.setProperty('--branding-background-image', backgroundValue)
    } else {
        root.style.removeProperty('--branding-background-image')
    }

    const preSessionVideo = document.querySelector('.pre-session-video')
    if (preSessionVideo) {
        if (backgroundValue) {
            preSessionVideo.style.backgroundImage = backgroundValue
        } else {
            preSessionVideo.style.removeProperty('background-image')
        }
    }

    const headerLogo = document.querySelector('#header img')
    if (headerLogo) {
        if (!headerLogo.dataset.defaultSrc) {
            headerLogo.dataset.defaultSrc = headerLogo.getAttribute('data-default-src') || headerLogo.getAttribute('src') || ''
        }
        if (!headerLogo.dataset.defaultAlt) {
            headerLogo.dataset.defaultAlt = headerLogo.getAttribute('data-default-alt') || headerLogo.getAttribute('alt') || ''
        }

        const logoUrl = typeof branding.logoUrl === 'string' ? resolveAssetUrl(branding.logoUrl).trim() : ''
        if (logoUrl) {
            headerLogo.src = logoUrl
            branding.logoUrl = logoUrl
        } else if (headerLogo.dataset.defaultSrc) {
            headerLogo.src = headerLogo.dataset.defaultSrc
        }

        const logoAlt = typeof branding.logoAlt === 'string' ? branding.logoAlt.trim() : ''
        if (logoAlt) {
            headerLogo.alt = logoAlt
        } else if (headerLogo.dataset.defaultAlt) {
            headerLogo.alt = headerLogo.dataset.defaultAlt
        }
    }

    activeBrandingConfig = Object.keys(branding).length > 0 ? branding : null
}

function computeAgentOptionBadge(label) {
    if (typeof label !== 'string' || !label.trim()) {
        return 'AI'
    }
    const parts = label
        .trim()
        .split(/\s+/)
        .filter((part) => part.length > 0)
        .slice(0, 2)
    const initials = parts.map((part) => part[0]).join('').toUpperCase()
    return initials.slice(0, 3) || 'AI'
}

function normalizeAgentOptions(agentConfig) {
    if (!agentConfig || typeof agentConfig !== 'object') {
        return []
    }

    const rawOptions = Array.isArray(agentConfig.options) ? agentConfig.options : []
    const normalized = []
    const seenKeys = new Set()

    rawOptions.forEach((entry, index) => {
        if (!entry || typeof entry !== 'object') {
            return
        }

        const agentIdCandidate = typeof entry.agentId === 'string' ? entry.agentId.trim() : ''
        const agentIdFallback = typeof entry.id === 'string' ? entry.id.trim() : ''
        const agentId = agentIdCandidate || agentIdFallback
        if (!agentId) {
            return
        }

        const labelSource = typeof entry.label === 'string' && entry.label.trim() !== ''
            ? entry.label.trim()
            : (typeof entry.name === 'string' && entry.name.trim() !== '' ? entry.name.trim() : agentId)

        const keyCandidate = typeof entry.key === 'string' && entry.key.trim() !== ''
            ? entry.key.trim()
            : (typeof entry.id === 'string' && entry.id.trim() !== '' ? entry.id.trim() : agentId)

        let key = keyCandidate
        while (seenKeys.has(key)) {
            key = `${keyCandidate}-${index}`
        }
        seenKeys.add(key)

        const projectId = typeof entry.projectId === 'string' ? entry.projectId.trim() : ''
        const endpoint = typeof entry.endpoint === 'string' ? entry.endpoint.trim() : ''
        const apiUrl = typeof entry.apiUrl === 'string' ? entry.apiUrl.trim() : ''
        const description = typeof entry.description === 'string' ? entry.description.trim() : ''
        const badge = typeof entry.badge === 'string' && entry.badge.trim() !== ''
            ? entry.badge.trim().slice(0, 3).toUpperCase()
            : computeAgentOptionBadge(labelSource)
        const branding = sanitizeBrandingConfig(entry.branding)
        const pluginId = typeof entry.pluginId === 'string' && entry.pluginId.trim() !== '' ? entry.pluginId.trim() : key
        const entryTags = []
        if (Array.isArray(entry.tags)) {
            entry.tags.forEach((tag) => {
                if (typeof tag === 'string') {
                    const trimmed = tag.trim()
                    if (trimmed) {
                        entryTags.push(trimmed)
                    }
                }
            })
        } else if (typeof entry.tags === 'string') {
            entry.tags.split(',').forEach((tag) => {
                const trimmed = tag.trim()
                if (trimmed) {
                    entryTags.push(trimmed)
                }
            })
        }

        normalized.push({
            key,
            label: labelSource,
            agentId,
            projectId,
            endpoint,
            apiUrl,
            description,
            badge,
            branding,
            pluginId,
            tags: entryTags,
            isDefault: entry.default === true
        })
    })

    return normalized
}

function filterAgentOptionsByTags(options, tagFilters) {
    if (!Array.isArray(options) || options.length === 0) {
        return []
    }
    if (!Array.isArray(tagFilters) || tagFilters.length === 0) {
        return options
    }

    const normalizedFilters = tagFilters
        .map((tag) => (typeof tag === 'string' ? tag.trim().toLowerCase() : ''))
        .filter(Boolean)

    if (normalizedFilters.length === 0) {
        return options
    }

    return options.filter((option) => {
        if (!Array.isArray(option.tags) || option.tags.length === 0) {
            return false
        }
        const optionTags = option.tags
            .map((tag) => (typeof tag === 'string' ? tag.trim().toLowerCase() : ''))
            .filter(Boolean)

        if (optionTags.length === 0) {
            return false
        }

        return normalizedFilters.every((filterTag) => optionTags.includes(filterTag))
    })
}

function initializeAgentSelectorElements() {
    if (agentSelectorState.initialized) {
        return agentSelectorElements.container !== null
    }

    const container = document.getElementById('preSessionAgentCarousel')
    const carouselTrack = document.getElementById('carouselTrack')
    const prevBtn = document.getElementById('carouselPrevBtn')
    const nextBtn = document.getElementById('carouselNextBtn')
    const indicators = document.getElementById('carouselIndicators')

    if (!container || !carouselTrack || !prevBtn || !nextBtn || !indicators) {
        return false
    }

    agentSelectorElements = {
        container,
        carouselTrack,
        prevBtn,
        nextBtn,
        indicators
    }

    prevBtn.addEventListener('click', (event) => {
        event.preventDefault()
        navigateCarousel('prev')
    })

    nextBtn.addEventListener('click', (event) => {
        event.preventDefault()
        navigateCarousel('next')
    })

    agentSelectorState.initialized = true
    return true
}

function renderAgentSelectorOptions() {
    if (!initializeAgentSelectorElements()) {
        return
    }

    const { carouselTrack, indicators } = agentSelectorElements
    carouselTrack.innerHTML = ''
    indicators.innerHTML = ''

    // Build carousel from plugin manifests directly
    const plugins = agentSelectorState.pluginMap || {}
    const pluginEntries = Object.entries(plugins)
    
    if (pluginEntries.length === 0) {
        // Fallback to old behavior if no plugins defined
        agentSelectorElements.container.hidden = true
        return
    }

    carouselState.slideCount = pluginEntries.length

    pluginEntries.forEach(([pluginId, plugin], index) => {
        // Create carousel slide
        const slide = document.createElement('div')
        slide.className = 'carousel-slide'
        slide.dataset.pluginId = pluginId
        slide.dataset.slideIndex = index

        // Get thumbnail from plugin manifest with fallbacks
        let thumbnailUrl = '/image/favicon.png' // Default fallback
        
        if (plugin.thumbnail && typeof plugin.thumbnail === 'string') {
            thumbnailUrl = resolveAssetUrl(plugin.thumbnail)
        } else if (plugin.image && typeof plugin.image === 'string') {
            thumbnailUrl = resolveAssetUrl(plugin.image)
        } else if (plugin.assets && plugin.assets.thumbnail && typeof plugin.assets.thumbnail === 'string') {
            thumbnailUrl = resolveAssetUrl(plugin.assets.thumbnail)
        } else if (plugin.branding && plugin.branding.logoUrl && typeof plugin.branding.logoUrl === 'string') {
            thumbnailUrl = resolveAssetUrl(plugin.branding.logoUrl)
        }

        // Get label from plugin manifest
        const pluginLabel = plugin.label && typeof plugin.label === 'string' 
            ? plugin.label.trim() 
            : pluginId

        // Create image element
        const img = document.createElement('img')
        img.className = 'carousel-slide__image'
        img.src = thumbnailUrl
        img.alt = pluginLabel
        img.loading = 'lazy'

        // Create label element
        const label = document.createElement('div')
        label.className = 'carousel-slide__label'
        label.textContent = pluginLabel

        // Assemble slide
        slide.appendChild(img)
        slide.appendChild(label)

        // Add click handler to select this plugin
        slide.addEventListener('click', () => {
            selectPluginById(pluginId)
        })

        carouselTrack.appendChild(slide)

        // Create indicator dot
        const indicator = document.createElement('button')
        indicator.className = 'carousel-indicator'
        indicator.type = 'button'
        indicator.setAttribute('aria-label', `Go to ${pluginLabel}`)
        indicator.dataset.slideIndex = index
        indicator.addEventListener('click', () => {
            goToSlide(index)
        })
        indicators.appendChild(indicator)
    })

    // Initialize carousel position - find the selected plugin or default to first
    const selectedPluginId = agentSelectorState.selectedPluginId || runtimeConfig?.agent?.selectedPluginId
    let selectedIndex = 0
    let initialPluginId = null
    
    if (selectedPluginId) {
        const foundIndex = pluginEntries.findIndex(([pluginId]) => pluginId === selectedPluginId)
        if (foundIndex >= 0) {
            selectedIndex = foundIndex
            initialPluginId = selectedPluginId
        }
    }
    
    if (!initialPluginId && pluginEntries.length > 0) {
        initialPluginId = pluginEntries[0][0]
    }
    
    carouselState.currentIndex = selectedIndex
    updateCarouselPosition()
    
    // Update pre-session header with initial plugin
    if (initialPluginId) {
        updatePreSessionHeader(initialPluginId)
    }

    agentSelectorElements.container.hidden = false
}

function navigateCarousel(direction) {
    if (carouselState.isTransitioning || carouselState.slideCount === 0) {
        return
    }

    carouselState.isTransitioning = true

    if (direction === 'next') {
        carouselState.currentIndex = (carouselState.currentIndex + 1) % carouselState.slideCount
    } else if (direction === 'prev') {
        carouselState.currentIndex = (carouselState.currentIndex - 1 + carouselState.slideCount) % carouselState.slideCount
    }

    updateCarouselPosition()

    setTimeout(() => {
        carouselState.isTransitioning = false
    }, 300)
}

function goToSlide(index) {
    if (carouselState.isTransitioning || index < 0 || index >= carouselState.slideCount) {
        return
    }

    carouselState.isTransitioning = true
    carouselState.currentIndex = index
    updateCarouselPosition()

    setTimeout(() => {
        carouselState.isTransitioning = false
    }, 300)
}

function updateCarouselPosition() {
    const { carouselTrack, prevBtn, nextBtn, indicators } = agentSelectorElements

    if (!carouselTrack) {
        return
    }

    // Update track position with transform
    // Since slides are flex-basis 100%, we translate by index * 100%
    const offset = -carouselState.currentIndex * 100
    carouselTrack.style.transform = `translateX(${offset}%)`

    // Update slide active states
    const slides = carouselTrack.querySelectorAll('.carousel-slide')
    slides.forEach((slide, index) => {
        slide.classList.toggle('carousel-slide--active', index === carouselState.currentIndex)
    })

    // Update navigation buttons visibility
    if (carouselState.slideCount > 1) {
        prevBtn.hidden = false
        nextBtn.hidden = false
        indicators.hidden = false
    } else {
        prevBtn.hidden = true
        nextBtn.hidden = true
        indicators.hidden = true
    }

    // Update indicator dots
    const indicatorButtons = indicators.querySelectorAll('.carousel-indicator')
    indicatorButtons.forEach((indicator, index) => {
        indicator.classList.toggle('carousel-indicator--active', index === carouselState.currentIndex)
    })

    // Auto-select the current slide's plugin
    const plugins = agentSelectorState.pluginMap || {}
    const pluginEntries = Object.entries(plugins)
    
    if (carouselState.currentIndex >= 0 && carouselState.currentIndex < pluginEntries.length) {
        const [pluginId] = pluginEntries[carouselState.currentIndex]
        if (pluginId && pluginId !== agentSelectorState.selectedPluginId) {
            selectPluginById(pluginId)
        }
    }
}

function updatePreSessionHeader(pluginId) {
    const headerElement = document.getElementById('preSessionHeader')
    const logoElement = document.getElementById('preSessionHeaderLogo')
    const labelElement = document.getElementById('preSessionHeaderLabel')
    
    if (!headerElement || !logoElement || !labelElement) {
        return
    }

    const plugin = agentSelectorState.pluginMap ? agentSelectorState.pluginMap[pluginId] : null
    if (!plugin) {
        // Hide header if no plugin
        headerElement.hidden = true
        return
    }

    // Show header
    headerElement.hidden = false

    // Update logo
    let logoUrl = '/image/logo.png'
    if (plugin.branding && plugin.branding.logoUrl) {
        logoUrl = resolveAssetUrl(plugin.branding.logoUrl)
    } else if (plugin.thumbnail) {
        logoUrl = resolveAssetUrl(plugin.thumbnail)
    }
    logoElement.src = logoUrl
    logoElement.alt = plugin.label || pluginId

    // Update label
    labelElement.textContent = plugin.label || pluginId

    // Update header background color if branding specifies it
    if (plugin.branding && plugin.branding.primaryColor) {
        headerElement.style.backgroundColor = plugin.branding.primaryColor
    } else {
        headerElement.style.backgroundColor = ''
    }
}

function updateAgentSelectorSummary() {
    // This function is no longer needed with carousel, but keeping stub for compatibility
    return
}

function openAgentSelectorPanel() {
    // This function is no longer needed with carousel, but keeping stub for compatibility
    return
}

function closeAgentSelectorPanel() {
    // This function is no longer needed with carousel, but keeping stub for compatibility
    return
}

function toggleAgentSelectorPanel() {
    // This function is no longer needed with carousel, but keeping stub for compatibility
    return
}

function applyAgentOptionSelection(option) {
    if (!option) {
        return
    }

    if (!runtimeConfig || typeof runtimeConfig !== 'object') {
        runtimeConfig = {}
    }

    if (!runtimeConfigDefaults) {
        captureRuntimeConfigDefaults(runtimeConfig || {}, agentSelectorState.baseBranding)
    }

    const pluginId = option.pluginId || option.key || option.agentId || null
    const plugin = pluginId && agentSelectorState.pluginMap ? agentSelectorState.pluginMap[pluginId] : null
    
    // Check if we're switching to a different plugin/agent
    const previousPluginId = runtimeConfig?.agent?.activePluginId || runtimeConfig?.agent?.selectedPluginId
    const isPluginSwitch = previousPluginId && previousPluginId !== pluginId

    // Capture current connection settings BEFORE updating runtimeConfig
    const previousConnectionSettings = isPluginSwitch ? {
        provider: normalizeAgentProvider(runtimeConfig?.agent?.provider || ''),
        baseEndpoint: runtimeConfig?.agent?.endpoint,
        projectId: runtimeConfig?.agent?.projectId,
        agentApiUrl: runtimeConfig?.agent?.apiUrl,
        pluginId: previousPluginId
    } : null

    runtimeConfig.agent = cloneDeep(runtimeConfigDefaults.agent) || {}
    if (agentSelectorState.pluginMap) {
        runtimeConfig.agent.plugins = cloneDeep(agentSelectorState.pluginMap) || {}
    } else if (runtimeConfig.agent.plugins === undefined) {
        runtimeConfig.agent.plugins = {}
    }
    if (Array.isArray(runtimeConfigDefaults.agent?.options)) {
        runtimeConfig.agent.options = cloneDeep(runtimeConfigDefaults.agent.options) || []
    }

    const connection = plugin && plugin.connection ? plugin.connection : {}
    
    // If switching agents and we have an active thread, reset it using previous settings
    if (isPluginSwitch && currentThreadId && previousConnectionSettings) {
        console.log(`[AgentSwitch] Switching from ${previousPluginId} to ${pluginId}, resetting thread`)
        resetAgentThread(previousConnectionSettings).catch((err) => {
            console.warn('Failed to reset agent thread during plugin switch', err)
        })
    }

    const resolvedAgentId = option.agentId || connection.agentId || runtimeConfig.agent.agentId
    if (resolvedAgentId) {
        runtimeConfig.agent.agentId = resolvedAgentId
    }

    const resolvedProjectId = option.projectId || connection.projectId || runtimeConfig.agent.projectId
    if (resolvedProjectId) {
        runtimeConfig.agent.projectId = resolvedProjectId
    }

    const resolvedEndpoint = option.endpoint || connection.endpoint || runtimeConfig.agent.endpoint
    if (resolvedEndpoint) {
        runtimeConfig.agent.endpoint = resolvedEndpoint
    }

    const resolvedApiUrl = option.apiUrl || connection.apiUrl || runtimeConfig.agent.apiUrl
    if (resolvedApiUrl) {
        runtimeConfig.agent.apiUrl = resolvedApiUrl
    }

    let systemPrompt = undefined
    if (connection && Object.prototype.hasOwnProperty.call(connection, 'systemPrompt')) {
        systemPrompt = connection.systemPrompt
    } else if (runtimeConfigDefaults.agent && Object.prototype.hasOwnProperty.call(runtimeConfigDefaults.agent, 'systemPrompt')) {
        systemPrompt = runtimeConfigDefaults.agent.systemPrompt
    } else if (Object.prototype.hasOwnProperty.call(runtimeConfig.agent, 'systemPrompt')) {
        systemPrompt = runtimeConfig.agent.systemPrompt
    }
    if (systemPrompt !== undefined) {
        runtimeConfig.agent.systemPrompt = systemPrompt
    } else if (runtimeConfig.agent && Object.prototype.hasOwnProperty.call(runtimeConfig.agent, 'systemPrompt')) {
        delete runtimeConfig.agent.systemPrompt
    }

    runtimeConfig.agent.selectedPluginId = pluginId
    runtimeConfig.agent.activePluginId = plugin ? plugin.id : pluginId

    const resolvedProvider = normalizeAgentProvider(
        (connection && connection.provider)
        || (plugin && plugin.provider)
        || (runtimeConfigDefaults.agent && runtimeConfigDefaults.agent.provider)
        || runtimeConfig.agent.provider
        || ''
    )
    runtimeConfig.agent.provider = resolvedProvider || 'azure_ai_foundry'

    if (runtimeConfig.agent.provider === 'copilot_studio') {
        const defaultDirectLine = cloneDeep(runtimeConfigDefaults.agent?.directLine) || {}
        const pluginDirectLine = connection && typeof connection.directLine === 'object'
            ? cloneDeep(connection.directLine)
            : {}

        const directLineConfig = {
            ...defaultDirectLine,
            ...pluginDirectLine
        }

        const applyDirectLineValue = (key, value) => {
            if (value === undefined || value === null) {
                return
            }
            const stringValue = typeof value === 'string' ? value.trim() : value
            if (stringValue === '' || stringValue === null) {
                return
            }
            directLineConfig[key] = stringValue
        }

        applyDirectLineValue('endpoint', connection.directLineEndpoint || connection.endpoint || directLineConfig.endpoint)
        applyDirectLineValue('secret', connection.directLineSecret ?? directLineConfig.secret)
        applyDirectLineValue('secretEnv', connection.directLineSecretEnv ?? directLineConfig.secretEnv)
        applyDirectLineValue('botId', connection.directLineBotId || connection.agentId || directLineConfig.botId)
        applyDirectLineValue('userId', connection.directLineUserId ?? directLineConfig.userId)
        applyDirectLineValue('scope', connection.directLineScope ?? directLineConfig.scope)
        applyDirectLineValue('region', connection.directLineRegion ?? directLineConfig.region)

        runtimeConfig.agent.directLine = directLineConfig
        runtimeConfig.agent.endpoint = directLineConfig.endpoint || ''
        runtimeConfig.agent.projectId = ''
        runtimeConfig.agent.agentId = ''
    } else {
        if (runtimeConfigDefaults.agent?.directLine) {
            runtimeConfig.agent.directLine = cloneDeep(runtimeConfigDefaults.agent.directLine) || {}
        } else if (runtimeConfig.agent.directLine) {
            runtimeConfig.agent.directLine = {}
        }
    }

    if (Array.isArray(plugin?.languages) && plugin.languages.length > 0) {
        runtimeConfig.agent.languages = plugin.languages.slice()
    } else if (Array.isArray(runtimeConfigDefaults.agent?.languages)) {
        runtimeConfig.agent.languages = runtimeConfigDefaults.agent.languages.slice()
    }

    if (plugin?.assets) {
        runtimeConfig.agent.assets = cloneDeep(plugin.assets) || {}
    } else if (runtimeConfigDefaults.agent && runtimeConfigDefaults.agent.assets) {
        runtimeConfig.agent.assets = cloneDeep(runtimeConfigDefaults.agent.assets) || {}
    }

    if (plugin?.content) {
        runtimeConfig.agent.content = cloneDeep(plugin.content) || {}
    } else if (runtimeConfigDefaults.agent && runtimeConfigDefaults.agent.content) {
        runtimeConfig.agent.content = cloneDeep(runtimeConfigDefaults.agent.content) || {}
    }

    if (runtimeConfig.agent.agentId !== undefined) {
        setElementValue('azureOpenAIDeploymentName', runtimeConfig.agent.agentId)
    }
    if (runtimeConfig.agent.projectId !== undefined) {
        setElementValue('azureOpenAIProjectId', runtimeConfig.agent.projectId)
    }
    if (runtimeConfig.agent.endpoint !== undefined) {
        setElementValue('azureOpenAIEndpoint', runtimeConfig.agent.endpoint)
    }
    if (runtimeConfig.agent.apiUrl !== undefined) {
        setElementValue('agentApiUrl', runtimeConfig.agent.apiUrl)
    }
    const promptValue = runtimeConfig.agent.systemPrompt !== undefined ? runtimeConfig.agent.systemPrompt : ''
    setElementValue('prompt', promptValue)

    const baseBranding = cloneBrandingConfig(agentSelectorState.baseBranding) || cloneBrandingConfig(runtimeConfigDefaults.branding)
    const pluginBranding = sanitizeBrandingConfig(plugin?.branding)
    const resolvedBranding = mergeBranding(baseBranding, pluginBranding) || baseBranding
    runtimeConfig.branding = cloneBrandingConfig(resolvedBranding)
    runtimeConfig.currentBranding = cloneBrandingConfig(resolvedBranding)

    runtimeConfig.avatar = cloneDeep(runtimeConfigDefaults.avatar) || {}
    if (plugin?.avatar) {
        runtimeConfig.avatar = runtimeConfig.avatar || {}
        deepMergeInto(runtimeConfig.avatar, plugin.avatar)
    }

    runtimeConfig.features = cloneDeep(runtimeConfigDefaults.features) || {}
    if (plugin?.features) {
        runtimeConfig.features = runtimeConfig.features || {}
        deepMergeInto(runtimeConfig.features, plugin.features)
    }

    runtimeConfig.ui = cloneDeep(runtimeConfigDefaults.ui) || {}
    if (plugin?.ui) {
        runtimeConfig.ui = runtimeConfig.ui || {}
        deepMergeInto(runtimeConfig.ui, plugin.ui)
    }

    runtimeConfig.conversation = cloneDeep(runtimeConfigDefaults.conversation) || {}
    if (plugin?.conversation) {
        runtimeConfig.conversation = runtimeConfig.conversation || {}
        deepMergeInto(runtimeConfig.conversation, plugin.conversation)
    }

    runtimeConfig.speech = cloneDeep(runtimeConfigDefaults.speech) || {}
    if (plugin?.speech) {
        runtimeConfig.speech = runtimeConfig.speech || {}
        deepMergeInto(runtimeConfig.speech, plugin.speech)
    }

    const manifestVoice = typeof plugin?.avatar?.voice === 'string' ? plugin.avatar.voice.trim() : ''
    const manifestSpeechVoice = typeof plugin?.speech?.ttsVoice === 'string' ? plugin.speech.ttsVoice.trim() : ''
    const resolvedVoice = manifestVoice || manifestSpeechVoice || runtimeConfig.speech?.ttsVoice
    if (resolvedVoice) {
        runtimeConfig.speech = runtimeConfig.speech || {}
        runtimeConfig.speech.ttsVoice = resolvedVoice
    }

    const manifestSttLocales = typeof plugin?.speech?.sttLocales === 'string' ? plugin.speech.sttLocales.trim() : ''
    const resolvedSttLocales = manifestSttLocales || runtimeConfig.speech?.sttLocales
    if (resolvedSttLocales) {
        runtimeConfig.speech = runtimeConfig.speech || {}
        runtimeConfig.speech.sttLocales = resolvedSttLocales
    }

    if (runtimeConfig.speech) {
        if (runtimeConfig.speech.ttsVoice !== undefined) {
            setElementValue('ttsVoice', runtimeConfig.speech.ttsVoice)
        }
        if (runtimeConfig.speech.customVoiceEndpointId !== undefined) {
            setElementValue('customVoiceEndpointId', runtimeConfig.speech.customVoiceEndpointId)
        }
        if (runtimeConfig.speech.sttLocales !== undefined) {
            setElementValue('sttLocales', runtimeConfig.speech.sttLocales)
        }
    }

    if (runtimeConfig.avatar) {
        if (runtimeConfig.avatar.character !== undefined) {
            setElementValue('talkingAvatarCharacter', runtimeConfig.avatar.character)
        }
        if (runtimeConfig.avatar.style !== undefined) {
            setElementValue('talkingAvatarStyle', runtimeConfig.avatar.style)
        }
        const avatarChanged = setCheckboxValue('customizedAvatar', runtimeConfig.avatar.customized)
        const builtInVoiceChanged = setCheckboxValue('useBuiltInVoice', runtimeConfig.avatar.useBuiltInVoice)
        setCheckboxValue('autoReconnectAvatar', runtimeConfig.avatar.autoReconnect)
        const localIdleChanged = setCheckboxValue('useLocalVideoForIdle', runtimeConfig.avatar.useLocalVideoForIdle)
        if (avatarChanged || builtInVoiceChanged) {
            invokeOptionalCallback('updateCustomAvatarBox')
        }
        if (localIdleChanged) {
            invokeOptionalCallback('updateLocalVideoForIdle')
        }
    }

    if (runtimeConfig.ui) {
        setCheckboxValue('showSubtitles', runtimeConfig.ui.showSubtitles)
    }

    if (runtimeConfig.conversation) {
        setCheckboxValue('continuousConversation', runtimeConfig.conversation.continuous)
    }

    applyFeatureSettings()
    applyBrandingSettings(runtimeConfig.currentBranding)
    applyAvatarBackground(getAvatarBackgroundImage())
    applyPreSessionVideo(getPreSessionVideoUrl())

    activateAgentPlugin(plugin).catch((err) => {
        console.error('Failed to activate agent plugin', pluginId || option.pluginId || option.key, err)
    })
}

function selectPluginById(pluginId) {
    if (!pluginId || !initializeAgentSelectorElements()) {
        return false
    }

    const plugin = agentSelectorState.pluginMap ? agentSelectorState.pluginMap[pluginId] : null
    if (!plugin) {
        console.warn('Plugin not found:', pluginId)
        return false
    }

    // Create a minimal option object from the plugin
    const option = {
        key: pluginId,
        pluginId: pluginId,
        label: plugin.label || pluginId,
        agentId: plugin.connection?.agentId,
        projectId: plugin.connection?.projectId,
        endpoint: plugin.connection?.endpoint,
        apiUrl: plugin.connection?.apiUrl,
        provider: plugin.provider || plugin.connection?.provider,
        branding: plugin.branding,
        config: plugin.config
    }

    agentSelectorState.selectedPluginId = pluginId
    agentSelectorState.selectedKey = pluginId

    applyAgentOptionSelection(option)

    renderAgentSelectorOptions()
    updateAgentSelectorSummary()
    updatePreSessionHeader(pluginId)

    // Store the selected plugin ID
    if (typeof localStorage !== 'undefined') {
        try {
            localStorage.setItem('agentSelectionKey', pluginId)
        } catch (err) {
            console.warn('Failed to store agent selection', err)
        }
    }

    return true
}

function selectAgentOption(optionKey, selectionOptions) {
    if (!initializeAgentSelectorElements()) {
        return false
    }

    const option = agentSelectorState.options.find((entry) => entry.key === optionKey)
    if (!option) {
        return false
    }

    agentSelectorState.selectedKey = option.key

    applyAgentOptionSelection(option)

    renderAgentSelectorOptions()
    updateAgentSelectorSummary()

    if (selectionOptions && selectionOptions.persist) {
        setStoredAgentSelectionKey(option.key)
    }

    if (selectionOptions && selectionOptions.closePanel) {
        closeAgentSelectorPanel()
    }

    if (selectionOptions && selectionOptions.focusToggle) {
        agentSelectorElements.toggleButton.focus()
    }

    return true
}

function updateAgentSelectorFromConfig(agentConfig) {
    if (!initializeAgentSelectorElements()) {
        return
    }

    const normalized = normalizeAgentOptions(agentConfig)
    const filtered = filterAgentOptionsByTags(normalized, agentSelectorState.tagFilters)
    agentSelectorState.options = filtered
    let lockSelector = false

    const defaultPluginId = typeof agentConfig?.defaultPluginId === 'string' ? agentConfig.defaultPluginId.trim() : ''
    const selectedPluginId = typeof agentConfig?.selectedPluginId === 'string' ? agentConfig.selectedPluginId.trim() : ''

    if (!filtered.length) {
        if (agentSelectorState.tagFilters && agentSelectorState.tagFilters.length) {
            console.warn('No agents matched requested tags:', agentSelectorState.tagFilters)
        }
        agentSelectorState.selectedKey = null
        agentSelectorElements.container.hidden = true
        agentSelectorState.locked = false
        return
    }

    agentSelectorElements.container.hidden = false

    const storedKey = getStoredAgentSelectionKey()
    const candidateByStorage = storedKey ? filtered.find((option) => option.key === storedKey) : null
    const candidateBySelectedPlugin = selectedPluginId
        ? filtered.find((option) => option.pluginId === selectedPluginId || option.key === selectedPluginId)
        : null
    const candidateByPluginDefault = defaultPluginId
        ? filtered.find((option) => option.pluginId === defaultPluginId || option.key === defaultPluginId)
        : null
    const candidateByDefault = filtered.find((option) => option.isDefault)
    const queryOverride = agentSelectorState.overrideQuery
    const candidateByOverride = queryOverride
        ? filtered.find((option) => {
            const queryLower = queryOverride.trim().toLowerCase()
            if (typeof option.key === 'string' && option.key.trim().toLowerCase() === queryLower) {
                return true
            }
            if (typeof option.agentId === 'string' && option.agentId.trim().toLowerCase() === queryLower) {
                return true
            }
            if (typeof option.label === 'string' && option.label.trim() !== '') {
                return option.label.trim().toLowerCase() === queryLower
            }
            return false
        })
        : null

    const configAgentId = typeof agentConfig?.agentId === 'string' ? agentConfig.agentId.trim() : ''
    const candidateByConfig = configAgentId ? filtered.find((option) => option.agentId === configAgentId) : null

    const existingSelected = agentSelectorState.selectedKey ? filtered.find((option) => option.key === agentSelectorState.selectedKey) : null

    const fallback = filtered[0]

    const targetOption = candidateByOverride
        || candidateByStorage
        || candidateBySelectedPlugin
        || candidateByPluginDefault
        || candidateByDefault
        || candidateByConfig
        || existingSelected
        || fallback

    if (targetOption) {
        const shouldPersist = !candidateByOverride
        selectAgentOption(targetOption.key, { persist: shouldPersist, closePanel: true })
        lockSelector = Boolean(candidateByOverride)
    } else {
        renderAgentSelectorOptions()
        updateAgentSelectorSummary()
    }

    if (lockSelector) {
        agentSelectorState.locked = true
        agentSelectorElements.container.hidden = true
    } else {
        agentSelectorState.locked = false
        if (queryOverride && !candidateByOverride) {
            console.warn('Agent override query parameter did not match any configured agent:', queryOverride)
        }
    }
}

function coerceBoolean(value) {
    if (typeof value === 'boolean') {
        return value
    }
    if (typeof value === 'string') {
        const lowered = value.trim().toLowerCase()
        if (['1', 'true', 'yes', 'y', 'on'].includes(lowered)) {
            return true
        }
        if (['0', 'false', 'no', 'n', 'off'].includes(lowered)) {
            return false
        }
    }
    if (typeof value === 'number') {
        return value !== 0
    }
    return undefined
}

function setElementValue(elementId, value) {
    if (value === undefined || value === null) {
        return
    }
    const element = document.getElementById(elementId)
    if (!element) {
        return
    }
    if ('value' in element) {
        element.value = value
    }
}

function setCheckboxValue(elementId, value) {
    const element = document.getElementById(elementId)
    if (!element) {
        return false
    }
    const normalized = coerceBoolean(value)
    if (normalized === undefined) {
        return false
    }
    const changed = element.checked !== normalized
    element.checked = normalized
    return changed
}

function getRuntimeSection(sectionName) {
    if (!runtimeConfig || typeof runtimeConfig !== 'object') {
        return {}
    }
    const section = runtimeConfig[sectionName]
    if (!section || typeof section !== 'object') {
        return {}
    }
    return section
}

function getServicesProxyBaseUrl() {
    if (servicesProxyBaseUrl) {
        return servicesProxyBaseUrl
    }
    if (typeof window.SERVICES_PROXY_BASE_URL === 'string') {
        const trimmed = window.SERVICES_PROXY_BASE_URL.trim()
        if (trimmed !== '') {
            servicesProxyBaseUrl = trimmed.replace(/\/$/, '')
            return servicesProxyBaseUrl
        }
    }
    return ''
}

function buildServicesProxyCandidates(path) {
    const candidates = []
    const base = getServicesProxyBaseUrl()
    if (base) {
        candidates.push(`${base}${path}`)
    }

    candidates.push(`http://localhost:4100${path}`)
    candidates.push(`https://localhost:4100${path}`)
    candidates.push(`http://127.0.0.1:4100${path}`)
    candidates.push(`https://127.0.0.1:4100${path}`)

    if (typeof window.location === 'object') {
        try {
            const origin = window.location.origin
            if (typeof origin === 'string' && origin.trim() !== '') {
                candidates.push(`${origin.replace(/\/$/, '')}${path}`)
            }
        } catch (_) {
            // Ignore location access errors
        }
    }

    const unique = []
    const seen = new Set()
    for (const candidate of candidates) {
        if (!candidate || seen.has(candidate)) {
            continue
        }
        seen.add(candidate)
        unique.push(candidate)
    }
    return unique
}

function getRuntimeString(sectionName, key, fallbackValue) {
    const section = getRuntimeSection(sectionName)
    const value = section[key]
    if (typeof value === 'string') {
        const trimmed = value.trim()
        if (trimmed !== '') {
            return trimmed
        }
        return ''
    }
    if (value !== undefined && value !== null) {
        return String(value)
    }
    return fallbackValue ?? ''
}

function getRuntimeBoolean(sectionName, key, fallbackValue) {
    const section = getRuntimeSection(sectionName)
    const value = section[key]
    if (typeof value === 'boolean') {
        return value
    }
    return fallbackValue
}

function getDomInputValue(elementId) {
    const element = document.getElementById(elementId)
    if (!element || typeof element.value !== 'string') {
        return ''
    }
    return element.value.trim()
}

function getRuntimeArray(sectionName, key, fallbackValue = []) {
    const section = getRuntimeSection(sectionName)
    const value = section[key]
    if (Array.isArray(value)) {
        return value
    }
    if (typeof value === 'string') {
        const parts = value
            .split(/[|,]/)
            .map((part) => part.trim())
            .filter((part) => part !== '')
        if (parts.length > 0) {
            return parts
        }
    }
    return fallbackValue
}

function getRuntimeCheckbox(sectionName, key, elementId, defaultValue = false) {
    const fallback = document.getElementById(elementId)?.checked ?? defaultValue
    return getRuntimeBoolean(sectionName, key, fallback)
}

function isAutoReconnectEnabled() {
    return getRuntimeCheckbox('avatar', 'autoReconnect', 'autoReconnectAvatar', false)
}

function isLocalIdleVideoEnabled() {
    return getRuntimeCheckbox('avatar', 'useLocalVideoForIdle', 'useLocalVideoForIdle', false)
}

function isSubtitlesEnabled() {
    return getRuntimeCheckbox('ui', 'showSubtitles', 'showSubtitles', false)
}

function isContinuousConversationEnabled() {
    return getRuntimeCheckbox('conversation', 'continuous', 'continuousConversation', false)
}

function applyRuntimeConfig(config) {
    runtimeConfig = config
    if (!config || typeof config !== 'object') {
        return
    }

    if (config.agent && typeof config.agent.plugins === 'object') {
        const clonedPlugins = cloneDeep(config.agent.plugins)
        agentSelectorState.pluginMap = clonedPlugins || { ...config.agent.plugins }
        config.agent.plugins = clonedPlugins || { ...config.agent.plugins }
    } else {
        agentSelectorState.pluginMap = {}
        if (config.agent) {
            delete config.agent.plugins
        }
    }

    const baseBranding = sanitizeBrandingConfig(config.branding)
    agentSelectorState.baseBranding = cloneBrandingConfig(baseBranding)
    runtimeConfig.branding = cloneBrandingConfig(baseBranding)
    runtimeConfig.currentBranding = cloneBrandingConfig(baseBranding)

    captureRuntimeConfigDefaults(config, baseBranding)

    if (typeof config.servicesProxyBaseUrl === 'string' && config.servicesProxyBaseUrl.trim() !== '') {
        servicesProxyBaseUrl = config.servicesProxyBaseUrl.trim().replace(/\/$/, '')
    }

    if (config.speech) {
        setElementValue('region', config.speech.region)
        setElementValue('APIKey', config.speech.apiKey)
        setCheckboxValue('enablePrivateEndpoint', config.speech.enablePrivateEndpoint)
        setElementValue('privateEndpoint', config.speech.privateEndpoint)
        setElementValue('sttLocales', config.speech.sttLocales)
        setElementValue('ttsVoice', config.speech.ttsVoice)
        setElementValue('customVoiceEndpointId', config.speech.customVoiceEndpointId)
        invokeOptionalCallback('updatePrivateEndpoint')
    }

    const pendingAgentConfig = config.agent && typeof config.agent === 'object' ? config.agent : null
    if (pendingAgentConfig) {
        setElementValue('azureOpenAIEndpoint', pendingAgentConfig.endpoint)
        setElementValue('azureOpenAIDeploymentName', pendingAgentConfig.agentId)
        setElementValue('azureOpenAIProjectId', pendingAgentConfig.projectId)
        setElementValue('agentApiUrl', pendingAgentConfig.apiUrl)
        setElementValue('prompt', pendingAgentConfig.systemPrompt ?? '')
    }

    if (config.search) {
        setCheckboxValue('enableOyd', config.search.enabled)
        setElementValue('azureCogSearchEndpoint', config.search.endpoint)
        setElementValue('azureCogSearchApiKey', config.search.apiKey)
        setElementValue('azureCogSearchIndexName', config.search.indexName)
        invokeOptionalCallback('updataEnableOyd')
    }

    if (config.avatar) {
        setElementValue('talkingAvatarCharacter', config.avatar.character)
        setElementValue('talkingAvatarStyle', config.avatar.style)
        const avatarChanged = setCheckboxValue('customizedAvatar', config.avatar.customized)
        const builtInVoiceChanged = setCheckboxValue('useBuiltInVoice', config.avatar.useBuiltInVoice)
        setCheckboxValue('autoReconnectAvatar', config.avatar.autoReconnect)
        const localIdleChanged = setCheckboxValue('useLocalVideoForIdle', config.avatar.useLocalVideoForIdle)
        if (avatarChanged || builtInVoiceChanged) {
            invokeOptionalCallback('updateCustomAvatarBox')
        }
        if (localIdleChanged) {
            invokeOptionalCallback('updateLocalVideoForIdle')
        }
    }

    if (config.ui) {
        setCheckboxValue('showSubtitles', config.ui.showSubtitles)
    }

    if (config.conversation) {
        setCheckboxValue('continuousConversation', config.conversation.continuous)
    }

    if (pendingAgentConfig) {
        updateAgentSelectorFromConfig(pendingAgentConfig)
    }

    applyFeatureSettings()

    applyBrandingSettings(runtimeConfig.currentBranding)
    applyAvatarBackground(getAvatarBackgroundImage())
    applyPreSessionVideo(getPreSessionVideoUrl())
}

async function loadRuntimeConfig() {
    setStartSessionButtonState({ disabled: true, busy: true, label: 'Loading configuration...' })
    const candidates = []

    if (typeof window.RUNTIME_CONFIG_URL === 'string' && window.RUNTIME_CONFIG_URL.trim() !== '') {
        candidates.push(window.RUNTIME_CONFIG_URL.trim())
    }

    const proxyBase = getServicesProxyBaseUrl()
    if (proxyBase) {
        candidates.push(`${proxyBase}/config`)
    }

    candidates.push('/config')
    candidates.push('http://localhost:8080/config')
    candidates.push('https://localhost:8080/config')
    candidates.push('http://localhost:4100/config')
    candidates.push('https://localhost:4100/config')
    candidates.push('http://localhost:4000/config')
    candidates.push('https://localhost:4000/config')

    const errors = []

    for (const url of candidates) {
        try {
            const response = await fetch(url, { cache: 'no-store' })
            if (!response.ok) {
                throw new Error(`Failed to load runtime configuration: ${response.status}`)
            }
            const config = await response.json()
            setRuntimeConfigAssetBase(response.url || url)
            applyRuntimeConfig(config)
            setStartSessionButtonState({ disabled: false, busy: false, label: null })
            return
        } catch (err) {
            errors.push({ url, error: err })
            console.warn(`Unable to load runtime configuration from ${url}`, err)
        }
    }

    console.error('Failed to load runtime configuration from all candidates', errors)
    setStartSessionButtonState({ disabled: true, busy: false, label: 'Configuration unavailable' })
}

document.addEventListener('DOMContentLoaded', () => {
    initializeAgentOverrideFromQuery()
    loadRuntimeConfig().catch((err) => {
        console.error('Unexpected error while loading runtime configuration', err)
        setStartSessionButtonState({ disabled: true, busy: false, label: 'Configuration unavailable' })
    })
})

async function requestSpeechSession() {
    const candidates = buildServicesProxyCandidates('/speech/token')
    const errors = []
    for (const endpoint of candidates) {
        try {
            const response = await fetch(endpoint, { method: 'POST' })
            if (!response.ok) {
                const text = await response.text()
                errors.push(`${endpoint} → ${response.status}`)
                continue
            }
            return await response.json()
        } catch (err) {
            errors.push(`${endpoint} → ${err.message ?? err}`)
        }
    }
    throw new Error(`Unable to obtain speech token. Attempts: ${errors.join('; ')}`)
}

// Define remoteVideoDiv globally
var remoteVideoDiv;

function formatAIProjectsError(err) {
    if (!err) {
        return 'Unknown error'
    }

    const parsed = err.response?.parsedBody || err.response?.body
    if (typeof parsed === 'string' && parsed.trim() !== '') {
        return parsed
    }

    if (parsed && typeof parsed === 'object') {
        if (typeof parsed.error === 'string' && parsed.error.trim() !== '') {
            return parsed.error
        }
        if (typeof parsed.message === 'string' && parsed.message.trim() !== '') {
            return parsed.message
        }
        if (parsed.error?.message) {
            return parsed.error.message
        }
        if (typeof parsed.error?.error === 'string' && parsed.error.error.trim() !== '') {
            return parsed.error.error
        }
        try {
            return JSON.stringify(parsed)
        } catch (jsonErr) {
            console.warn('Unable to stringify parsed error body', jsonErr)
        }
    }

    if (err.message) {
        return err.message
    }

    try {
        return JSON.stringify(err)
    } catch (jsonErr) {
        console.warn('Unable to stringify error object', jsonErr)
    }

    return String(err)
}

function buildAgentApiUrl(agentApiUrl, path, query) {
    const normalizedBase = agentApiUrl.endsWith('/') ? agentApiUrl : `${agentApiUrl}/`
    const url = new URL(path.replace(/^\//, ''), normalizedBase)
    if (query) {
        Object.entries(query).forEach(([key, value]) => {
            if (value !== undefined && value !== null && value !== '') {
                url.searchParams.set(key, value)
            }
        })
    }
    return url.toString()
}

function normalizeAgentProvider(value) {
    if (typeof value !== 'string') {
        return 'azure_ai_foundry'
    }
    const trimmed = value.trim().toLowerCase()
    if (!trimmed) {
        return 'azure_ai_foundry'
    }
    if (trimmed === 'copilot_studio' || trimmed === 'copilotstudio' || trimmed === 'copilot' || trimmed === 'directline' || trimmed === 'direct-line' || trimmed === 'direct_line') {
        return 'copilot_studio'
    }
    if (trimmed === 'azure' || trimmed === 'azure-openai' || trimmed === 'azure_ai_foundry') {
        return 'azure_ai_foundry'
    }
    return trimmed
}

function readAgentConfigValues() {
    const providerRaw = getRuntimeString('agent', 'provider', runtimeConfig?.agent?.provider || '')
    const provider = normalizeAgentProvider(providerRaw)
    const pluginId = runtimeConfig?.agent?.activePluginId
        || runtimeConfig?.agent?.selectedPluginId
        || runtimeConfig?.agent?.defaultPluginId
        || ''
    const directLineConfig = runtimeConfig?.agent?.directLine && typeof runtimeConfig.agent.directLine === 'object'
        ? cloneDeep(runtimeConfig.agent.directLine)
        : {}

    return {
        provider,
        baseEndpoint: getRuntimeString('agent', 'endpoint', getDomInputValue('azureOpenAIEndpoint')),
        agentId: getRuntimeString('agent', 'agentId', getDomInputValue('azureOpenAIDeploymentName')),
        projectId: getRuntimeString('agent', 'projectId', getDomInputValue('azureOpenAIProjectId')),
        agentApiUrl: getRuntimeString('agent', 'apiUrl', getDomInputValue('agentApiUrl')),
        pluginId: typeof pluginId === 'string' ? pluginId.trim() : '',
        directLine: directLineConfig || {}
    }
}

function getAgentConfig() {
    const {
        provider,
        baseEndpoint,
        agentId,
        projectId,
        agentApiUrl,
        pluginId,
        directLine
    } = readAgentConfigValues()

    if (!agentApiUrl) {
        alert('Please provide the agent API URL so the app can communicate with the backend.')
        return null
    }

    if (provider === 'copilot_studio') {
        return {
            provider,
            agentApiUrl,
            pluginId,
            directLine: directLine || {}
        }
    }

    if (!baseEndpoint || !agentId || !projectId) {
        alert('Please fill in the Azure AI Foundry endpoint, project ID, and agent ID.')
        return null
    }

    return {
        provider,
        baseEndpoint,
        agentId,
        projectId,
        agentApiUrl,
        pluginId,
        directLine: directLine || {}
    }
}

function getAgentConfiguration() {
    return getAgentConfig()
}

function tryGetAgentConnectionSettings() {
    const { provider, baseEndpoint, projectId, agentApiUrl, pluginId } = readAgentConfigValues()
    if (!agentApiUrl) {
        return null
    }
    if (provider === 'copilot_studio') {
        return {
            provider,
            agentApiUrl,
            pluginId
        }
    }
    if (!baseEndpoint || !projectId) {
        return null
    }
    return {
        provider,
        baseEndpoint,
        projectId,
        agentApiUrl,
        pluginId
    }
}

async function ensureAgentThread(connection) {
    const agentApiUrl = connection?.agentApiUrl || readAgentConfigValues().agentApiUrl
    const provider = normalizeAgentProvider(connection?.provider || runtimeConfig?.agent?.provider || '')
    const pluginId = connection?.pluginId
        || runtimeConfig?.agent?.activePluginId
        || runtimeConfig?.agent?.selectedPluginId
        || runtimeConfig?.agent?.defaultPluginId
        || ''
    const baseEndpoint = connection?.baseEndpoint
    const projectId = connection?.projectId
    
    // If we have a thread ID, verify it's for the current plugin/provider
    if (currentThreadId) {
        // Check if the existing thread is for the same plugin
        const currentPluginId = runtimeConfig?.agent?.activePluginId || runtimeConfig?.agent?.selectedPluginId
        const currentProvider = normalizeAgentProvider(runtimeConfig?.agent?.provider || '')
        
        // If plugin or provider changed, reset the thread first
        if (currentPluginId && pluginId && currentPluginId !== pluginId) {
            console.log(`[AgentThread] Plugin mismatch (${currentPluginId} → ${pluginId}), resetting thread`)
            await resetAgentThread()
        } else if (currentProvider && provider && currentProvider !== provider) {
            console.log(`[AgentThread] Provider mismatch (${currentProvider} → ${provider}), resetting thread`)
            await resetAgentThread()
        } else {
            // Thread is valid for current plugin/provider
            return true
        }
    }

    try {
        const requestBody = { provider }
        if (pluginId) {
            requestBody.pluginId = pluginId
        }

        if (provider !== 'copilot_studio') {
            if (!baseEndpoint || !projectId) {
                throw new Error('Azure AI Foundry endpoint or project ID is missing.')
            }
            requestBody.endpoint = baseEndpoint
            requestBody.projectId = projectId
        }

        const threadResponse = await agentApiRequest(agentApiUrl, 'POST', '/thread', requestBody)
        if (!threadResponse?.id) {
            throw new Error('Thread creation returned no identifier.')
        }

        currentThreadId = threadResponse.id
        spokenAssistantMessageIds = new Set()
        assistantMessageBuffers.clear()
        lastAssistantMessageKey = null
        console.log('Initialized agent thread:', currentThreadId, 'for plugin:', pluginId, 'provider:', provider)
        return true
    } catch (err) {
        console.error('Failed to initialize agent thread', err)
        const formattedError = provider === 'copilot_studio'
            ? (err?.message || 'Unable to create Copilot Studio conversation.')
            : formatAIProjectsError(err)
        alert(`Failed to create agent thread: ${formattedError}`)
        return false
    }
}

async function agentApiRequest(agentApiUrl, method, path, body, query) {
    const url = buildAgentApiUrl(agentApiUrl, path, query)
    const options = {
        method,
        headers: { 'Content-Type': 'application/json' }
    }
    if (body !== undefined) {
        options.body = JSON.stringify(body)
    }

    const response = await fetch(url, options)
    if (!response.ok) {
        let errorPayload = null
        try {
            errorPayload = await response.json()
        } catch (jsonErr) {
            try {
                errorPayload = await response.text()
            } catch (_) {
                errorPayload = null
            }
        }
        const error = new Error(typeof errorPayload === 'object' && errorPayload?.error ? errorPayload.error : response.statusText)
        error.response = {
            status: response.status,
            parsedBody: errorPayload
        }
        throw error
    }

    if (response.status === 204) {
        return null
    }

    try {
        return await response.json()
    } catch (jsonErr) {
        return null
    }
}

async function resetAgentThread(previousConnection = null) {
    const threadId = currentThreadId
    currentThreadId = null
    spokenAssistantMessageIds = new Set()
    assistantMessageBuffers.clear()
    lastAssistantMessageKey = null

    // Use provided previous connection settings, or try to get current ones
    const connection = previousConnection || tryGetAgentConnectionSettings()
    if (!threadId || !connection) {
        if (!connection) {
            console.warn('Unable to reset agent thread because connection settings are unavailable.')
        }
        return
    }

    try {
        const deletePayload = {
            provider: normalizeAgentProvider(connection.provider || runtimeConfig?.agent?.provider || ''),
            pluginId: connection.pluginId
        }
        if (deletePayload.provider !== 'copilot_studio') {
            deletePayload.endpoint = connection.baseEndpoint
            deletePayload.projectId = connection.projectId
        }
        await agentApiRequest(connection.agentApiUrl, 'DELETE', `/thread/${encodeURIComponent(threadId)}`, deletePayload)
        console.log('Deleted agent thread:', threadId, 'for plugin:', connection.pluginId)
    } catch (err) {
        console.warn('Failed to delete agent thread', err)
    }

    // Only call ensureAgentThread if we're not switching plugins (previousConnection not provided)
    if (!previousConnection) {
        await ensureAgentThread(connection)
    }
}

// Connect to avatar service
async function connectAvatar() {
    const speechRegion = getRuntimeString('speech', 'region', getDomInputValue('region')) || 'eastus2'
    const privateEndpointEnabled = getRuntimeBoolean('speech', 'enablePrivateEndpoint', document.getElementById('enablePrivateEndpoint')?.checked ?? false)
    const privateEndpointRaw = getRuntimeString('speech', 'privateEndpoint', getDomInputValue('privateEndpoint'))
    const privateEndpoint = privateEndpointRaw.startsWith('https://') ? privateEndpointRaw.slice(8) : privateEndpointRaw
    if (privateEndpointEnabled && privateEndpoint === '') {
        alert('Please fill in the Azure Speech endpoint.')
        return false
    }

    try {
        const session = await requestSpeechSession()
        const speechToken = session.speechToken
        if (!speechToken) {
            throw new Error('Speech token is missing from services proxy response.')
        }

        const effectiveRegion = (session.region && session.region.trim()) || speechRegion
        const useManagedIdentity = Boolean(session.useManagedIdentity)
        const speechEndpoint = (session.speechEndpoint ?? '').trim()

        const speechSynthesisConfig = SpeechSDK.SpeechConfig.fromAuthorizationToken(speechToken, effectiveRegion)

        // When using Managed Identity, the custom domain endpoint is required for
        // Entra ID (AAD) token authentication. Regional endpoints don't support AAD tokens.
        let ttsEndpoint
        if (useManagedIdentity && speechEndpoint) {
            const normalizedEndpoint = speechEndpoint.replace(/^https?:\/\//, '')
            ttsEndpoint = `wss://${normalizedEndpoint}/tts/cognitiveservices/websocket/v1?enableTalkingAvatar=true`
        } else if (privateEndpointEnabled) {
            ttsEndpoint = `wss://${privateEndpoint}/tts/cognitiveservices/websocket/v1?enableTalkingAvatar=true`
        } else {
            ttsEndpoint = `wss://${effectiveRegion}.tts.speech.microsoft.com/cognitiveservices/websocket/v1?enableTalkingAvatar=true`
        }
        speechSynthesisConfig.setProperty(SpeechSDK.PropertyId.SpeechServiceConnection_Endpoint, ttsEndpoint)
        speechSynthesisConfig.endpointId = getRuntimeString('speech', 'customVoiceEndpointId', getDomInputValue('customVoiceEndpointId'))

        const talkingAvatarCharacter = getRuntimeString('avatar', 'character', getDomInputValue('talkingAvatarCharacter') || 'lisa') || 'lisa'
        const talkingAvatarStyle = getRuntimeString('avatar', 'style', getDomInputValue('talkingAvatarStyle') || 'casual-sitting') || 'casual-sitting'
        const transparentBackground = isTransparentBackgroundEnabled()
        const avatarConfig = new SpeechSDK.AvatarConfig(talkingAvatarCharacter, talkingAvatarStyle)
        avatarConfig.customized = getRuntimeBoolean('avatar', 'customized', document.getElementById('customizedAvatar')?.checked ?? false)
        avatarConfig.useBuiltInVoice = getRuntimeBoolean('avatar', 'useBuiltInVoice', document.getElementById('useBuiltInVoice')?.checked ?? false)
        const avatarBackgroundImage = getAvatarBackgroundImage()
        console.log('Avatar background image:', avatarBackgroundImage)
        console.log('Avatar background transparent:', transparentBackground)

        if (!transparentBackground && avatarBackgroundImage) {
            const resolvedBackgroundImage = resolveAssetUrl(avatarBackgroundImage)
            if (resolvedBackgroundImage) {
                avatarConfig.backgroundImage = resolvedBackgroundImage
            }
        }
        if (transparentBackground) {
            avatarConfig.videoCodec = 'vp9'
            avatarConfig.videoFormat = 'webm'
            avatarConfig.backgroundColor = '#FFFFFF00'
        }
        avatarSynthesizer = new SpeechSDK.AvatarSynthesizer(speechSynthesisConfig, avatarConfig)
        avatarSynthesizer.authorizationToken = speechToken
        avatarSynthesizer.avatarEventReceived = function (s, e) {
            var offsetMessage = ", offset from session start: " + e.offset / 10000 + "ms."
            if (e.offset === 0) {
                offsetMessage = ""
            }

            console.log("Event received: " + e.description + offsetMessage)
        }

        const speechRecognitionConfig = SpeechSDK.SpeechConfig.fromAuthorizationToken(speechToken, effectiveRegion)
        let sttEndpoint
        if (useManagedIdentity && speechEndpoint) {
            const normalizedEndpoint = speechEndpoint.replace(/^https?:\/\//, '')
            sttEndpoint = `wss://${normalizedEndpoint}/stt/speech/universal/v2`
        } else if (privateEndpointEnabled) {
            sttEndpoint = `wss://${privateEndpoint}/stt/speech/universal/v2`
        } else {
            sttEndpoint = `wss://${effectiveRegion}.stt.speech.microsoft.com/speech/universal/v2`
        }
        speechRecognitionConfig.setProperty(SpeechSDK.PropertyId.SpeechServiceConnection_Endpoint, sttEndpoint)
        speechRecognitionConfig.setProperty(SpeechSDK.PropertyId.SpeechServiceConnection_LanguageIdMode, "Continuous")
        const sttLocalesSource = getRuntimeString('speech', 'sttLocales', getDomInputValue('sttLocales') || 'en-US,da-DK')
        var sttLocales = sttLocalesSource.split(',')
        var autoDetectSourceLanguageConfig = SpeechSDK.AutoDetectSourceLanguageConfig.fromLanguages(sttLocales)
        speechRecognizer = SpeechSDK.SpeechRecognizer.FromConfig(speechRecognitionConfig, autoDetectSourceLanguageConfig, SpeechSDK.AudioConfig.fromDefaultMicrophoneInput())
    speechRecognizer.authorizationToken = speechToken

        const agentThreadConfig = getAgentConfiguration()
        if (!agentThreadConfig) {
            return false
        }

        dataSources = []
        const onYourDataEnabled = getRuntimeBoolean('search', 'enabled', document.getElementById('enableOyd')?.checked ?? false)
        if (onYourDataEnabled) {
            const azureCogSearchEndpoint = getRuntimeString('search', 'endpoint', getDomInputValue('azureCogSearchEndpoint'))
            const azureCogSearchApiKey = getRuntimeString('search', 'apiKey', getDomInputValue('azureCogSearchApiKey'))
            const azureCogSearchIndexName = getRuntimeString('search', 'indexName', getDomInputValue('azureCogSearchIndexName'))
            if (azureCogSearchEndpoint === "" || azureCogSearchApiKey === "" || azureCogSearchIndexName === "") {
                alert('Please fill in the Azure Cognitive Search endpoint, API key and index name.')
                return false
            } else {
                setDataSources(azureCogSearchEndpoint, azureCogSearchApiKey, azureCogSearchIndexName)
            }
        }

        if (!messageInitiated) {
            initMessages()
            messageInitiated = true
        }

        document.getElementById('startSession').disabled = true

        const relay = session.relay || {}
        const iceServerUrl = Array.isArray(relay.Urls) ? relay.Urls[0] : undefined
        const iceServerUsername = relay.Username
        const iceServerCredential = relay.Password
        if (!iceServerUrl || !iceServerUsername || !iceServerCredential) {
            throw new Error('Services proxy did not provide valid relay credentials.')
        }

        setupWebRTC(iceServerUrl, iceServerUsername, iceServerCredential, agentThreadConfig)
        return true
    } catch (err) {
        console.error('Failed to establish avatar connection', err)
        alert(`Failed to establish avatar connection: ${err.message ?? err}`)
        return false
    }
}

// Disconnect from avatar service
function disconnectAvatar() {
    if (avatarSynthesizer !== undefined) {
        avatarSynthesizer.close()
        avatarSynthesizer = undefined
    }

    if (speechRecognizer !== undefined) {
        speechRecognizer.stopContinuousRecognitionAsync()
        speechRecognizer.close()
        speechRecognizer = undefined
    }

    sessionActive = false
}

// Setup WebRTC
function setupWebRTC(iceServerUrl, iceServerUsername, iceServerCredential, agentThreadConfig) {
    // Initialize remoteVideoDiv
    remoteVideoDiv = document.getElementById('remoteVideo');

    // Create WebRTC peer connection
    peerConnection = new RTCPeerConnection({
        iceServers: [{
            urls: [ iceServerUrl ],
            username: iceServerUsername,
            credential: iceServerCredential
        }]
    })

    // Fetch WebRTC video stream and mount it to an HTML video element
    peerConnection.ontrack = function (event) {
        if (event.track.kind === 'audio') {
            let audioElement = document.createElement('audio')
            audioElement.id = 'audioPlayer'
            audioElement.srcObject = event.streams[0]
            audioElement.autoplay = false
            audioElement.addEventListener('loadeddata', () => {
                audioElement.play()
            })

            audioElement.onplaying = () => {
                console.log(`WebRTC ${event.track.kind} channel connected.`)
            }

            // Clean up existing audio element if there is any
            for (var i = 0; i < remoteVideoDiv.childNodes.length; i++) {
                if (remoteVideoDiv.childNodes[i].localName === event.track.kind) {
                    remoteVideoDiv.removeChild(remoteVideoDiv.childNodes[i])
                }
            }

            // Append the new audio element
            document.getElementById('remoteVideo').appendChild(audioElement)
        }

        if (event.track.kind === 'video') {
            let videoElement = document.createElement('video')
            videoElement.id = 'videoPlayer'
            videoElement.srcObject = event.streams[0]
            videoElement.autoplay = false
            videoElement.addEventListener('loadeddata', () => {
                videoElement.play()
            })

            videoElement.playsInline = true
            videoElement.style.width = '0px'
            document.getElementById('remoteVideo').appendChild(videoElement)

            // Continue speaking if there are unfinished sentences
            if (repeatSpeakingSentenceAfterReconnection) {
                if (speakingText !== '') {
                    speakNext(speakingText, 0, true)
                }
            } else {
                if (spokenTextQueue.length > 0) {
                    speakNext(spokenTextQueue.shift())
                }
            }

            videoElement.onplaying = async () => {
                // Clean up existing video element if there is any
                for (var i = 0; i < remoteVideoDiv.childNodes.length; i++) {
                    if (remoteVideoDiv.childNodes[i].localName === event.track.kind) {
                        remoteVideoDiv.removeChild(remoteVideoDiv.childNodes[i])
                    }
                }
                const spinner = document.getElementById('loadingSpinner')
                if (spinner) {
                    spinner.style.visibility = 'hidden'
                }
                // Append the new video element
                videoElement.style.width = '100%'
                videoElement.style.height = '100vh'
                videoElement.style["object-fit"] = "cover";
                document.getElementById('remoteVideo').appendChild(videoElement)
                const container = document.getElementById('videoContainer')
                if (container) {
                    container.hidden = false
                }

                console.log(`WebRTC ${event.track.kind} channel connected.`)
                document.getElementById('microphone').disabled = false
                document.getElementById('stopSession').disabled = false
                document.getElementById('remoteVideo').style.height = '100vh'
                const chatHistory = document.getElementById('chatHistory')
                const toggleButton = document.getElementById('toggleChatHistory')
                if (toggleButton) {
                    const nextLabel = chatHistory && !chatHistory.hidden ? 'Hide Chat History' : 'Show Chat History'
                    updateButtonLabel(toggleButton, nextLabel)
                }
                const microphoneButton = document.getElementById('microphone')
                if (microphoneButton && microphoneButton.dataset.state !== 'active') {
                    setMicrophoneState('inactive')
                }
                ensureUserMessageInput()

                if (isLocalIdleVideoEnabled()) {
                    hideLocalIdleVideo()
                    if (lastSpeakTime === undefined) {
                        lastSpeakTime = new Date()
                    }
                }

                if (agentThreadConfig && !currentThreadId) {
                    const ensurePayload = {
                        provider: agentThreadConfig.provider,
                        agentApiUrl: agentThreadConfig.agentApiUrl,
                        baseEndpoint: agentThreadConfig.baseEndpoint,
                        projectId: agentThreadConfig.projectId,
                        pluginId: agentThreadConfig.pluginId
                    }
                    if (ensurePayload.agentApiUrl) {
                        await ensureAgentThread(ensurePayload)
                    }
                }

                isReconnecting = false
                setTimeout(() => { sessionActive = true }, 5000) // Set session active after 5 seconds
            }
        }
    }
    
     // Listen to data channel, to get the event from the server
    peerConnection.addEventListener("datachannel", event => {
        peerConnectionDataChannel = event.channel
        peerConnectionDataChannel.onmessage = e => {
            let subtitles = document.getElementById('subtitles')
            const webRTCEvent = JSON.parse(e.data)
            const eventType = webRTCEvent.event.eventType
            if (eventType === 'EVENT_TYPE_TURN_START') {
                hideLocalIdleVideo()
                if (isSubtitlesEnabled()) {
                    subtitles.hidden = false
                    subtitles.innerHTML = speakingText
                }
            } else if (eventType === 'EVENT_TYPE_SWITCH_TO_IDLE') {
                subtitles.hidden = true
                showLocalIdleVideo()
            } else if (eventType === 'EVENT_TYPE_SESSION_END') {
                subtitles.hidden = true
                showLocalIdleVideo()
                console.log(`[${(new Date()).toISOString()}] Session end event received from avatar.`)
                transitionToPreSessionState({ reason: 'avatar-session-end' }).catch((err) => {
                    console.error('Failed to reset state after avatar session ended', err)
                })
            }

            console.log("[" + (new Date()).toISOString() + "] WebRTC event received: " + e.data)
        }
    })

    // This is a workaround to make sure the data channel listening is working by creating a data channel from the client side
    const c = peerConnection.createDataChannel("eventChannel")

    // Make necessary update to the web page when the connection state changes
    peerConnection.oniceconnectionstatechange = e => {
        console.log("WebRTC status: " + peerConnection.iceConnectionState)
        if (peerConnection.iceConnectionState === 'disconnected') {
            showLocalIdleVideo()
        } else if (peerConnection.iceConnectionState === 'connected' || peerConnection.iceConnectionState === 'completed') {
            hideLocalIdleVideo()
        }
    }

    // Offer to receive 1 audio, and 1 video track
    peerConnection.addTransceiver('video', { direction: 'sendrecv' })
    peerConnection.addTransceiver('audio', { direction: 'sendrecv' })

    // start avatar, establish WebRTC connection
    avatarSynthesizer.startAvatarAsync(peerConnection).then((r) => {
        if (r.reason === SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
            console.log("[" + (new Date()).toISOString() + "] Avatar started. Result ID: " + r.resultId)
        } else {
            console.log("[" + (new Date()).toISOString() + "] Unable to start avatar. Result ID: " + r.resultId)
            const spinner = document.getElementById('loadingSpinner')
            if (spinner) {
                spinner.style.visibility = 'hidden'
            }
            const preSession = document.getElementById('preSession')
            if (preSession) {
                preSession.hidden = false
                const startButton = preSession.querySelector('#startSession')
                if (startButton) {
                    startButton.style.visibility = 'visible'
                }
            }
            if (!document.body.classList.contains('intro-mode')) {
                document.body.classList.add('intro-mode')
            }
            if (r.reason === SpeechSDK.ResultReason.Canceled) {
                let cancellationDetails = SpeechSDK.CancellationDetails.fromResult(r)
                if (cancellationDetails.reason === SpeechSDK.CancellationReason.Error) {
                    console.log(cancellationDetails.errorDetails)
                };

                console.log("Unable to start avatar: " + cancellationDetails.errorDetails);
            }
            document.getElementById('startSession').disabled = false;
        }
    }).catch(
        (error) => {
            console.log("[" + (new Date()).toISOString() + "] Avatar failed to start. Error: " + error)
            document.getElementById('startSession').disabled = false
            const spinner = document.getElementById('loadingSpinner')
            if (spinner) {
                spinner.style.visibility = 'hidden'
            }
            const preSession = document.getElementById('preSession')
            if (preSession) {
                preSession.hidden = false
                const startButton = preSession.querySelector('#startSession')
                if (startButton) {
                    startButton.style.visibility = 'visible'
                }
            }
            if (!document.body.classList.contains('intro-mode')) {
                document.body.classList.add('intro-mode')
            }
        }
    )
}

// Initialize messages
function initMessages() {
    messages = []
    spokenAssistantMessageIds = new Set()

    if (dataSources.length === 0) {
        let systemPrompt = getRuntimeString('agent', 'systemPrompt', document.getElementById('prompt')?.value || '')
        let systemMessage = {
            role: 'system',
            content: systemPrompt
        }

        messages.push(systemMessage)
    }
}

// Set data sources for chat API
function setDataSources(azureCogSearchEndpoint, azureCogSearchApiKey, azureCogSearchIndexName) {
    let dataSource = {
        type: 'AzureCognitiveSearch',
        parameters: {
            endpoint: azureCogSearchEndpoint,
            key: azureCogSearchApiKey,
            indexName: azureCogSearchIndexName,
            semanticConfiguration: '',
            queryType: 'simple',
            fieldsMapping: {
                contentFieldsSeparator: '\n',
                contentFields: ['content'],
                filepathField: null,
                titleField: 'title',
                urlField: null
            },
            inScope: true,
            roleInformation: getRuntimeString('agent', 'systemPrompt', document.getElementById('prompt')?.value || '')
        }
    }

    dataSources.push(dataSource)
}

// Do HTML encoding on given text
function htmlEncode(text) {
    const entityMap = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
      '/': '&#x2F;'
    };

    return String(text).replace(/[&<>"'\/]/g, (match) => entityMap[match])
}

// Speak the given text
function speak(text, endingSilenceMs = 0) {
    const rawText = typeof text === 'string' ? text : (text === undefined || text === null ? '' : String(text))
    const sanitizedText = stripInlineCitations(rawText)
    const finalText = typeof sanitizedText === 'string' ? sanitizedText : rawText

    if (isSpeaking) {
        spokenTextQueue.push(finalText)
        return
    }

    speakNext(finalText, endingSilenceMs)
}

function speakNext(text, endingSilenceMs = 0, skipUpdatingChatHistory = false) {
    const rawText = typeof text === 'string' ? text : (text === undefined || text === null ? '' : String(text))
    const sanitizedText = stripInlineCitations(rawText)
    const finalText = typeof sanitizedText === 'string' ? sanitizedText : rawText
    const ttsVoice = getRuntimeString('speech', 'ttsVoice', 'en-US-AvaMultilingualNeural')
    let ssml = `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='http://www.w3.org/2001/mstts' xml:lang='en-US'><voice name='${ttsVoice}'><mstts:leadingsilence-exact value='0'/>${htmlEncode(finalText)}</voice></speak>`
    if (endingSilenceMs > 0) {
        ssml = `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='http://www.w3.org/2001/mstts' xml:lang='en-US'><voice name='${ttsVoice}'><mstts:leadingsilence-exact value='0'/>${htmlEncode(finalText)}<break time='${endingSilenceMs}ms' /></voice></speak>`
    }

    if (enableDisplayTextAlignmentWithSpeech && !skipUpdatingChatHistory) {
        scrollChatHistoryToBottom()
    }

    lastSpeakTime = new Date()
    isSpeaking = true
    speakingText = finalText
    lockMicrophoneWhileSpeaking()
    if (!avatarSynthesizer || typeof avatarSynthesizer.speakSsmlAsync !== 'function') {
        console.warn('Avatar synthesizer is not ready to speak.')
        isSpeaking = false
        unlockMicrophoneAfterSpeaking()
        return
    }
    avatarSynthesizer.speakSsmlAsync(ssml).then(
        (result) => {
            if (result.reason === SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
                console.log(`Speech synthesized to speaker for text [ ${finalText} ]. Result ID: ${result.resultId}`)
                lastSpeakTime = new Date()
            } else {
                console.log(`Error occurred while speaking the SSML. Result ID: ${result.resultId}`)
            }

            speakingText = ''

            if (spokenTextQueue.length > 0) {
                speakNext(spokenTextQueue.shift())
            } else {
                isSpeaking = false
                unlockMicrophoneAfterSpeaking()
            }
        }).catch(
            (error) => {
                console.log(`Error occurred while speaking the SSML: [ ${error} ]`)

                speakingText = ''

                if (spokenTextQueue.length > 0) {
                    speakNext(spokenTextQueue.shift())
                } else {
                    isSpeaking = false
                    unlockMicrophoneAfterSpeaking()
                }
            }
        )
}

function stopSpeaking() {
    lastInteractionTime = new Date()
    spokenTextQueue = []
    const microphoneButton = document.getElementById('microphone')
    if (microphoneButton && microphoneButton.dataset.avatarSpeaking === 'true') {
        microphoneButton.disabled = true
    }
    if (!avatarSynthesizer || typeof avatarSynthesizer.stopSpeakingAsync !== 'function') {
        isSpeaking = false
        unlockMicrophoneAfterSpeaking()
        return
    }
    avatarSynthesizer.stopSpeakingAsync().then(
        () => {
            isSpeaking = false
            unlockMicrophoneAfterSpeaking()
            console.log("[" + (new Date()).toISOString() + "] Stop speaking request sent.")
        }
    ).catch(
        (error) => {
            console.log("Error occurred while stopping speaking: " + error)
            unlockMicrophoneAfterSpeaking()
        }
    )
}

window.stopSpeaking = stopSpeaking

async function handleUserQuery(userQuery, userQueryHTML, imgUrlPath) {
    const agentConfig = getAgentConfiguration()
    if (!agentConfig) {
        releaseMicrophoneAwaitingResponse(true)
        return
    }
    const {
        provider,
        agentApiUrl,
        baseEndpoint,
        projectId,
        agentId,
        pluginId
    } = agentConfig

    if (runInProgress) {
        alert('Please wait for the assistant to finish responding before asking another question.')
        releaseMicrophoneAwaitingResponse(true)
        return
    }

    const avatarReady = await ensureAvatarConnectionReady()
    if (!avatarReady) {
        alert('Unable to connect to the avatar service. Please try again.')
        releaseMicrophoneAwaitingResponse(true)
        return
    }

    let contentMessage = userQuery;
    if (imgUrlPath.trim()) {
        contentMessage = [
            { type: "text", text: userQuery },
            { type: "image_url", image_url: { url: imgUrlPath } }
        ];
    }

    let chatMessage = { role: 'user', content: contentMessage };
    messages.push(chatMessage);

    let chatHistoryTextArea = document.getElementById('chatHistory');
    if (chatHistoryTextArea.innerHTML !== '' && !chatHistoryTextArea.innerHTML.endsWith('\n\n')) {
        chatHistoryTextArea.innerHTML += '\n\n';
    }
    chatHistoryTextArea.innerHTML += imgUrlPath.trim() ? `<br/><br/>User: ${userQueryHTML}` : `<br/><br/>User: ${userQuery}<br/>`;
    scrollChatHistoryToBottom()

    // Stop previous speaking if there is any
    if (isSpeaking) {
        stopSpeaking();
    }

    // For 'bring your data' scenario, chat API currently has long (4s+) latency
    // We return some quick reply here before the chat API returns to mitigate.
    if (dataSources.length > 0 && enableQuickReply) {
        speak(getQuickReply(), 2000);
    }

    try {
        runInProgress = true
        const ensurePayload = {
            provider,
            agentApiUrl,
            baseEndpoint,
            projectId,
            pluginId
        }
        const threadReady = await ensureAgentThread(ensurePayload)
        if (!threadReady) {
            runInProgress = false
            releaseMicrophoneAwaitingResponse(true)
            return
        }
        const messagePayload = Array.isArray(contentMessage) || typeof contentMessage === 'string'
            ? contentMessage
            : String(contentMessage)
        const messageRequest = {
            provider,
            pluginId,
            threadId: currentThreadId,
            role: 'user',
            content: messagePayload
        }
        if (provider !== 'copilot_studio') {
            messageRequest.endpoint = baseEndpoint
            messageRequest.projectId = projectId
        }
        await agentApiRequest(agentApiUrl, 'POST', '/message', messageRequest)
        console.log('Message posted to thread:', currentThreadId)

        if (currentRunEventSource) {
            currentRunEventSource.close()
            currentRunEventSource = null
        }

        assistantMessageBuffers = new Map()
        lastAssistantMessageKey = null
        const completedMessageIds = new Set()
        const spokenDuringRun = new Set()

        const streamQuery = {
            provider,
            pluginId,
            threadId: currentThreadId
        }
        if (provider !== 'copilot_studio') {
            streamQuery.endpoint = baseEndpoint
            streamQuery.projectId = projectId
            streamQuery.agentId = agentId
        }

        const streamUrl = buildAgentApiUrl(agentApiUrl, '/run-stream', streamQuery)

        let isFirstAssistantMessageForRun = true
        const eventSource = new EventSource(streamUrl)
        currentRunEventSource = eventSource

        const finalizeStream = () => {
            if (currentRunEventSource === eventSource) {
                currentRunEventSource.close()
                currentRunEventSource = null
            }
            runInProgress = false
            assistantMessageBuffers.forEach((entry) => {
                if (!entry) {
                    return
                }

                if (!Array.isArray(entry.attachments)) {
                    entry.attachments = []
                }

                renderAssistantAttachments(entry, entry.attachments)

                if (entry.hasSpoken) {
                    return
                }

                const identifier = entry.messageId || entry.key
                if (identifier && (spokenAssistantMessageIds.has(identifier) || spokenDuringRun.has(identifier))) {
                    entry.hasSpoken = true
                    return
                }

                const trimmed = (entry.text || '').trim()
                const hasAttachments = entry.attachments.length > 0
                if (!trimmed && !hasAttachments) {
                    return
                }

                if (!spokenDuringRun.size) {
                    chatHistoryTextArea.innerHTML += '<br/>'
                }

                if (trimmed) {
                    const encoded = htmlEncode(trimmed)
                    const spanElement = document.getElementById(entry.spanId)
                    if (spanElement) {
                        spanElement.innerHTML = encoded
                    } else {
                        chatHistoryTextArea.innerHTML += `Assistant: ${encoded}<br/>`
                    }
                    speak(trimmed)
                }

                if (identifier) {
                    spokenAssistantMessageIds.add(identifier)
                    if (trimmed) {
                        spokenDuringRun.add(identifier)
                    }
                }

                const attachmentsForHistory = entry.attachments.length
                    ? entry.attachments.slice(0, MAX_ASSISTANT_ATTACHMENTS)
                    : []
                messages.push({ role: 'assistant', content: entry.text, attachments: attachmentsForHistory })
                entry.hasSpoken = true
            })
            assistantMessageBuffers.clear()
            lastAssistantMessageKey = null

            const microphoneButton = document.getElementById('microphone')
            if (microphoneButton && microphoneButton.dataset.awaitingResponse === 'true') {
                microphoneButton.dataset.awaitingReleaseReady = 'true'
            }

            if (!spokenDuringRun.size) {
                releaseMicrophoneAwaitingResponse()
            }
        }

        const generateTempMessageKey = () => `temp-${Date.now()}-${Math.random().toString(36).slice(2)}`

        const ensureAssistantMessageEntry = (payload) => {
            const payloadMessageId = payload?.message?.id ?? payload?.data?.message?.id ?? null
            const candidateKeys = [
                payloadMessageId,
                payload?.message_id,
                payload?.messageId,
                payload?.parent_id,
                lastAssistantMessageKey,
                payload?.id
            ].filter((value) => typeof value === 'string' && value.trim() !== '')

            let resolvedKey = null
            for (const candidate of candidateKeys) {
                if (assistantMessageBuffers.has(candidate)) {
                    resolvedKey = candidate
                    break
                }
            }

            if (!resolvedKey) {
                resolvedKey = candidateKeys.length > 0 ? candidateKeys[0] : generateTempMessageKey()
            }

            const safeKeyFragment = resolvedKey.replace(/[^a-zA-Z0-9_-]/g, '') || Date.now().toString(36)

            let entry = assistantMessageBuffers.get(resolvedKey)

            if (!entry) {
                const spanId = `assistant-msg-${safeKeyFragment}-text`
                const attachmentsId = `assistant-msg-${safeKeyFragment}-attachments`
                const prefix = isFirstAssistantMessageForRun && imgUrlPath.trim() ? 'Assistant: ' : '<br/>Assistant: '
                chatHistoryTextArea.innerHTML += `${prefix}<span id="${spanId}" class="assistant-output-text"></span><div id="${attachmentsId}" class="assistant-attachments" hidden></div>`
                scrollChatHistoryToBottom()

                entry = {
                    key: resolvedKey,
                    spanId,
                    text: '',
                    messageId: payloadMessageId ?? payload?.message_id ?? payload?.messageId ?? null,
                    attachmentsId,
                    attachments: [],
                    attachmentsRendered: false,
                    hasSpoken: false
                }

                assistantMessageBuffers.set(resolvedKey, entry)
                isFirstAssistantMessageForRun = false
            } else if (!entry.messageId && payloadMessageId) {
                entry.messageId = payloadMessageId
                if (!entry.attachmentsId) {
                    entry.attachmentsId = `assistant-msg-${safeKeyFragment}-attachments`
                }
                if (!Array.isArray(entry.attachments)) {
                    entry.attachments = []
                }
            } else {
                if (!entry.attachmentsId) {
                    entry.attachmentsId = `assistant-msg-${safeKeyFragment}-attachments`
                }
                if (!Array.isArray(entry.attachments)) {
                    entry.attachments = []
                }
            }

            lastAssistantMessageKey = entry.key
            return entry
        }

        const updateAssistantMessageDisplay = (entry) => {
            const spanElement = document.getElementById(entry.spanId)
            if (spanElement) {
                spanElement.innerHTML = htmlEncode(entry.text).replace(/\n/g, '<br/>')
            }
            scrollChatHistoryToBottom()
        }

        const handleAssistantDelta = (payload) => {
            notifyAgentContentObservers(payload)
            const deltaResult = extractTextFromStreamDelta(payload)
            if (!deltaResult || !deltaResult.text) {
                return
            }

            const entry = ensureAssistantMessageEntry(payload)
            const incoming = deltaResult.text

            if (deltaResult.replace) {
                entry.text = incoming
            } else {
                const existing = entry.text || ''
                if (incoming.startsWith(existing)) {
                    entry.text = incoming
                } else {
                    let overlapLength = 0
                    const maxOverlap = Math.min(existing.length, incoming.length)
                    for (let i = maxOverlap; i > 0; i--) {
                        if (existing.endsWith(incoming.slice(0, i))) {
                            overlapLength = i
                            break
                        }
                    }
                    entry.text = existing + incoming.slice(overlapLength)
                }
            }

            updateAssistantMessageDisplay(entry)
        }

        const handleAssistantCompletion = (payload) => {
            if (!payload || (payload.role && payload.role !== 'assistant')) {
                notifyAgentContentObservers(payload)
                return
            }

            if (typeof window !== 'undefined' && window.DEBUG_AGENT_RESPONSES === true) {
                try {
                    const serialized = JSON.stringify(payload, null, 2)
                    console.log('[AgentResponse]', serialized)
                } catch (err) {
                    console.log('[AgentResponse] (unserializable payload)', payload)
                }
            }

            const completionIdentifier = payload?.message?.id
                ?? payload?.data?.message?.id
                ?? payload?.message_id
                ?? payload?.messageId
                ?? payload?.id
                ?? lastAssistantMessageKey
            if (completionIdentifier && completedMessageIds.has(completionIdentifier)) {
                notifyAgentContentObservers(payload)
                return
            }
            if (completionIdentifier) {
                completedMessageIds.add(completionIdentifier)
            }

            const entry = ensureAssistantMessageEntry(payload)
            const messageContent = resolveAssistantMessageContent(payload)
            const replyText = extractTextFromMessageContent(messageContent)
            if (replyText) {
                entry.text = replyText
            }

            const attachments = extractAssistantAttachmentsFromPayload(payload)
            if (Array.isArray(attachments) && attachments.length > 0) {
                payload.__attachments = attachments
            } else if (payload && Object.prototype.hasOwnProperty.call(payload, '__attachments')) {
                delete payload.__attachments
            }

            // Pass the structured content array to plugins (for Copilot Studio parsed JSON)
            if (Array.isArray(messageContent) && messageContent.length > 0) {
                payload.__structuredContent = messageContent
            }

            notifyAgentContentObservers(payload)

            entry.attachments = attachments
            renderAssistantAttachments(entry, attachments)

            updateAssistantMessageDisplay(entry)

            const identifier = entry.messageId || entry.key
            if (!identifier || spokenAssistantMessageIds.has(identifier) || entry.hasSpoken) {
                return
            }

            const trimmedReply = (entry.text || '').trim()
            const hasAttachments = attachments.length > 0
            if (trimmedReply) {
                if (!spokenDuringRun.size) {
                    chatHistoryTextArea.innerHTML += '<br/>'
                }
                const encodedReply = htmlEncode(trimmedReply)
                const spanElement = document.getElementById(entry.spanId)
                if (spanElement) {
                    spanElement.innerHTML = encodedReply
                } else {
                    chatHistoryTextArea.innerHTML += `Assistant: ${encodedReply}<br/>`
                }
                speak(trimmedReply)
            } else if (hasAttachments && !spokenDuringRun.size) {
                chatHistoryTextArea.innerHTML += '<br/>'
            }

            spokenAssistantMessageIds.add(identifier)
            if (trimmedReply) {
                spokenDuringRun.add(identifier)
            }
            const attachmentsForHistory = hasAttachments
                ? attachments.slice(0, MAX_ASSISTANT_ATTACHMENTS)
                : []
            messages.push({ role: 'assistant', content: entry.text, attachments: attachmentsForHistory })
            entry.hasSpoken = true
            assistantMessageBuffers.delete(entry.key)
            if (assistantMessageBuffers.size === 0) {
                lastAssistantMessageKey = null
            }
        }

        const safeParse = (raw) => {
            if (typeof raw !== 'string' || raw.trim() === '') {
                return null
            }
            try {
                return JSON.parse(raw)
            } catch (parseError) {
                console.error('Failed to parse run stream payload', parseError, raw)
                return null
            }
        }

        eventSource.addEventListener('thread.message.delta', (evt) => {
            const payload = safeParse(evt.data)
            handleAssistantDelta(payload)
        })

        eventSource.addEventListener('message.delta', (evt) => {
            const payload = safeParse(evt.data)
            handleAssistantDelta(payload)
        })

        eventSource.addEventListener('thread.message.completed', (evt) => {
            const payload = safeParse(evt.data)
            handleAssistantCompletion(payload)
        })

        eventSource.addEventListener('message.completed', (evt) => {
            const payload = safeParse(evt.data)
            handleAssistantCompletion(payload)
        })

        eventSource.addEventListener('thread.run.completed', () => {
            finalizeStream()
        })

        eventSource.addEventListener('run.completed', () => {
            finalizeStream()
        })

        eventSource.addEventListener('thread.run.failed', (evt) => {
            finalizeStream()
            const payload = safeParse(evt.data)
            const message = payload?.lastError?.message ?? 'Agent run failed. Please try again.'
            setTimeout(() => alert(message), 0)
        })

        eventSource.addEventListener('error', (evt) => {
            const payload = safeParse(evt.data)
            if (payload?.error) {
                console.error('Run stream error:', payload)
                setTimeout(() => alert(payload.error), 0)
            } else {
                console.error('Run stream error:', evt)
            }
            finalizeStream()
        })

        eventSource.onerror = (evt) => {
            console.error('Run stream network error:', evt)
            finalizeStream()
        }

        eventSource.addEventListener('done', () => {
            finalizeStream()
        })
    } catch (err) {
        console.error('Agent API flow error:', err)
        runInProgress = false
        if (err?.response) {
            console.error('Agent API flow error detail:', err.response.status, err.response.parsedBody)
        }
        const formattedError = provider === 'copilot_studio'
            ? (err?.message || 'Agent API call failed.')
            : formatAIProjectsError(err)
        alert(`Agent API call failed: ${formattedError}`)
        releaseMicrophoneAwaitingResponse(true)
    }
}

function getQuickReply() {
    return quickReplies[Math.floor(Math.random() * quickReplies.length)]
}

function checkHung() {
    // Check whether the avatar video stream is hung, by checking whether the video time is advancing
    let videoElement = document.getElementById('videoPlayer')
    if (videoElement !== null && videoElement !== undefined && sessionActive) {
        let videoTime = videoElement.currentTime
        setTimeout(() => {
            // Check whether the video time is advancing
            if (videoElement.currentTime === videoTime) {
                // Check whether the session is active to avoid duplicatedly triggering reconnect
                if (sessionActive) {
                    sessionActive = false
                    if (isAutoReconnectEnabled()) {
                        // No longer reconnect when there is no interaction for a while
                        if (new Date() - lastInteractionTime < 300000) {
                            console.log(`[${(new Date()).toISOString()}] The video stream got disconnected, need reconnect.`)
                            isReconnecting = true
                            // Remove data channel onmessage callback to avoid duplicatedly triggering reconnect
                            peerConnectionDataChannel.onmessage = null
                            // Release the existing avatar connection
                            if (avatarSynthesizer !== undefined) {
                                avatarSynthesizer.close()
                            }
    
                            // Setup a new avatar connection
                            connectAvatar()
                                .then((started) => {
                                    if (!started) {
                                        isReconnecting = false
                                    }
                                })
                                .catch((err) => {
                                    console.error('Failed to reconnect avatar after hang detection', err)
                                    isReconnecting = false
                                })
                        }
                    }
                }
            }
        }, 2000)
    }
}

function checkLastSpeak() {
    if (lastSpeakTime === undefined) {
        return
    }

    let currentTime = new Date()
    if (currentTime - lastSpeakTime > 15000) {
        if (isLocalIdleVideoEnabled() && sessionActive && !isSpeaking) {
            disconnectAvatar()
            showLocalIdleVideo()
            sessionActive = false
        }
    }
}

window.onload = () => {
    setInterval(() => {
        checkHung()
        checkLastSpeak()
    }, 2000) // Check session activity every 2 seconds
    setupTextInputToggle()
    setupMobileSessionActions()
    ensureUserMessageInput()
}

window.startSession = async () => {
    const spinner = document.getElementById('loadingSpinner')
    if (spinner) {
        spinner.style.visibility = 'visible'
    }
    
    const preSession = document.getElementById('preSession')
    if(preSession) {
        //preSession.style.height = '100%'
    }
    lastInteractionTime = new Date()
    const body = document.body
    if (body.classList.contains('intro-mode')) {
        body.classList.remove('intro-mode')
    }

    const appContent = document.getElementById('appContent')
    if (appContent) {
        appContent.hidden = false
    }

    if (preSession) {
        preSession.hidden = true
        const startButton = preSession.querySelector('#startSession')
        if (startButton) {
            startButton.style.visibility = 'hidden'
        }
    }

    const agentConfig = getAgentConfiguration()
    if (!agentConfig) {
        const configPanel = document.getElementById('configuration')
        if (configPanel) {
            configPanel.hidden = false
        }
        if (preSession) {
            preSession.hidden = false
            const startButton = preSession.querySelector('#startSession')
            if (startButton) {
                startButton.style.visibility = 'visible'
            }
        }
        if (spinner) {
            spinner.style.visibility = 'hidden'
        }
        if (appContent) {
            appContent.hidden = true
        }
        if (appContent) {
            appContent.hidden = true
        }
        if (!body.classList.contains('intro-mode')) {
            body.classList.add('intro-mode')
        }
        return
    }

    const threadReady = await ensureAgentThread({
        provider: agentConfig.provider,
        agentApiUrl: agentConfig.agentApiUrl,
        baseEndpoint: agentConfig.baseEndpoint,
        projectId: agentConfig.projectId,
        pluginId: agentConfig.pluginId
    })
    if (!threadReady) {
        if (preSession) {
            preSession.hidden = false
            const startButton = preSession.querySelector('#startSession')
            if (startButton) {
                startButton.style.visibility = 'visible'
            }
        }
        if (spinner) {
            spinner.style.visibility = 'hidden'
        }
        if (appContent) {
            appContent.hidden = true
        }
        if (!body.classList.contains('intro-mode')) {
            body.classList.add('intro-mode')
        }
        return
    }


    const sessionButtons = document.getElementById('sessionButtons')
    const chatControls = document.getElementById('chatControls')
    const videoContainer = document.getElementById('videoContainer')

    if (isLocalIdleVideoEnabled()) {
        if (sessionButtons) {
            sessionButtons.hidden = false
        }
        if (chatControls) {
            chatControls.hidden = false
        }
        if (videoContainer) {
            videoContainer.hidden = false
        }
        document.getElementById('startSession').disabled = true
        document.getElementById('microphone').disabled = false
        document.getElementById('stopSession').disabled = false
    showLocalIdleVideo()
        document.getElementById('chatHistory').hidden = true
        const toggleButton = document.getElementById('toggleChatHistory')
        if (toggleButton) {
            updateButtonLabel(toggleButton, 'Show Chat History')
        }
        setMicrophoneState('inactive')
        ensureUserMessageInput()
        if (spinner) {
            spinner.style.visibility = 'hidden'
        }
        return
    }

    userClosedSession = false
    const connectionStarted = await connectAvatar()
    if (!connectionStarted) {
        if (sessionButtons) {
            sessionButtons.hidden = true
        }
        if (chatControls) {
            chatControls.hidden = true
        }
        if (videoContainer) {
            videoContainer.hidden = true
        }
        if (configPanel) {
            configPanel.hidden = false
        }
        if (preSession) {
            preSession.hidden = false
            const startButton = preSession.querySelector('#startSession')
            if (startButton) {
                startButton.style.visibility = 'visible'
            }
        }
        setMicrophoneState('inactive')
        if (spinner) {
            spinner.style.visibility = 'hidden'
        }
        if (appContent) {
            appContent.hidden = true
        }
        if (!body.classList.contains('intro-mode')) {
            body.classList.add('intro-mode')
        }
        const startBtn = document.getElementById('startSession')
        if (startBtn) {
            startBtn.disabled = false
        }
        return
    }

    if (sessionButtons) {
        sessionButtons.hidden = false
    }
    if (chatControls) {
        chatControls.hidden = false
        if (typeof applyTextInputToggleLayout === 'function') {
            applyTextInputToggleLayout(true)
        }
    }
    if (preSession) {
        preSession.hidden = true
    }
    ensureUserMessageInput()
}

async function transitionToPreSessionState({ reason = 'manual-stop' } = {}) {
    if (sessionShutdownPromise) {
        return sessionShutdownPromise
    }

    const shutdownTask = (async () => {
        console.log(`[${(new Date()).toISOString()}] Resetting to pre-session state (reason: ${reason}).`)
        lastInteractionTime = new Date()

        const startSessionButton = document.getElementById('startSession')
        if (startSessionButton) {
            startSessionButton.disabled = false
        }

        const microphoneButton = document.getElementById('microphone')
        if (microphoneButton) {
            microphoneButton.disabled = true
            delete microphoneButton.dataset.awaitingResponse
            delete microphoneButton.dataset.awaitingReleaseReady
        }

        const stopSessionButton = document.getElementById('stopSession')
        if (stopSessionButton) {
            stopSessionButton.disabled = true
        }

        if (typeof window.closeMobileActionsMenu === 'function') {
            window.closeMobileActionsMenu(false)
        }

        const mobileToggleButton = document.getElementById('mobileTextInputToggle')
        if (mobileToggleButton) {
            mobileToggleButton.setAttribute('aria-pressed', 'false')
        }

        setMicrophoneState('inactive')

        const chatHistory = document.getElementById('chatHistory')
        if (chatHistory) {
            chatHistory.hidden = true
        }

        const messageBox = document.getElementById('userMessageBox')
        if (messageBox) {
            messageBox.hidden = true
            messageBox.dataset.deferredDisplay = 'true'
            messageBox.classList.remove('floating-input')
            messageBox.innerHTML = ''
        }

        const floatingBackdrop = document.getElementById('textInputBackdrop')
        if (floatingBackdrop) {
            floatingBackdrop.hidden = true
        }

        const sessionButtons = document.getElementById('sessionButtons')
        if (sessionButtons) {
            sessionButtons.hidden = true
        }

        const chatControls = document.getElementById('chatControls')
        if (chatControls) {
            chatControls.hidden = true
            chatControls.classList.remove('show-input')
        }

        const textToggleButton = document.getElementById('textInputToggle')
        if (textToggleButton) {
            textToggleButton.setAttribute('aria-pressed', 'false')
        }

        if (typeof applyTextInputToggleLayout === 'function') {
            applyTextInputToggleLayout(true)
        }

        const videoContainer = document.getElementById('videoContainer')
        if (videoContainer) {
            videoContainer.hidden = true
        }

        const appContent = document.getElementById('appContent')
        if (appContent) {
            appContent.hidden = true
        }

        const preSession = document.getElementById('preSession')
        if (preSession) {
            preSession.hidden = false
            const preSessionStartButton = preSession.querySelector('#startSession')
            if (preSessionStartButton) {
                preSessionStartButton.style.visibility = 'visible'
            }
        }

        if (!document.body.classList.contains('intro-mode')) {
            document.body.classList.add('intro-mode')
        }

        if (isLocalIdleVideoEnabled()) {
            hideLocalIdleVideo()
        }

        userClosedSession = true
        isReconnecting = false

        const toggleChatHistoryButton = document.getElementById('toggleChatHistory')
        if (toggleChatHistoryButton) {
            updateButtonLabel(toggleChatHistoryButton, 'Show Chat History')
        }

        if (currentRunEventSource) {
            currentRunEventSource.close()
            currentRunEventSource = null
        }

        assistantMessageBuffers.clear()
        lastAssistantMessageKey = null
        runInProgress = false

        attachmentOverlayState.attachments = []
        if (pluginHostState.activeInstance && typeof pluginHostState.activeInstance.clearOverlayContent === 'function') {
            try {
                pluginHostState.activeInstance.clearOverlayContent()
            } catch (err) {
                pluginDebugLog('Plugin clearOverlayContent failed during session reset', err?.message || err)
            }
        }
        setPluginOverlayDescriptor(null)

        disconnectAvatar()
        spokenTextQueue.length = 0
        messageInitiated = false
        initMessages()

        if (chatHistory) {
            chatHistory.innerHTML = ''
        }

        scrollChatHistoryToBottom(true)

        try {
            await resetAgentThread()
        } catch (err) {
            console.warn('Failed to reset agent thread after stopping the session', err)
        }
    })()

    sessionShutdownPromise = shutdownTask

    return shutdownTask.finally(() => {
        sessionShutdownPromise = null
        console.log(`[${(new Date()).toISOString()}] Transitioned to pre-session state (reason: ${reason}).`)
    })
}

window.stopSession = async () => {
    await transitionToPreSessionState({ reason: 'user-stop' })
}

window.clearChatHistory = async () => {
    lastInteractionTime = new Date()
    const chatHistory = document.getElementById('chatHistory')
    if (chatHistory) {
        chatHistory.innerHTML = ''
    }
    scrollChatHistoryToBottom(true)
    initMessages()
    setPluginOverlayDescriptor(null)
    if (currentRunEventSource) {
        currentRunEventSource.close()
        currentRunEventSource = null
    }
    runInProgress = false
    await resetAgentThread()
}

window.microphone = async () => {
    lastInteractionTime = new Date()
    const microphoneButton = document.getElementById('microphone')
    if (!microphoneButton) {
        return
    }

    if (microphoneButton.dataset.avatarSpeaking === 'true') {
        microphoneButton.disabled = true
        stopSpeaking()
        return
    }

    const isActive = microphoneButton.dataset.state === 'active'

    if (isActive) {
        if (!speechRecognizer || typeof speechRecognizer.stopContinuousRecognitionAsync !== 'function') {
            setMicrophoneState('inactive')
            microphoneButton.disabled = false
            return
        }

        microphoneButton.disabled = true
        if (!microphoneButton.dataset.stopReason) {
            microphoneButton.dataset.stopReason = 'manual'
        }
        speechRecognizer.stopContinuousRecognitionAsync(
            () => {
                const stopReason = microphoneButton.dataset.stopReason
                delete microphoneButton.dataset.stopReason
                setMicrophoneState('inactive')
                if (stopReason === 'auto') {
                    engageMicrophoneAwaitingResponse()
                } else {
                    microphoneButton.disabled = false
                }
            }, (err) => {
                console.log("Failed to stop continuous recognition:", err)
                delete microphoneButton.dataset.stopReason
                microphoneButton.disabled = false
                setMicrophoneState('inactive')
            })

        return
    }

    if (!speechRecognizer || typeof speechRecognizer.startContinuousRecognitionAsync !== 'function') {
        const ready = await ensureAvatarConnectionReady()
        if (!ready) {
            alert('Unable to connect to the avatar service. Please try again.')
            return
        }
    }

    if (isLocalIdleVideoEnabled()) {
        if (!sessionActive) {
            const ready = await ensureAvatarConnectionReady()
            if (!ready) {
                alert('Unable to connect to the avatar service. Please try again.')
                return
            }
        }

        setTimeout(() => {
            document.getElementById('audioPlayer').play()
        }, 5000)
    } else {
        document.getElementById('audioPlayer').play()
    }

    microphoneButton.disabled = true
    speechRecognizer.recognized = async (s, e) => {
        if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
            let userQuery = e.result.text.trim()
            if (userQuery === '') {
                return
            }

            // Auto stop microphone when a phrase is recognized, when it's not continuous conversation mode
            if (!isContinuousConversationEnabled()) {
                microphoneButton.disabled = true
                microphoneButton.dataset.stopReason = 'auto'
                speechRecognizer.stopContinuousRecognitionAsync(
                    () => {
                        const stopReason = microphoneButton.dataset.stopReason
                        delete microphoneButton.dataset.stopReason
                        setMicrophoneState('inactive')
                        if (stopReason === 'auto') {
                            engageMicrophoneAwaitingResponse()
                        } else {
                            microphoneButton.disabled = false
                        }
                    }, (err) => {
                        console.log("Failed to stop continuous recognition:", err)
                        delete microphoneButton.dataset.stopReason
                        microphoneButton.disabled = false
                        setMicrophoneState('inactive')
                    })
            }

            handleUserQuery(userQuery,"","")
        }
    }

    speechRecognizer.startContinuousRecognitionAsync(
        () => {
            setMicrophoneState('active')
            microphoneButton.disabled = false
        }, (err) => {
            console.log("Failed to start continuous recognition:", err)
            setMicrophoneState('inactive')
            microphoneButton.disabled = false
        })
}

window.updataEnableOyd = () => {
    const enableOyd = document.getElementById('enableOyd')
    const cogSearchConfig = document.getElementById('cogSearchConfig')
    if (!enableOyd || !cogSearchConfig) {
        return
    }
    if (enableOyd.checked) {
        cogSearchConfig.hidden = false
    } else {
        cogSearchConfig.hidden = true
    }
}

window.updateLocalVideoForIdle = () => {
    // No-op retained for compatibility now that the type message control is always visible
}

window.updatePrivateEndpoint = () => {
    const enablePrivateEndpoint = document.getElementById('enablePrivateEndpoint')
    const showPrivateEndpointCheckBox = document.getElementById('showPrivateEndpointCheckBox')
    if (!enablePrivateEndpoint || !showPrivateEndpointCheckBox) {
        return
    }
    if (enablePrivateEndpoint.checked) {
        showPrivateEndpointCheckBox.hidden = false
    } else {
        showPrivateEndpointCheckBox.hidden = true
    }
}

window.updateCustomAvatarBox = () => {
    if (document.getElementById('customizedAvatar').checked) {
        document.getElementById('useBuiltInVoice').disabled = false
    } else {
        document.getElementById('useBuiltInVoice').disabled = true
        document.getElementById('useBuiltInVoice').checked = false
    }
}

window.toggleChatHistory = () => {
    lastInteractionTime = new Date()
    const chatHistory = document.getElementById('chatHistory')
    const toggleButton = document.getElementById('toggleChatHistory')
    if (!chatHistory || !toggleButton) {
        return
    }

    const willShow = chatHistory.hidden
    chatHistory.hidden = !chatHistory.hidden
    updateButtonLabel(toggleButton, willShow ? 'Hide Chat History' : 'Show Chat History')
}


function extractTextFromMessageContent(content, depth = 0) {
    const debugEnabled = typeof window !== 'undefined' && window.DEBUG_SPEECH_EXTRACTION === true
    const log = (...args) => {
        if (debugEnabled) {
            console.log(`[SpeechText][${depth}]`, ...args)
        }
    }

    if (depth > 6) {
        log('max recursion depth reached')
        return ''
    }

    log('raw content:', content)

    if (!content) {
        log('no content provided')
        return ''
    }

    const tryParseJson = (text) => {
        const trimmed = text.trim()
        if (!trimmed) {
            log('string trimmed to empty')
            return ''
        }

        const firstChar = trimmed[0]
        if (firstChar === '{' || firstChar === '[') {
            try {
                const parsed = JSON.parse(trimmed)
                log('parsed JSON string successfully')
                return extractTextFromMessageContent(parsed, depth + 1)
            } catch (err) {
                log('failed to parse JSON string, returning literal text')
                return trimmed
            }
        }

        log('returning literal text')
        return trimmed
    }

    if (typeof content === 'string') {
        return tryParseJson(content)
    }

    if (Array.isArray(content)) {
        if (!content.length) {
            log('array content empty')
            return ''
        }
        const combined = content
            .map((entry) => extractTextFromMessageContent(entry, depth + 1))
            .filter(Boolean)
            .join('\n')
        log('array combined result:', combined)
        return combined
    }

    if (typeof content !== 'object') {
        log('unsupported content type')
        return ''
    }

    if (typeof content.value === 'string') {
        log('found value property containing string')
        return tryParseJson(content.value)
    }

    if (typeof content.text === 'string' && !Array.isArray(content.content)) {
        log('found text property containing string')
        return tryParseJson(content.text)
    }

    if (Array.isArray(content.text)) {
        log('found text property containing array')
        return extractTextFromMessageContent(content.text, depth + 1)
    }

    if (typeof content.text === 'object' && content.text !== null) {
        log('found text property containing nested object')
        const nested = extractTextFromMessageContent(content.text, depth + 1)
        if (nested) {
            return nested
        }
    }

    const items = Array.isArray(content.content) ? content.content : []
    if (!items.length) {
        log('no content items present')
        return ''
    }

    const segments = []
    for (const item of items) {
        if (!item) {
            continue
        }

        if (item.name === 'text') {
            const resolved = extractTextFromMessageContent(item.text, depth + 1)
            if (resolved) {
                segments.push(resolved)
                log('appended text segment from name field:', resolved)
            } else {
                log('text name item resolved empty')
            }
            continue
        }

        if (item.type === 'text') {
            log('encountered type="text" item without name, inspecting nested text payload')
            const nested = extractTextFromMessageContent(item.text, depth + 1)
            if (nested) {
                segments.push(nested)
                log('appended text segment from type field:', nested)
            } else {
                log('type="text" item produced no speech text')
            }
            continue
        }

        log('skipping non-text item', item)
    }

    const result = segments.filter(Boolean).join('\n')
    log('final result:', result)
    return result
}

function extractTextFromStreamDelta(payload) {
    if (!payload) {
        return null
    }

    const candidateContent = Array.isArray(payload?.delta?.content)
        ? payload.delta.content
        : Array.isArray(payload?.content)
            ? payload.content
            : Array.isArray(payload?.data?.content)
                ? payload.data.content
                : []

    if (!Array.isArray(candidateContent)) {
        return null
    }

    let replace = false
    let segments = []

    const appendSegment = (value) => {
        if (typeof value === 'string' && value.length > 0) {
            segments.push(value)
        }
    }

    for (const item of candidateContent) {
        if (!item) {
            continue
        }

        if (typeof item === 'string') {
            appendSegment(item)
            continue
        }

        if (item?.type === 'output_text.delta') {
            if (typeof item?.text?.content === 'string') {
                appendSegment(item.text.content)
            } else if (typeof item?.text === 'string') {
                appendSegment(item.text)
            }
            continue
        }

        if (item?.type === 'output_text') {
            let value = null
            if (typeof item?.text?.value === 'string') {
                value = item.text.value
            } else if (typeof item?.text?.content === 'string') {
                value = item.text.content
            } else if (typeof item?.text === 'string') {
                value = item.text
            }
            if (typeof value === 'string') {
                replace = true
                segments = [value]
            }
            continue
        }

        if (item?.type === 'text.delta' || item?.type === 'text_delta') {
            if (typeof item?.text === 'string') {
                appendSegment(item.text)
            } else if (typeof item?.delta === 'string') {
                appendSegment(item.delta)
            }
            continue
        }

        if (typeof item?.delta === 'string') {
            appendSegment(item.delta)
            continue
        }

        if (typeof item?.value === 'string') {
            appendSegment(item.value)
        }
    }

    if (segments.length === 0) {
        return null
    }

    return {
        text: segments.join(''),
        replace
    }
}

function resolveAssistantMessageContent(payload) {
    if (!payload) {
        return null
    }
    
    // For Copilot Studio, check if there's structured content in data.activity.text
    const activityText = payload.data?.activity?.text
    if (typeof activityText === 'string' && activityText.trim()) {
        const trimmed = activityText.trim()
        // Check if it's JSON array
        if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
            try {
                const parsed = JSON.parse(trimmed)
                if (Array.isArray(parsed) && parsed.length > 0) {
                    // Return the parsed array as content
                    return parsed
                }
            } catch (err) {
                // Not valid JSON, fall through to other candidates
            }
        }
    }
    
    const candidates = [
        payload.message?.content,
        payload.data?.message?.content,
        payload.data?.content,
        payload.response?.content,
        payload.content
    ]

    for (const candidate of candidates) {
        if (candidate !== undefined) {
            return candidate
        }
    }

    return null
}

function extractAssistantAttachmentsFromPayload(payload) {
    if (!payload) {
        return []
    }

    // Normalizes any structured attachment objects that arrive alongside the agent's reply
    // so the browser can render images, charts, and other UI widgets inline.
    const results = []
    const seen = new Set()

    const addAttachment = (attachment) => {
        if (!attachment) {
            return
        }
        const keyParts = [
            attachment.type || 'unknown',
            attachment.url || '',
            attachment.kind || '',
            attachment.title || ''
        ]
        const key = keyParts.join('|')
        if (seen.has(key)) {
            return
        }
        seen.add(key)
        results.push(attachment)
    }

    const candidateContents = [
        resolveAssistantMessageContent(payload),
        payload?.response?.content,
        payload?.output,
        payload?.data?.output
    ]

    for (const candidate of candidateContents) {
        if (!candidate) {
            continue
        }
        const extracted = extractAssistantAttachments(candidate)
        if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
            console.log('[AgentAttachments] Extracted from candidate:', extracted)
        }
        for (const attachment of extracted) {
            addAttachment(attachment)
            if (results.length >= MAX_ASSISTANT_ATTACHMENTS) {
                break
            }
        }
        if (results.length >= MAX_ASSISTANT_ATTACHMENTS) {
            break
        }
    }

    return results.slice(0, MAX_ASSISTANT_ATTACHMENTS)
}

function extractAssistantAttachments(content) {
    if (!content) {
        return []
    }

    const results = []
    const visited = new WeakSet()

    const visit = (node) => {
        if (node === null || node === undefined) {
            return
        }

        if (typeof node === 'string') {
            const trimmed = node.trim()
            if (trimmed.length > 2 && (trimmed.startsWith('{') || trimmed.startsWith('['))) {
                try {
                    const parsed = JSON.parse(trimmed)
                    visit(parsed)
                } catch (_) {
                    // Ignore invalid JSON snippets
                }
            }
            return
        }

        if (typeof node !== 'object') {
            return
        }

        if (visited.has(node)) {
            return
        }
        visited.add(node)

        const normalized = normalizeAssistantAttachment(node)
        if (normalized) {
            results.push(normalized)
        }

        const nestedCandidates = [
            node.attachments,
            node.items,
            node.elements,
            node.entries,
            node.children,
            node.values,
            node.content,
            node.contents,
            node.payload,
            node.data,
            node.json,
            node.detail,
            node.details,
            node.blocks,
            node.parts,
            node.cards,
            node.results,
            node.response,
            node.asset,
            node.assets,
            node.text,
            node.message,
            node.output,
            node.value
        ]

        for (const candidate of nestedCandidates) {
            if (!candidate) {
                continue
            }
            if (Array.isArray(candidate)) {
                candidate.forEach(visit)
            } else if (typeof candidate === 'string') {
                const trimmed = candidate.trim()
                if (trimmed.length > 2 && (trimmed.startsWith('{') || trimmed.startsWith('['))) {
                    try {
                        const parsed = JSON.parse(trimmed)
                        visit(parsed)
                    } catch (_) {
                        // Ignore invalid JSON snippets
                    }
                }
            } else {
                visit(candidate)
            }
        }
    }

    if (Array.isArray(content)) {
        content.forEach(visit)
    } else {
        visit(content)
    }

    return results
}

function normalizeAssistantAttachment(raw) {
    if (!raw || typeof raw !== 'object') {
        return null
    }

    // Ignore container nodes that just group attachments; we'll recurse into their children instead.
    if (Array.isArray(raw.attachments) || Array.isArray(raw.items) || Array.isArray(raw.elements) || Array.isArray(raw.children) || Array.isArray(raw.values)) {
        const rawType = typeof raw.type === 'string' ? raw.type.toLowerCase() : ''
        if (!rawType || ['collection', 'attachments', 'list', 'group'].includes(rawType)) {
            return null
        }
    }

    const type = typeof raw.type === 'string' ? raw.type.toLowerCase() : ''
    const role = typeof raw.role === 'string' ? raw.role.toLowerCase() : ''

    if (type === 'placeholder' || raw.placeholder === true) {
        const title = coerceDisplayString(raw.title || raw.name || 'Content unavailable')
        const description = coerceDisplayString(raw.description || raw.reason || raw.message || raw.fallbackText)
        return {
            type: 'placeholder',
            title: title || 'Content unavailable',
            description: description || 'The assistant referenced additional content that could not be displayed.'
        }
    }

    const title = coerceDisplayString(raw.title || raw.name || raw.label)
    const description = coerceDisplayString(raw.description || raw.caption || raw.summary || raw.excerpt)
    const fallbackText = coerceDisplayString(raw.fallbackText || raw.fallback || raw.note || raw.message)

    const urlCandidates = [
        raw.url,
        raw.href,
        raw.link,
        raw.source?.url,
        raw.image_url?.url,
    typeof raw.image_url === 'string' ? raw.image_url : null,
        raw.imageUrl?.url,
    typeof raw.imageUrl === 'string' ? raw.imageUrl : null,
        raw.chartUrl,
        raw.assetUrl,
        raw.previewUrl,
        raw.webUrl,
        raw.path,
        raw.data?.url,
        raw.payload?.url
    ]

    let resolvedUrl = null
    for (const candidate of urlCandidates) {
        const safeUrl = resolveSafeHttpUrl(candidate)
        if (safeUrl) {
            resolvedUrl = safeUrl
            break
        }
    }

    const contentType = (raw.mimeType || raw.contentType || '').toLowerCase()
    const kind = coerceDisplayString(raw.kind || raw.category || raw.widgetType || raw.visualizationType || raw.chartType || raw.component)
    const altText = coerceDisplayString(raw.alt ?? raw.altText ?? raw.caption ?? description ?? title)

    const dataObject = typeof raw.data === 'object' && raw.data !== null
        ? raw.data
        : typeof raw.payload === 'object' && raw.payload !== null
            ? raw.payload
            : typeof raw.details === 'object' && raw.details !== null
                ? raw.details
                : typeof raw.context === 'object' && raw.context !== null
                    ? raw.context
                    : null

    const isImageType = type.includes('image')
        || role.includes('image')
        || (kind && kind.toLowerCase().includes('image'))
        || contentType.startsWith('image/')
        || (resolvedUrl && /\.(png|jpe?g|gif|svg|webp|bmp)$/i.test(resolvedUrl))

    if (isImageType) {
        if (!resolvedUrl) {
            return {
                type: 'placeholder',
                title: title || 'Image unavailable',
                description: description || fallbackText || 'The assistant referenced an image, but no valid URL was provided.'
            }
        }
        if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
            console.log('[AgentAttachments] Normalized image attachment:', {
                title,
                description,
                url: resolvedUrl,
                altText
            })
        }
        return {
            type: 'image',
            title,
            description,
            url: resolvedUrl,
            altText: altText || 'Assistant shared an image',
            fallbackText
        }
    }

    const widgetKeywords = ['chart', 'widget', 'dashboard', 'table', 'report', 'summary', 'analytics', 'visualization', 'graph']
    const candidateStrings = [type, role, kind, title, description].filter(Boolean).map((value) => value.toLowerCase())
    const isWidgetType = widgetKeywords.some((keyword) => candidateStrings.some((value) => value.includes(keyword)))

    if (isWidgetType || dataObject || resolvedUrl) {
        return {
            type: 'widget',
            title,
            description: description || fallbackText,
            kind: kind || (isWidgetType ? (type || role || 'widget') : ''),
            url: resolvedUrl,
            data: dataObject,
            fallbackText
        }
    }

    return null
}

function resolveSafeHttpUrl(candidate) {
    if (typeof candidate !== 'string') {
        return null
    }
    const trimmed = candidate.trim()
    if (!trimmed) {
        return null
    }
    try {
        const base = typeof window !== 'undefined' && window.location ? window.location.origin : undefined
        const url = new URL(trimmed, base)
        if (url.protocol === 'http:' || url.protocol === 'https:') {
            return url.href
        }
        return null
    } catch (_) {
        return null
    }
}

function coerceDisplayString(value) {
    if (typeof value === 'string') {
        return value.trim()
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
        return String(value)
    }
    return ''
}

function truncateForDisplay(text, maxLength = 200) {
    if (!text) {
        return ''
    }
    if (text.length <= maxLength) {
        return text
    }
    const safeLength = Math.max(0, maxLength - 3)
    return `${text.slice(0, safeLength)}...`
}

function summarizeAttachmentData(data) {
    if (!data || typeof data !== 'object') {
        return ''
    }
    try {
        const serialized = JSON.stringify(data, null, 2)
        if (!serialized) {
            return ''
        }
        return truncateForDisplay(serialized, ATTACHMENT_SUMMARY_LIMIT)
    } catch (_) {
        return ''
    }
}

function capitalizeLabel(value) {
    const text = coerceDisplayString(value)
    if (!text) {
        return ''
    }
    return text
        .split(/[\s_-]+/)
        .filter(Boolean)
        .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
        .join(' ')
}

function formatAttachmentTitle(attachment) {
    if (attachment?.title) {
        return attachment.title
    }
    if (attachment?.kind) {
        return capitalizeLabel(attachment.kind)
    }
    if (attachment?.type === 'image') {
        return 'Shared image'
    }
    if (attachment?.type === 'widget') {
        return 'Shared data'
    }
    return 'Assistant content'
}

function renderAssistantAttachments(entry, attachments) {
    if (!entry || !entry.attachmentsId) {
        return
    }
    const container = document.getElementById(entry.attachmentsId)
    if (!container) {
        return
    }

    const normalized = Array.isArray(attachments) ? attachments.filter(Boolean) : []
    if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
        console.log('[AgentAttachments] Rendering attachments for entry:', {
            entryId: entry.key || entry.messageId,
            attachments: normalized
        })
    }
    if (normalized.length === 0) {
        entry.attachmentsRendered = false
        if (typeof container.replaceChildren === 'function') {
            container.replaceChildren()
        } else {
            container.innerHTML = ''
        }
        container.hidden = true
        renderAssistantAttachmentOverlay([])
        return
    }

    const nodes = normalized.slice(0, MAX_ASSISTANT_ATTACHMENTS).map((attachment) => buildAttachmentCard(attachment)).filter(Boolean)
    if (typeof container.replaceChildren === 'function') {
        container.replaceChildren(...nodes)
    } else {
        container.innerHTML = ''
        nodes.forEach((node) => container.appendChild(node))
    }
    container.hidden = false
    entry.attachmentsRendered = true

    if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
        console.log('[AgentAttachments] Attachment nodes rendered:', nodes)
    }

    renderAssistantAttachmentOverlay(normalized)
}

function renderAssistantAttachmentOverlay(attachments) {
    const normalized = Array.isArray(attachments) ? attachments.filter(Boolean) : []
    attachmentOverlayState.attachments = normalized

    if (attachmentOverlayState.pluginDescriptor) {
        if (typeof window !== 'undefined' && window.DEBUG_AGENT_ATTACHMENTS === true) {
            console.log('[AgentAttachments] Plugin overlay active; stored attachment update for fallback.', normalized)
        }
        return
    }

    renderOverlayFromAttachments(normalized)
}

function buildAttachmentCard(attachment) {
    if (!attachment) {
        return createAttachmentPlaceholderCard('Content unavailable', 'The assistant referenced additional content that could not be displayed.')
    }

    if (attachment.type === 'image') {
        return createImageAttachmentCard(attachment)
    }

    if (attachment.type === 'widget') {
        return createWidgetAttachmentCard(attachment)
    }

    if (attachment.type === 'placeholder') {
        return createAttachmentPlaceholderCard(attachment.title, attachment.description || attachment.fallbackText)
    }

    return createAttachmentPlaceholderCard(
        attachment.title || 'Content unavailable',
        attachment.description || attachment.fallbackText || 'This content type cannot be rendered in the current experience.'
    )
}

function createImageAttachmentCard(attachment) {
    const figure = document.createElement('figure')
    figure.className = 'attachment-card attachment-card--image'

    const img = document.createElement('img')
    img.loading = 'lazy'
    img.decoding = 'async'
    img.alt = attachment.altText || attachment.title || 'Assistant shared an image'
    if (attachment.url) {
        img.src = attachment.url
    }

    img.addEventListener('error', () => {
        const placeholder = createAttachmentPlaceholderCard(
            attachment.title || 'Image unavailable',
            'The referenced image could not be loaded.'
        )
        figure.replaceWith(placeholder)
    })

    figure.appendChild(img)

    const captionText = attachment.title || attachment.description || ''
    if (captionText) {
        const caption = document.createElement('figcaption')
        caption.className = 'attachment-card__caption'
        caption.textContent = captionText
        figure.appendChild(caption)
    }

    if (attachment.url) {
        const footer = document.createElement('div')
        footer.className = 'attachment-card__footer'
        const link = document.createElement('a')
        link.href = attachment.url
        link.target = '_blank'
        link.rel = 'noopener noreferrer'
        link.textContent = ''
        footer.appendChild(link)
        figure.appendChild(footer)
    }

    return figure
}

function createImageAttachmentCarousel(images) {
    const carousel = document.createElement('div')
    carousel.className = 'assistant-attachment-carousel'
    carousel.dataset.index = '0'

    const markCarouselInteraction = () => {
        lastInteractionTime = new Date()
    }

    carousel.addEventListener('pointerdown', markCarouselInteraction)
    carousel.addEventListener('keydown', markCarouselInteraction)

    const track = document.createElement('div')
    track.className = 'assistant-attachment-carousel__track'
    carousel.appendChild(track)

    images.forEach((image, index) => {
        const slide = document.createElement('div')
        slide.className = 'assistant-attachment-carousel__slide'
        slide.dataset.index = String(index)
        if (index === 0) {
            slide.dataset.active = 'true'
        }
        slide.setAttribute('role', 'group')
        slide.setAttribute('aria-label', `Image ${index + 1} of ${images.length}`)

        const card = createImageAttachmentCard(image)
        card.classList.add('attachment-card--overlay')
        slide.appendChild(card)

        track.appendChild(slide)
    })

    if (images.length > 1) {
        const previous = document.createElement('button')
        previous.type = 'button'
        previous.className = 'assistant-attachment-carousel__control assistant-attachment-carousel__control--prev'
        previous.setAttribute('aria-label', 'Show previous image')
        previous.innerHTML = '<span aria-hidden="true">&lt;</span>'
        previous.addEventListener('click', () => {
            markCarouselInteraction()
            rotateCarousel(carousel, -1)
        })

        const next = document.createElement('button')
        next.type = 'button'
        next.className = 'assistant-attachment-carousel__control assistant-attachment-carousel__control--next'
        next.setAttribute('aria-label', 'Show next image')
        next.innerHTML = '<span aria-hidden="true">&gt;</span>'
        next.addEventListener('click', () => {
            markCarouselInteraction()
            rotateCarousel(carousel, 1)
        })

        track.appendChild(previous)
        track.appendChild(next)

        const indicators = document.createElement('div')
        indicators.className = 'assistant-attachment-carousel__indicators'

        images.forEach((_, index) => {
            const indicator = document.createElement('button')
            indicator.type = 'button'
            indicator.className = 'assistant-attachment-carousel__indicator'
            indicator.setAttribute('aria-label', `Show image ${index + 1}`)
            indicator.dataset.index = String(index)
            if (index === 0) {
                indicator.dataset.active = 'true'
            }
            indicator.addEventListener('click', () => {
                markCarouselInteraction()
                setCarouselActiveSlide(carousel, index)
            })
            indicators.appendChild(indicator)
        })

        carousel.appendChild(indicators)
    }

    return carousel
}

function rotateCarousel(carousel, offset) {
    const currentIndex = parseInt(carousel.dataset.index || '0', 10) || 0
    setCarouselActiveSlide(carousel, currentIndex + offset)
}

function setCarouselActiveSlide(carousel, index) {
    const slides = carousel.querySelectorAll('.assistant-attachment-carousel__slide')
    const total = slides.length
    if (!total) {
        return
    }

    const normalizedIndex = ((index % total) + total) % total
    carousel.dataset.index = String(normalizedIndex)

    slides.forEach((slide, slideIndex) => {
        if (slideIndex === normalizedIndex) {
            slide.dataset.active = 'true'
        } else {
            delete slide.dataset.active
        }
    })

    const indicators = carousel.querySelectorAll('.assistant-attachment-carousel__indicator')
    indicators.forEach((indicator, indicatorIndex) => {
        if (indicatorIndex === normalizedIndex) {
            indicator.dataset.active = 'true'
        } else {
            delete indicator.dataset.active
        }
    })
}

function createWidgetAttachmentCard(attachment) {
    const card = document.createElement('article')
    card.className = 'attachment-card attachment-card--widget'

    const header = document.createElement('div')
    header.className = 'attachment-card__header'

    const title = document.createElement('span')
    title.className = 'attachment-card__title'
    title.textContent = formatAttachmentTitle(attachment)
    header.appendChild(title)

    if (attachment.kind) {
        const badge = document.createElement('span')
        badge.className = 'attachment-card__badge'
        badge.textContent = capitalizeLabel(attachment.kind)
        header.appendChild(badge)
    }

    card.appendChild(header)

    const body = document.createElement('div')
    body.className = 'attachment-card__body'

    const description = attachment.description || attachment.fallbackText
    if (description) {
        const paragraph = document.createElement('p')
        paragraph.textContent = description
        body.appendChild(paragraph)
    }

    if (attachment.url) {
        const link = document.createElement('a')
        link.href = attachment.url
        link.target = '_blank'
        link.rel = 'noopener noreferrer'
        link.textContent = 'Open in new tab'
        body.appendChild(link)
    }

    const summary = summarizeAttachmentData(attachment.data)
    if (summary) {
        const pre = document.createElement('pre')
        pre.className = 'attachment-card__data'
        pre.textContent = summary
        body.appendChild(pre)
    }

    card.appendChild(body)
    return card
}

function createAttachmentPlaceholderCard(title, message) {
    const card = document.createElement('article')
    card.className = 'attachment-card attachment-card--placeholder'

    const heading = document.createElement('div')
    heading.className = 'attachment-card__title'
    heading.textContent = title || 'Content unavailable'
    card.appendChild(heading)

    if (message) {
        const body = document.createElement('div')
        body.className = 'attachment-card__body'
        body.textContent = message
        card.appendChild(body)
    }

    return card
}

function updateMicrophoneIcons(button) {
    if (!button) {
        return
    }

    const defaultIcon = button.querySelector('.icon-default')
    const activeIcon = button.querySelector('.icon-active')
    const speakingIcon = button.querySelector('.icon-speaking')
    const isAvatarSpeaking = button.dataset.avatarSpeaking === 'true'
    const isActive = button.dataset.state === 'active'

    if (speakingIcon) {
        if (isAvatarSpeaking) {
            if (defaultIcon) {
                defaultIcon.hidden = true
            }
            if (activeIcon) {
                activeIcon.hidden = true
            }
            speakingIcon.hidden = false
            return
        }
        speakingIcon.hidden = true
    }

    if (defaultIcon && activeIcon) {
        defaultIcon.hidden = isActive
        activeIcon.hidden = !isActive
    } else if (defaultIcon) {
        defaultIcon.hidden = false
    } else if (activeIcon) {
        activeIcon.hidden = !isActive
    }
}

function releaseMicrophoneAwaitingResponse(force = false) {
    const button = document.getElementById('microphone')
    if (!button) {
        return false
    }

    const awaitingResponse = button.dataset.awaitingResponse === 'true'
    const isActive = button.dataset.state === 'active'
    const isSpeaking = button.dataset.avatarSpeaking === 'true'
    const releaseReady = button.dataset.awaitingReleaseReady === 'true'

    if (!force && (!awaitingResponse || isActive || isSpeaking || !releaseReady)) {
        return false
    }

    if (force && isActive) {
        return false
    }

    delete button.dataset.awaitingResponse
    delete button.dataset.awaitingReleaseReady
    button.disabled = false
    updateButtonLabel(button, 'Start Microphone')
    updateMicrophoneIcons(button)
    return true
}

function engageMicrophoneAwaitingResponse() {
    const button = document.getElementById('microphone')
    if (!button) {
        return
    }

    button.dataset.awaitingResponse = 'true'
    delete button.dataset.awaitingReleaseReady
    button.disabled = true
    updateButtonLabel(button, 'Processing Response')
    updateMicrophoneIcons(button)
}

function lockMicrophoneWhileSpeaking() {
    const button = document.getElementById('microphone')
    if (!button) {
        return
    }
    if (button.dataset.state === 'active') {
        return
    }
    if (typeof button.dataset.preSpeakingDisabled === 'undefined') {
        button.dataset.preSpeakingDisabled = button.disabled ? 'true' : 'false'
    }
    button.dataset.avatarSpeaking = 'true'
    button.disabled = false
    updateButtonLabel(button, 'Stop Speaking')
    updateMicrophoneIcons(button)
}

function unlockMicrophoneAfterSpeaking() {
    const button = document.getElementById('microphone')
    if (!button) {
        return
    }
    const awaitingResponse = button.dataset.awaitingResponse === 'true'

    delete button.dataset.avatarSpeaking

    if (button.dataset.state === 'active') {
        delete button.dataset.preSpeakingDisabled
        button.disabled = false
        updateButtonLabel(button, 'Stop Microphone')
        updateMicrophoneIcons(button)
        return
    }

    if (typeof button.dataset.preSpeakingDisabled !== 'undefined') {
        const wasDisabled = button.dataset.preSpeakingDisabled === 'true'
        button.disabled = awaitingResponse ? true : wasDisabled
        delete button.dataset.preSpeakingDisabled
    }

    if (awaitingResponse) {
        const releaseReady = button.dataset.awaitingReleaseReady === 'true'
        if (releaseReady) {
            releaseMicrophoneAwaitingResponse(true)
        } else {
            updateButtonLabel(button, 'Processing Response')
            updateMicrophoneIcons(button)
        }
        return
    }

    updateButtonLabel(button, 'Start Microphone')
    updateMicrophoneIcons(button)
}

function setMicrophoneState(state) {
    const button = document.getElementById('microphone')
    if (!button) {
        return
    }

    button.dataset.state = state

    const label = state === 'active' ? 'Stop Microphone' : 'Start Microphone'
    updateButtonLabel(button, label)

    if (state === 'active') {
        delete button.dataset.awaitingResponse
        delete button.dataset.awaitingReleaseReady
        delete button.dataset.avatarSpeaking
        button.disabled = false
        delete button.dataset.preSpeakingDisabled
    } else if (!isSpeaking && typeof button.dataset.preSpeakingDisabled !== 'undefined') {
        unlockMicrophoneAfterSpeaking()
        return
    }

    updateMicrophoneIcons(button)
}

function ensureUserMessageInput() {
    const messageBox = document.getElementById('userMessageBox')
    if (!messageBox) {
        return
    }

    const deferDisplay = messageBox.dataset.deferredDisplay === 'true'
    if (!deferDisplay) {
        messageBox.hidden = false
    }

    if (messageBox.dataset.listenerAttached === 'true') {
        const controls = document.getElementById('chatControls')
        if (controls && !controls.hidden && !messageBox.hidden) {
            requestAnimationFrame(() => messageBox.focus())
        }
        return
    }

    messageBox.dataset.listenerAttached = 'true'
    messageBox.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter' || e.shiftKey) {
            return
        }

        e.preventDefault()
        const messageBoxEl = e.currentTarget
        const userQuery = messageBoxEl.innerText.trim()
        const childImg = messageBoxEl.querySelector('#picInput')
        if (childImg) {
            childImg.style.width = '200px'
            childImg.style.height = '200px'
        }

        let userQueryHTML = messageBoxEl.innerHTML.trim()
        if (userQueryHTML.startsWith('<img')) {
            userQueryHTML = '<br/>' + userQueryHTML
        }

        if (userQuery !== '') {
            handleUserQuery(userQuery, userQueryHTML, imgUrl)
            messageBoxEl.innerHTML = ''
            imgUrl = ''
            const controls = document.getElementById('chatControls')
            const toggleButton = document.getElementById('textInputToggle')
            const backdrop = document.getElementById('textInputBackdrop')
            const mobileToggle = document.getElementById('mobileTextInputToggle')

            if (controls && controls.classList.contains('floating-mode') && controls.classList.contains('show-input')) {
                controls.classList.remove('show-input')
                messageBoxEl.hidden = true
                messageBoxEl.dataset.deferredDisplay = 'true'
                messageBoxEl.classList.remove('floating-input')
                messageBoxEl.blur()
                if (backdrop) {
                    backdrop.hidden = true
                }
                if (toggleButton && toggleButton.hidden === false) {
                    toggleButton.setAttribute('aria-pressed', 'false')
                    try {
                        toggleButton.focus({ preventScroll: true })
                    } catch (_) {
                        toggleButton.focus()
                    }
                }
            }

            if (controls && !controls.classList.contains('floating-mode') && isSmallScreenMobileLayout() && controls.classList.contains('show-input')) {
                controls.classList.remove('show-input')
                messageBoxEl.hidden = true
                messageBoxEl.dataset.deferredDisplay = 'true'
                messageBoxEl.classList.remove('floating-input')
                messageBoxEl.blur()
                if (typeof window.closeMobileActionsMenu === 'function') {
                    window.closeMobileActionsMenu(false)
                }
                if (mobileToggle) {
                    mobileToggle.setAttribute('aria-pressed', 'false')
                    try {
                        mobileToggle.focus({ preventScroll: true })
                    } catch (_) {
                        mobileToggle.focus()
                    }
                }
            }
        }
    })

    const controls = document.getElementById('chatControls')
    if (controls && !controls.hidden && !messageBox.hidden) {
        requestAnimationFrame(() => messageBox.focus())
    }
}

const largePortraitMediaQuery = typeof window !== 'undefined'
    ? window.matchMedia('(orientation: portrait) and (min-width: 900px)')
    : null

const smallScreenMobileActionsQuery = typeof window !== 'undefined'
    ? window.matchMedia('(max-width: 600px)')
    : null

function isSmallScreenMobileLayout() {
    return smallScreenMobileActionsQuery ? smallScreenMobileActionsQuery.matches : false
}

let applyTextInputToggleLayout = null

function setupTextInputToggle() {
    if (!largePortraitMediaQuery) {
        return
    }

    const toggleButton = document.getElementById('textInputToggle')
    const controls = document.getElementById('chatControls')
    const messageBox = document.getElementById('userMessageBox')
    const backdrop = document.getElementById('textInputBackdrop')
    if (!toggleButton || !controls || !messageBox) {
        return
    }

    let previousIsLargePortrait = null

    const hideFloatingInput = (focusToggle = false) => {
        controls.classList.remove('show-input')
        toggleButton.setAttribute('aria-pressed', 'false')
        messageBox.hidden = true
        messageBox.classList.remove('floating-input')
        messageBox.dataset.deferredDisplay = 'true'
        messageBox.blur()
        if (backdrop) {
            backdrop.hidden = true
        }
        if (focusToggle) {
            try {
                toggleButton.focus({ preventScroll: true })
            } catch (_) {
                toggleButton.focus()
            }
        }
    }

    const showFloatingInput = () => {
        controls.classList.add('show-input')
        toggleButton.setAttribute('aria-pressed', 'true')
        messageBox.hidden = false
        messageBox.classList.add('floating-input')
        messageBox.dataset.deferredDisplay = 'false'
        if (backdrop) {
            backdrop.hidden = false
        }
        ensureUserMessageInput()
    }

    const applyLayoutState = (resetVisibility = false) => {
        const isLargePortrait = largePortraitMediaQuery.matches
        const isSmallMobile = isSmallScreenMobileLayout()

        controls.classList.toggle('floating-mode', isLargePortrait)

        if (isLargePortrait) {
            toggleButton.hidden = false
            if (resetVisibility || previousIsLargePortrait !== true) {
                hideFloatingInput(false)
            }

            if (controls.classList.contains('show-input')) {
                showFloatingInput()
            } else {
                hideFloatingInput(false)
            }
        } else {
            toggleButton.hidden = true
            toggleButton.setAttribute('aria-pressed', 'false')
            controls.classList.remove('floating-mode')
            messageBox.classList.remove('floating-input')
            if (isSmallMobile) {
                if (resetVisibility) {
                    controls.classList.remove('show-input')
                }
                const shouldShow = controls.classList.contains('show-input')
                messageBox.hidden = !shouldShow
                messageBox.dataset.deferredDisplay = shouldShow ? 'false' : 'true'
                if (shouldShow) {
                    ensureUserMessageInput()
                } else {
                    messageBox.blur()
                }
                if (backdrop) {
                    backdrop.hidden = true
                }
            } else {
                controls.classList.add('show-input')
                if (messageBox.hidden) {
                    messageBox.hidden = false
                }
                messageBox.dataset.deferredDisplay = 'false'
                if (backdrop) {
                    backdrop.hidden = true
                }
                ensureUserMessageInput()
            }
        }

        previousIsLargePortrait = isLargePortrait
    }

    const handleToggleClick = () => {
        if (!largePortraitMediaQuery.matches) {
            return
        }

        const willShow = !controls.classList.contains('show-input')
        if (willShow) {
            showFloatingInput()
        } else {
            messageBox.blur()
            hideFloatingInput(false)
        }
    }

    toggleButton.addEventListener('click', handleToggleClick)

    if (backdrop) {
        backdrop.addEventListener('click', () => {
            if (!largePortraitMediaQuery.matches) {
                return
            }
            if (controls.classList.contains('show-input')) {
                hideFloatingInput(true)
            }
        })
    }

    document.addEventListener('keydown', (event) => {
        if (event.key !== 'Escape') {
            return
        }
        if (!largePortraitMediaQuery.matches) {
            return
        }
        if (!controls.classList.contains('show-input')) {
            return
        }
        hideFloatingInput(true)
    })

    const registerMediaListener = (mediaQuery, callback) => {
        if (!mediaQuery) {
            return
        }
        if (typeof mediaQuery.addEventListener === 'function') {
            mediaQuery.addEventListener('change', callback)
        } else if (typeof mediaQuery.addListener === 'function') {
            mediaQuery.addListener(callback)
        }
    }

    applyTextInputToggleLayout = (resetVisibility = false) => applyLayoutState(resetVisibility)

    applyLayoutState(true)
    registerMediaListener(largePortraitMediaQuery, () => applyLayoutState(true))
}

function invokeOptionalCallback(callbackName, ...args) {
    const candidate = window[callbackName]
    if (typeof candidate !== 'function') {
        return
    }
    try {
        candidate(...args)
    } catch (err) {
        console.warn(`Optional callback ${callbackName} failed`, err)
    }
}

async function ensureAvatarConnectionReady() {
    if (avatarSynthesizer && typeof avatarSynthesizer.speakSsmlAsync === 'function') {
        return true
    }

    if (pendingAvatarConnection) {
        return pendingAvatarConnection
    }

    pendingAvatarConnection = (async () => {
        try {
            return await connectAvatar()
        } finally {
            pendingAvatarConnection = null
        }
    })()

    return pendingAvatarConnection
}

function setupMobileSessionActions() {
    const sessionButtons = document.getElementById('sessionButtons')
    const moreButton = document.getElementById('moreActions')
    const secondaryActions = document.getElementById('sessionSecondaryActions')

    const controls = document.getElementById('chatControls')
    const messageBox = document.getElementById('userMessageBox')
    const mobileToggle = document.getElementById('mobileTextInputToggle')

    if (!sessionButtons || !moreButton || !secondaryActions || !controls || !messageBox || !mobileToggle) {
        return
    }

    const isMobileView = () => smallScreenMobileActionsQuery ? smallScreenMobileActionsQuery.matches : false

    const focusButtonSafely = (button) => {
        if (!button) {
            return
        }
        try {
            button.focus({ preventScroll: true })
        } catch (_) {
            button.focus()
        }
    }

    const updateMobileToggleState = (pressed) => {
        mobileToggle.setAttribute('aria-pressed', pressed ? 'true' : 'false')
    }

    const showMobileInput = ({ focus = true } = {}) => {
        if (!isMobileView()) {
            return
        }
        controls.classList.add('show-input')
        messageBox.hidden = false
        messageBox.classList.remove('floating-input')
        messageBox.dataset.deferredDisplay = 'false'
        updateMobileToggleState(true)
        if (focus) {
            ensureUserMessageInput()
        }
    }

    const hideMobileInput = ({ focusToggle = false } = {}) => {
        if (!isMobileView()) {
            return
        }
        controls.classList.remove('show-input')
        messageBox.hidden = true
        messageBox.classList.remove('floating-input')
        messageBox.dataset.deferredDisplay = 'true'
        messageBox.blur()
        updateMobileToggleState(false)
        if (focusToggle) {
            focusButtonSafely(mobileToggle)
        }
    }

    const closeMenu = (focusToggle = false) => {
        if (!sessionButtons.classList.contains('mobile-actions-open')) {
            if (focusToggle) {
                focusButtonSafely(moreButton)
            }
            return
        }

        sessionButtons.classList.remove('mobile-actions-open')
        moreButton.setAttribute('aria-expanded', 'false')
        updateButtonLabel(moreButton, 'More Actions')
        document.body.classList.remove('mobile-actions-menu-open')

        if (focusToggle) {
            focusButtonSafely(moreButton)
        }
    }

    const openMenu = () => {
        if (!isMobileView()) {
            return
        }
        if (sessionButtons.classList.contains('mobile-actions-open')) {
            return
        }

        sessionButtons.classList.add('mobile-actions-open')
        moreButton.setAttribute('aria-expanded', 'true')
        updateButtonLabel(moreButton, 'Close Actions')
        document.body.classList.add('mobile-actions-menu-open')

        updateMobileToggleState(controls.classList.contains('show-input'))

        const firstInteractive = secondaryActions.querySelector('button:not([disabled])') || secondaryActions.querySelector('button')
        if (firstInteractive) {
            requestAnimationFrame(() => focusButtonSafely(firstInteractive))
        }
    }

    const toggleMenu = () => {
        if (!isMobileView()) {
            return
        }
        if (sessionButtons.classList.contains('mobile-actions-open')) {
            closeMenu(true)
        } else {
            openMenu()
        }
    }

    moreButton.addEventListener('click', (event) => {
        event.preventDefault()
        toggleMenu()
    })

    moreButton.addEventListener('keydown', (event) => {
        if (event.key === 'ArrowDown' && isMobileView()) {
            event.preventDefault()
            openMenu()
        }
    })

    mobileToggle.addEventListener('click', (event) => {
        event.preventDefault()
        if (!isMobileView()) {
            return
        }
        const isShowing = controls.classList.contains('show-input')
        if (isShowing) {
            hideMobileInput({ focusToggle: false })
        } else {
            showMobileInput({ focus: true })
        }
        closeMenu(false)
    })

    secondaryActions.addEventListener('click', (event) => {
        const button = event.target instanceof HTMLElement ? event.target.closest('button') : null
        if (!button) {
            return
        }
        if (!isMobileView()) {
            return
        }
        // Defer closing so that the button's action runs first
        setTimeout(() => closeMenu(false), 0)
    })

    document.addEventListener('keydown', (event) => {
        if (event.key !== 'Escape') {
            return
        }
        if (!sessionButtons.classList.contains('mobile-actions-open')) {
            return
        }
        event.preventDefault()
        closeMenu(true)
    })

    const handleViewportChange = () => {
        if (!isMobileView()) {
            closeMenu(false)
            moreButton.setAttribute('aria-expanded', 'false')
            updateButtonLabel(moreButton, 'More Actions')
            document.body.classList.remove('mobile-actions-menu-open')
            updateMobileToggleState(false)
            if (typeof applyTextInputToggleLayout === 'function') {
                applyTextInputToggleLayout(true)
            }
        } else {
            const shouldShow = controls.classList.contains('show-input')
            messageBox.hidden = !shouldShow
            messageBox.classList.remove('floating-input')
            messageBox.dataset.deferredDisplay = shouldShow ? 'false' : 'true'
            if (!shouldShow) {
                messageBox.blur()
            }
            updateMobileToggleState(shouldShow)
        }
    }

    if (smallScreenMobileActionsQuery) {
        if (typeof smallScreenMobileActionsQuery.addEventListener === 'function') {
            smallScreenMobileActionsQuery.addEventListener('change', handleViewportChange)
        } else if (typeof smallScreenMobileActionsQuery.addListener === 'function') {
            smallScreenMobileActionsQuery.addListener(handleViewportChange)
        }
    }

    handleViewportChange()

    window.closeMobileActionsMenu = closeMenu
}

window.copyChatHistory = async () => {
    lastInteractionTime = new Date()
    const chatHistory = document.getElementById('chatHistory')
    if (!chatHistory) {
        return
    }

    const walker = document.createTreeWalker(chatHistory, NodeFilter.SHOW_TEXT, null)
    const segments = []
    while (walker.nextNode()) {
        const textContent = walker.currentNode.nodeValue
        if (textContent) {
            segments.push(textContent)
        }
    }

    const plainText = segments.join('\n').replace(/\n{3,}/g, '\n\n').trim()

    try {
        if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(plainText)
            if (typeof window !== 'undefined' && window.DEBUG_CHAT_HISTORY === true) {
                console.log('[ChatHistory] Copied via clipboard API:', plainText)
            }
        } else {
            throw new Error('Clipboard API not available')
        }
    } catch (error) {
        const textarea = document.createElement('textarea')
        textarea.value = plainText
        textarea.setAttribute('readonly', '')
        textarea.style.position = 'absolute'
        textarea.style.left = '-9999px'
        document.body.appendChild(textarea)
        textarea.select()
        try {
            document.execCommand('copy')
            if (typeof window !== 'undefined' && window.DEBUG_CHAT_HISTORY === true) {
                console.log('[ChatHistory] Copied via execCommand fallback:', plainText)
            }
        } catch (fallbackError) {
            console.error('Failed to copy chat history:', fallbackError)
        }
        document.body.removeChild(textarea)
    }
}

