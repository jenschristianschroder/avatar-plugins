# Avatar Session Lifecycle

This document explains the functions and events used to manage the Azure AI Avatar video stream, including initialization, speaking control, idle fallback, and session termination.

## Table of Contents

- [Session Initialization](#session-initialization)
- [Speaking Control](#speaking-control)
- [Idle Video Management](#idle-video-management)
- [Session Termination](#session-termination)
- [WebRTC Events](#webrtc-events)

---

## Session Initialization

### `connectAvatar()`

**Purpose**: Establishes a connection to Azure Speech Services and initializes the avatar session with WebRTC.

**Location**: `js/chat-agent.js` (line 3064)

**Returns**: `Promise<boolean>` - `true` if connection succeeds, `false` otherwise

**Process**:

1. **Configuration Retrieval**
   - Gets speech region, private endpoint settings, and authentication
   - Retrieves avatar character, style, and background settings
   - Obtains agent thread configuration

2. **Speech SDK Initialization**
   - Creates `SpeechConfig` with authorization token
   - Configures TTS endpoint (WebSocket connection)
   - Sets up custom voice endpoint if specified

3. **Avatar Configuration**
   - Creates `AvatarConfig` with character and style
   - Configures background (image or transparent)
   - Sets video codec (`vp9` for transparent background, default otherwise)

4. **Avatar Synthesizer Setup**
   - Instantiates `SpeechSDK.AvatarSynthesizer`
   - Registers `avatarEventReceived` callback for WebRTC events

5. **Speech Recognizer Setup** (for voice input)
   - Creates `SpeechRecognizer` with continuous language detection
   - Configures STT endpoint and supported locales

6. **WebRTC Connection**
   - Requests session from services proxy (ICE server credentials)
   - Calls `setupWebRTC()` to establish peer connection

**Usage**:

```javascript
const success = await connectAvatar();
if (success) {
    console.log('Avatar session initialized successfully');
}
```

**Error Handling**:
- Shows alert on failure
- Re-enables start session button
- Logs detailed error messages

---

### `setupWebRTC(iceServerUrl, iceServerUsername, iceServerCredential, agentThreadConfig)`

**Purpose**: Establishes the WebRTC peer connection for avatar video/audio streaming.

**Location**: `js/chat-agent.js` (line 3192)

**Parameters**:
- `iceServerUrl`: TURN/STUN server URL
- `iceServerUsername`: Authentication username
- `iceServerCredential`: Authentication password
- `agentThreadConfig`: Agent configuration object

**Process**:

1. **Peer Connection Creation**
   - Creates `RTCPeerConnection` with ICE server configuration
   - Adds transceivers for audio and video (sendrecv direction)

2. **Track Handling** (`ontrack` event)
   - **Audio Track**: Creates `<audio>` element, enables autoplay
   - **Video Track**: Creates `<video>` element, sets to playsInline
   - Appends elements to `#remoteVideo` container

3. **Data Channel Setup**
   - Listens for `datachannel` event
   - Handles WebRTC events via data channel messages

4. **ICE Connection State Monitoring**
   - `disconnected`: Shows local idle video
   - `connected`/`completed`: Hides local idle video

5. **Avatar Start**
   - Calls `avatarSynthesizer.startAvatarAsync(peerConnection)`
   - Handles success/failure with appropriate UI updates

**Data Channel Message Handling**:
- Receives JSON messages with event type and metadata
- Updates UI based on avatar state (speaking, idle, session end)

---

## Speaking Control

### `speakNext(text, endingSilenceMs = 0, skipUpdatingChatHistory = false)`

**Purpose**: Synthesizes speech and makes the avatar speak the provided text.

**Location**: `js/chat-agent.js` (line 3488)

**Parameters**:
- `text`: Text to be spoken
- `endingSilenceMs`: Optional silence duration after speech (milliseconds)
- `skipUpdatingChatHistory`: Whether to skip chat history update

**Process**:

1. **Text Preparation**
   - Strips inline citations from text
   - Sanitizes content for SSML

2. **SSML Generation**
   - Creates SSML markup with selected TTS voice
   - Adds leading silence (`<mstts:leadingsilence-exact value='0'/>`)
   - Adds optional ending silence (`<break time='...ms' />`)

3. **Speech Synthesis**
   - Calls `avatarSynthesizer.speakSsmlAsync(ssml)`
   - Sets `isSpeaking = true`
   - Locks microphone during speaking

4. **Queue Management**
   - If speech queue (`spokenTextQueue`) has items, processes next
   - Otherwise, unlocks microphone

**Usage**:

```javascript
speakNext("Hello, how can I help you today?");
speakNext("Please wait.", 500); // with 500ms pause
```

**State Management**:
- `speakingText`: Tracks currently speaking text
- `isSpeaking`: Boolean flag for speaking state
- `lastSpeakTime`: Timestamp of last speech

---

### `stopSpeaking()`

**Purpose**: Immediately stops the avatar from speaking and clears the speech queue.

**Location**: `js/chat-agent.js` (line 3545)

**Process**:

1. **Queue Clearing**
   - Empties `spokenTextQueue` array
   - Updates `lastInteractionTime`

2. **Microphone Handling**
   - Temporarily disables microphone button
   - Removes speaking lock

3. **Avatar Stop**
   - Calls `avatarSynthesizer.stopSpeakingAsync()`
   - Sets `isSpeaking = false`
   - Unlocks microphone after completion

**Usage**:

```javascript
stopSpeaking(); // Interrupt current speech
```

**Exposed Globally**:
```javascript
window.stopSpeaking = stopSpeaking; // Available in browser console
```

---

## Idle Video Management

### `showLocalIdleVideo()`

**Purpose**: Displays the local idle video (fallback when avatar is not actively speaking).

**Location**: `js/chat-agent.js` (line 1059)

**Process**:

1. **Checks if Enabled**
   - Returns early if `isLocalIdleVideoEnabled()` returns `false`

2. **Show Local Video**
   - Unhides `#localVideo` element
   - Sets dimensions to full viewport (`100%` width, `100vh` height)
   - Sets visibility to `visible`

3. **Hide Remote Video**
   - Hides `#remoteVideo` container (WebRTC stream)

**When Called**:
- Avatar switches to idle state (`EVENT_TYPE_SWITCH_TO_IDLE`)
- Avatar session ends (`EVENT_TYPE_SESSION_END`)
- WebRTC connection is disconnected
- Agent streaming completes or errors

**Usage**:

```javascript
showLocalIdleVideo(); // Switch to idle state
```

---

### `hideLocalIdleVideo()`

**Purpose**: Hides the local idle video and shows the WebRTC avatar stream.

**Location**: `js/chat-agent.js` (line 1077)

**Process**:

1. **Hide Local Video**
   - Hides `#localVideo` element
   - Sets dimensions to zero (`0px` width, `0vh` height)
   - Sets visibility to `hidden`

2. **Show Remote Video**
   - Unhides `#remoteVideo` container
   - Resets width and visibility

**When Called**:
- Avatar starts speaking (`EVENT_TYPE_TURN_START`)
- WebRTC connection is established (`connected`/`completed`)
- Video track is received

**Usage**:

```javascript
hideLocalIdleVideo(); // Switch to active avatar
```

---

## Session Termination

### `disconnectAvatar()`

**Purpose**: Closes the avatar synthesizer and speech recognizer, ending the session.

**Location**: `js/chat-agent.js` (line 3177)

**Process**:

1. **Avatar Synthesizer Cleanup**
   - Calls `avatarSynthesizer.close()`
   - Sets `avatarSynthesizer = undefined`

2. **Speech Recognizer Cleanup**
   - Calls `speechRecognizer.stopContinuousRecognitionAsync()`
   - Calls `speechRecognizer.close()`
   - Sets `speechRecognizer = undefined`

3. **Session State**
   - Sets `sessionActive = false`

**Usage**:

```javascript
disconnectAvatar(); // Clean up and end session
```

**Note**: Does not close WebRTC peer connection directly. Connection cleanup happens via ICE state changes and session end events.

---

## WebRTC Events

The avatar service sends events via the WebRTC data channel. These events are received in the `datachannel.onmessage` handler within `setupWebRTC()`.

### Event Structure

```json
{
  "event": {
    "eventType": "EVENT_TYPE_...",
    "offset": "14.610s",
    "duration": "0s",
    "turnID": "835557A750FB4AC0981ADA59006A5B48"
  }
}
```

### Event Types

#### `EVENT_TYPE_TURN_START`

**When**: Avatar starts speaking

**Actions**:
- Calls `hideLocalIdleVideo()` - switches to WebRTC stream
- Shows subtitles if enabled
- Updates subtitles with current speaking text

**Example**:
```json
{
  "event": {
    "eventType": "EVENT_TYPE_TURN_START",
    "offset": "14.610s",
    "turnID": "835557A750FB4AC0981ADA59006A5B48"
  }
}
```

---

#### `EVENT_TYPE_SWITCH_TO_IDLE`

**When**: Avatar finishes speaking and enters idle state

**Actions**:
- Hides subtitles
- Calls `showLocalIdleVideo()` - switches to local idle video

**Example**:
```json
{
  "event": {
    "eventType": "EVENT_TYPE_SWITCH_TO_IDLE",
    "offset": "27.430s"
  }
}
```

---

#### `EVENT_TYPE_SESSION_END`

**When**: Avatar session terminates (timeout, error, or explicit closure)

**Actions**:
- Hides subtitles
- Calls `showLocalIdleVideo()`
- Logs session end
- Calls `transitionToPreSessionState()` to reset UI

**Example**:
```json
{
  "event": {
    "eventType": "EVENT_TYPE_SESSION_END",
    "offset": "600.000s"
  }
}
```

---

#### `EVENT_TYPE_SWITCH_TO_SPEAKING`

**When**: Avatar transitions from idle to speaking (accompanies `EVENT_TYPE_TURN_START`)

**Actions**:
- Logged for debugging
- No direct UI changes (handled by `EVENT_TYPE_TURN_START`)

---

### ICE Connection State Events

The `RTCPeerConnection.oniceconnectionstatechange` event monitors WebRTC connection health:

#### `disconnected`

**When**: Network interruption or connection loss

**Actions**:
- Calls `showLocalIdleVideo()` - fallback to idle

#### `connected` / `completed`

**When**: WebRTC connection established successfully

**Actions**:
- Calls `hideLocalIdleVideo()` - show avatar stream

---

## Complete Lifecycle Example

### Starting a Session

```javascript
// 1. User clicks "Talk to Avatar" button
document.getElementById('startSession').onclick = async () => {
    // 2. Initialize avatar connection
    const connected = await connectAvatar();
    
    if (connected) {
        // 3. Avatar starts and WebRTC is established
        // 4. Local idle video is hidden, remote stream is shown
        console.log('Avatar session active');
    }
};
```

### Speaking Flow

```javascript
// 1. Agent generates response
const agentResponse = "Hello! How can I help you today?";

// 2. Trigger avatar speech
speakNext(agentResponse);

// 3. Events received:
//    - EVENT_TYPE_TURN_START -> hideLocalIdleVideo()
//    - (avatar speaks)
//    - EVENT_TYPE_SWITCH_TO_IDLE -> showLocalIdleVideo()
```

### Interrupting Speech

```javascript
// User clicks stop button
document.getElementById('stopButton').onclick = () => {
    stopSpeaking(); // Immediately stops avatar
};
```

### Ending Session

```javascript
// Session timeout or user exits
disconnectAvatar();

// WebRTC connection closes
// EVENT_TYPE_SESSION_END received
// UI transitions to pre-session state
```

---

## Best Practices

1. **Always await `connectAvatar()`** before attempting to speak
2. **Use `stopSpeaking()`** before disconnecting to prevent audio glitches
3. **Monitor WebRTC events** for connection health
4. **Handle `EVENT_TYPE_SESSION_END`** to gracefully reset UI
5. **Check `isSpeaking` flag** before queuing new speech
6. **Use `spokenTextQueue`** for multiple sequential utterances

---

## Troubleshooting

### Avatar Not Starting

- Verify speech token is valid
- Check WebRTC ICE server credentials
- Ensure avatar character/style is supported
- Review browser console for connection errors

### Video Not Displaying

- Confirm `#remoteVideo` element exists
- Check video track is received (`ontrack` event)
- Verify `hideLocalIdleVideo()` is called
- Inspect ICE connection state

### Speech Not Playing

- Ensure `avatarSynthesizer` is initialized
- Check SSML is valid XML
- Verify TTS voice is supported
- Monitor `speakSsmlAsync()` promise resolution

---

## Related Files

- `js/chat-agent.js` - Core avatar session logic
- `shared/pluginBase.js` - Plugin integration for overlay content
- `docs/interaction-sequence.md` - Detailed interaction flow diagrams

---

## API References

- [Azure Speech SDK Documentation](https://learn.microsoft.com/azure/ai-services/speech-service/)
- [WebRTC API](https://developer.mozilla.org/docs/Web/API/WebRTC_API)
- [RTCPeerConnection](https://developer.mozilla.org/docs/Web/API/RTCPeerConnection)
- [Azure AI Avatar](https://learn.microsoft.com/azure/ai-services/speech-service/avatar)
