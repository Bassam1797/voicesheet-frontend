// Local microphone implementation without using the Web Speech API.
// This module exposes a single function `wireMicrophoneControls()` on
// the global `window` object. It attaches a click handler to an element
// with the ID `micBtn`, toggling recording from the user's microphone.
// When recording stops, it will send the captured audio to a local STT
// endpoint defined by `window.LOCAL_STT_ENDPOINT`. If defined, it
// expects the endpoint to respond with JSON containing a `text` field.
// After receiving a transcription, it calls a global callback
// `window.onMicResult(text)` if provided.

export function wireMicrophoneControls() {
  const micBtn = document.getElementById('micBtn');
  if (!micBtn) return;
    // Set default local transcription endpoint if not defined
    if (!window.LOCAL_STT_ENDPOINT) {

    window.LOCAL_STT_ENDPOINT = 'http://localhost:8000/api/transcribe';
  }

  let recording = false;
  let mediaRecorder;
  let audioChunks = [];

  async function startRecording() {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      mediaRecorder = new MediaRecorder(stream);
      audioChunks = [];
      mediaRecorder.ondataavailable = (event) => {
        if (event.data && event.data.size > 0) {
          audioChunks.push(event.data);
        }
      };
      mediaRecorder.onstop = async () => {
        // Combine recorded chunks into a single Blob
        const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
        if (window.LOCAL_STT_ENDPOINT) {
          // Send the blob to the configured STT endpoint
          const formData = new FormData();
          formData.append('audio', audioBlob, 'speech.webm');
          try {
            const response = await fetch(window.LOCAL_STT_ENDPOINT, {
              method: 'POST',
              body: formData,
            });
            const data = await response.json();
            if (data && data.text && typeof window.onMicResult === 'function') {
              window.onMicResult(data.text);
            }
          } catch (err) {
            console.error('Error sending audio to STT endpoint:', err);
          }
        }
      };
      mediaRecorder.start();
      micBtn.classList.add('listening');
    } catch (err) {
      console.error('Microphone error:', err);
    }
  }

  function stopRecording() {
    if (mediaRecorder) {
      mediaRecorder.stop();
      micBtn.classList.remove('listening');
    }
  }

  micBtn.addEventListener('click', () => {
    if (recording) {
      stopRecording();
    } else {
      startRecording();
    }
    recording = !recording;
  });
}

// Expose function on window for legacy code integration
window.wireMicrophoneControls = wireMicrophoneControls;
