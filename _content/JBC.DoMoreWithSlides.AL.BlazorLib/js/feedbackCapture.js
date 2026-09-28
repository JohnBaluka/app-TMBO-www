// Screen capture and control picking for the Feedback tab.
//
// Everything here is browser-only and permission-gated: getDisplayMedia shows the browser's own
// picker, and there is no way to capture anything without the user choosing what to share. That is
// deliberate and not worked around - a "silent screenshot" is not a thing this application should
// be able to do.
//
// Every entry point resolves rather than throws. A head with no getDisplayMedia (an Office web
// add-in pane, an older WebView) gets a disabled button with a reason, not a broken panel.

let recorder = null;
let recorderChunks = [];
let recorderStream = null;
let picker = null;

/** The import probe every module in this repo answers, so a stubbed runtime is detectable. */
export function ready() {
    return true;
}

/** Whether this browser can capture the screen at all. */
export function canCapture() {
    return !!(navigator.mediaDevices && navigator.mediaDevices.getDisplayMedia);
}

/** Whether it can record as well - a codec and MediaRecorder are both needed. */
export function canRecord() {
    return canCapture() && typeof MediaRecorder !== 'undefined';
}

function stopStream(stream) {
    if (!stream) return;

    for (const track of stream.getTracks()) {
        try { track.stop(); } catch { /* already ended */ }
    }
}

/**
 * One frame of whatever the user chose to share, as a PNG data URI.
 *
 * The stream is stopped the moment the frame is drawn: leaving it running would leave the
 * browser's "sharing your screen" indicator up over a screenshot that was already taken.
 */
export async function takeScreenshot() {
    if (!canCapture()) return null;

    let stream = null;

    try {
        stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false });

        const video = document.createElement('video');

        video.srcObject = stream;
        video.muted = true;

        await video.play();

        // One animation frame, so the first painted frame is the one captured rather than a black
        // one - a video element reports itself playing before it has anything to show.
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));

        const canvas = document.createElement('canvas');

        canvas.width = video.videoWidth || 1280;
        canvas.height = video.videoHeight || 720;

        canvas.getContext('2d').drawImage(video, 0, 0, canvas.width, canvas.height);

        video.pause();
        video.srcObject = null;

        return canvas.toDataURL('image/png');
    } catch {
        // The user cancelled the picker, or the browser refused. Neither is an error worth a dialog.
        return null;
    } finally {
        stopStream(stream);
    }
}

/** True while a recording is running. */
export function isRecording() {
    return !!recorder && recorder.state === 'recording';
}

/**
 * Starts recording whatever the user chose to share.
 *
 * Returns false when the picker was cancelled, which is not a failure.
 */
export async function startRecording() {
    if (!canRecord() || isRecording()) return false;

    try {
        recorderStream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: true });
    } catch {
        recorderStream = null;

        return false;
    }

    recorderChunks = [];

    // Let the browser choose the container it actually supports. Chromium gives webm; naming a
    // codec it does not have makes the constructor throw.
    try {
        recorder = new MediaRecorder(recorderStream);
    } catch {
        stopStream(recorderStream);

        recorderStream = null;
        recorder = null;

        return false;
    }

    recorder.ondataavailable = event => {
        if (event.data && event.data.size > 0) recorderChunks.push(event.data);
    };

    // Stopping the share from the browser's own bar ends the track rather than the recorder, so the
    // recorder is told as well - without this, Stop in the panel would wait for data that never
    // arrives.
    for (const track of recorderStream.getVideoTracks()) {
        track.addEventListener('ended', () => {
            if (recorder && recorder.state === 'recording') recorder.stop();
        });
    }

    recorder.start();

    return true;
}

/**
 * Stops the recording and hands back what it captured, as a data URI.
 *
 * The blob is read through a FileReader rather than assembled by hand: a recording is megabytes,
 * and base64 built in JS string concatenation is what makes a browser tab stop responding.
 */
export async function stopRecording() {
    if (!recorder) return null;

    const finished = new Promise(resolve => {
        recorder.onstop = () => resolve();
    });

    try {
        if (recorder.state === 'recording') recorder.stop();

        await finished;
    } catch {
        return null;
    } finally {
        stopStream(recorderStream);

        recorderStream = null;
    }

    const type = recorder.mimeType || 'video/webm';

    recorder = null;

    if (recorderChunks.length === 0) return null;

    const blob = new Blob(recorderChunks, { type });

    recorderChunks = [];

    return await new Promise(resolve => {
        const reader = new FileReader();

        reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : null);
        reader.onerror = () => resolve(null);

        reader.readAsDataURL(blob);
    });
}

/** How a picked element describes itself to somebody reading the report later. */
function describe(element) {
    if (!element) return null;

    const label =
        element.getAttribute('aria-label') ||
        element.getAttribute('title') ||
        (element.innerText || element.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 80);

    // The nearest ancestor that has one, because the click very often lands on a span inside a
    // button and the id worth reporting is the button's.
    const withId = element.closest('[id]');

    return {
        id: withId ? withId.id : '',
        label: label || element.tagName.toLowerCase(),
        tag: element.tagName.toLowerCase(),
    };
}

/**
 * Puts the page into "click the control you mean" mode, once.
 *
 * Resolves with the picked control, or null when the user pressed Escape. The listener is capturing
 * and cancels the event, so clicking a button to report it does not also press it.
 */
export function pickControl(dotNetRef) {
    if (picker) picker.cancel();

    const overlay = document.createElement('div');

    overlay.setAttribute('id', 'dmwsFeedbackPickerOutline');
    overlay.style.cssText =
        'position:fixed;pointer-events:none;z-index:2147483647;border:2px solid #ff6d00;' +
        'background:rgba(255,109,0,0.12);border-radius:3px;transition:all 60ms ease-out;';

    const hint = document.createElement('div');

    hint.setAttribute('id', 'dmwsFeedbackPickerHint');
    hint.textContent = 'Click the control you want to report — Escape to cancel';
    hint.style.cssText =
        'position:fixed;left:50%;top:12px;transform:translateX(-50%);z-index:2147483647;' +
        'background:#ff6d00;color:#fff;font:13px system-ui,sans-serif;padding:6px 12px;' +
        'border-radius:14px;pointer-events:none;box-shadow:0 2px 8px rgba(0,0,0,0.3);';

    document.body.appendChild(overlay);
    document.body.appendChild(hint);

    let done = false;

    const finish = result => {
        if (done) return;

        done = true;

        document.removeEventListener('mousemove', onMove, true);
        document.removeEventListener('click', onClick, true);
        document.removeEventListener('keydown', onKey, true);

        overlay.remove();
        hint.remove();

        picker = null;

        // Told rather than returned, because the pick happens long after the call: a promise across
        // the interop boundary would hold a circuit open for as long as the user took to decide.
        try { dotNetRef.invokeMethodAsync('OnControlPicked', result); } catch { /* gone */ }
    };

    const onMove = event => {
        const target = event.target;

        if (!target || target === overlay || target === hint) return;

        const box = target.getBoundingClientRect();

        overlay.style.left = box.left + 'px';
        overlay.style.top = box.top + 'px';
        overlay.style.width = box.width + 'px';
        overlay.style.height = box.height + 'px';
    };

    const onClick = event => {
        event.preventDefault();
        event.stopPropagation();

        finish(describe(event.target));
    };

    const onKey = event => {
        if (event.key !== 'Escape') return;

        event.preventDefault();
        event.stopPropagation();

        finish(null);
    };

    document.addEventListener('mousemove', onMove, true);
    document.addEventListener('click', onClick, true);
    document.addEventListener('keydown', onKey, true);

    picker = { cancel: () => finish(null) };

    return true;
}

/** Cancels a pick that is in progress. */
export function cancelPick() {
    if (picker) picker.cancel();

    return true;
}

/**
 * Hands a mailto: URL to whatever this machine uses for mail.
 *
 * This is the fallback for a machine with no mail server configured, and it is genuinely a
 * fallback rather than an equal alternative: NO mail client accepts an attachment through a
 * mailto: URL, because a link on a web page could otherwise attach a file off your disk. The
 * caller therefore names the files and their folder in the body, and records the send as
 * "handed to your mail program" rather than as "sent".
 *
 * location.href rather than window.open: a popup blocker stops the second one, and a mailto:
 * navigation does not replace the page - the browser hands it to the protocol handler and stays
 * where it is.
 */
export function openMailClient(uri) {
    if (!uri) return false;

    try {
        window.location.href = uri;

        return true;
    } catch {
        return false;
    }
}
