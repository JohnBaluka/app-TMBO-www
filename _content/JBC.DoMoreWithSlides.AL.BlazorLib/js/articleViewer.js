// Article Viewer JavaScript Interop for Blazor
let dotNetRef = null;
let youtubePlayer = null;
let youtubePlayerReady = false;
let presentationReveal = null;
let videoTimeTracker = null;

// The page's own film (the Video view) — the <video> its listeners were bound to, so a re-render that
// replaced the element binds the new one and one that did not is not bound twice.
let filmPlayerBound = null;

// Reveal.js module reference (loaded dynamically)
let Reveal = null;
let RevealZoom = null;
let RevealNotes = null;
let RevealSearch = null;
let RevealHighlight = null;
let revealLoaded = false;
let revealLoadPromise = null;

// Store all article view reveal instances
const articleRevealInstances = [];

/**
 * Reports an article slide's fragment position to .NET and writes it into the slide's counter directly.
 * The article view redraws only when one of its own parameters changes, which a fragment never does, so
 * the counter the .NET side keeps would otherwise stay at "0 / n" for as long as the slide played — the
 * reason the timestamps beside it are written straight into the page too. The text node Blazor drew is
 * the one changed, so its next render of the counter finds its own node.
 */
function reportArticleFragment(slideIndex, fragmentIndex, total) {
    const badge = document.getElementById(`fragmentNumber-${slideIndex}`);
    if (badge) {
        const text = `${fragmentIndex >= 0 ? fragmentIndex + 1 : 0} / ${total}`;
        const walker = document.createTreeWalker(badge, NodeFilter.SHOW_TEXT);
        let node = walker.nextNode();
        while (node && !node.nodeValue.trim()) node = walker.nextNode();
        if (node) node.nodeValue = text;
    }

    if (dotNetRef) {
        dotNetRef.invokeMethodAsync('UpdateArticleSlideFragmentState', slideIndex, fragmentIndex, total);
    }
}

/**
 * Which spoken line a matched span is — counted by line, not by span. A slide's notes are split into
 * lines at each animation break, so line n is the slide's n-th step; but a line written over two
 * paragraphs is two spans with one `data-note-id`, and counting spans put every step after it one late.
 */
function spokenLineIndex(spans, spanIndex) {
    const seen = new Set();
    let index = -1;

    for (let i = 0; i <= spanIndex && i < spans.length; i++) {
        const id = spans[i].getAttribute('data-note-id') || spans[i].id || `#${i}`;
        if (!seen.has(id)) {
            seen.add(id);
            index++;
        }
    }

    return index;
}

/**
 * Shows the article slide's step for its n-th spoken line — the first line is before any step — and
 * only when that step changes. The narration's time update runs several times a second, and every
 * slide() call re-synced the whole slide; a step the reader moved by hand is also left alone until the
 * next line.
 */
function applyArticleFragment(slideData, slideIndex, lineIndex) {
    const fragmentIndex = lineIndex <= 0 ? -1 : lineIndex - 1;

    if (slideData.appliedFragment === fragmentIndex) return;

    slideData.appliedFragment = fragmentIndex;
    slideData.instance.slide(0, 0, fragmentIndex);

    // Reported here as well: fragmentshown/fragmenthidden do not always fire when slide() targets a step.
    if (slideData.lastReportedFragment !== fragmentIndex) {
        slideData.lastReportedFragment = fragmentIndex;
        const section = slideData.element.querySelector('section');
        reportArticleFragment(slideIndex, fragmentIndex, countUniqueFragments(section));
    }
}

/**
 * The presentation view's own form of applyArticleFragment: the step for the n-th spoken line, only
 * when it is not the one showing — every slide() re-syncs and re-lays out the deck.
 */
function applyPresentationFragment(lineIndex) {
    if (!presentationReveal) return;

    const state = presentationReveal.getState();
    const target = lineIndex <= 0 ? -1 : lineIndex - 1;
    const current = typeof state.indexf === 'number' ? state.indexf : -1;

    if (current !== target) {
        presentationReveal.slide(state.indexh, state.indexv, target);
    }
}

// Count unique fragment indexes in a slide element (matches static site logic)
// Uses data-fragment-index attributes to count unique fragment steps,
// looking at both .fragment and .fragment-off elements
function countUniqueFragments(slideElement) {
    if (!slideElement) return 0;
    const fragmentElements = slideElement.querySelectorAll('[data-fragment-index]');
    const uniqueIndexes = new Set();
    fragmentElements.forEach(el => {
        const index = el.getAttribute('data-fragment-index');
        if (index !== null) {
            uniqueIndexes.add(index);
        }
    });
    return uniqueIndexes.size;
}

// Throttle for timestamp updates to Blazor (avoid excessive re-renders)
let lastTimestampUpdate = 0;
const TIMESTAMP_UPDATE_INTERVAL = 250; // ms

// Cached total duration for progress bar updates 
let cachedTotalDuration = 0;
let presentationResizeBound = false;
// The article view's equivalent: bound once, and re-applies the overrides Reveal's own resize handling
// rewrites on every embedded instance. Without it a resized window left every article slide drawn at
// Reveal's own scale.
let articleResizeBound = false;
let articleRelayoutTimer = null;

function scheduleArticleRelayout() {
    if (articleRevealInstances.length === 0) {
        return;
    }

    // A drag of a window edge raises resize dozens of times a second; one relayout at the end is enough.
    clearTimeout(articleRelayoutTimer);
    articleRelayoutTimer = setTimeout(() => {
        articleRelayoutTimer = null;
        layoutAllArticleSlides();
    }, 150);
}
let lastHighlightedNoteId = null;
let lastHighlightedSlideIndex = null;
let presentationResizeObserver = null;

// User seek override: suppresses timeupdate-driven note updates after user click
let userSeekNoteId = null;
let userSeekSlideIndex = null;
let userSeekTimestamp = 0;
const USER_SEEK_TIMEOUT_MS = 2000;

// Transition suppression: blocks ALL note updates during view switch / mode toggle
let noteUpdateSuppressedUntil = 0;
const NOTE_SUPPRESS_TRANSITION_MS = 1500;
let pendingPresentationSeek = null;

// Logging configuration
// 0 = None, 1 = Error, 2 = Warning, 3 = Info, 4 = Debug, 5 = Trace
let logLevel = 3; // Default to Info

// Background retry timers for deferred initialization when DOM elements
// are not yet available (Blazor WASM may flush renders after JS interop).
let _articleRetryTimer = null;
let _videoPlayerRetryTimer = null;

// Logging helper functions
function logError(message, ...args) {
    if (logLevel >= 1) console.error(`[ArticleViewer] ${message}`, ...args);
}

function logWarning(message, ...args) {
    if (logLevel >= 2) console.warn(`[ArticleViewer] ${message}`, ...args);
}

function logInfo(message, ...args) {
    if (logLevel >= 3) console.info(`[ArticleViewer] ${message}`, ...args);
}

function logDebug(message, ...args) {
    if (logLevel >= 4) console.debug(`[ArticleViewer] ${message}`, ...args);
}

function logTrace(message, ...args) {
    if (logLevel >= 5) console.log(`[ArticleViewer][TRACE] ${message}`, ...args);
}

export function setLogLevel(level) {
    if (level == 4) { // 4 = Debug
        //debugger;
    }
    logLevel = level;
    logInfo(`Log level set to: ${level}`);
}

function getPresentationAspectRatio() {
    const container = document.getElementById('presentationContainer');
    if (container) {
        const width = parseFloat(container.dataset.aspectWidth);
        const height = parseFloat(container.dataset.aspectHeight);
        if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
            return width / height;
        }
    }
    return 1280 / 720;
}

function resizePresentationLayout() {
    const wrapper = document.getElementById('presentationViewWrapper');
    const container = document.getElementById('presentationContainer');
    const revealEl = document.getElementById('presentationReveal');

    if (!wrapper || !container || !revealEl) {
        return;
    }

    // Guard: skip layout when the wrapper is hidden (display:none gives zero rect)
    const wrapperRect = wrapper.getBoundingClientRect();
    if (wrapperRect.width === 0 && wrapperRect.height === 0) {
        logTrace('resizePresentationLayout: wrapper hidden, skipping');
        return;
    }

    const toolbar = document.getElementById('presentationToolbar');
    const styles = window.getComputedStyle(wrapper);
    const paddingX = parseFloat(styles.paddingLeft) + parseFloat(styles.paddingRight);
    const paddingY = parseFloat(styles.paddingTop) + parseFloat(styles.paddingBottom);
    const availableWidth = Math.max(0, wrapperRect.width - paddingX);
    const availableHeight = Math.max(0, wrapperRect.height - paddingY);
    const toolbarHeight = toolbar ? toolbar.offsetHeight : 0;
    const aspectRatio = getPresentationAspectRatio();

    const maxRevealHeight = Math.max(0, availableHeight - toolbarHeight);
    const revealWidth = Math.min(availableWidth, maxRevealHeight * aspectRatio);
    const revealHeight = revealWidth / aspectRatio;
    const containerHeight = revealHeight + toolbarHeight;

    container.style.width = `${revealWidth}px`;
    container.style.height = `${containerHeight}px`;
    revealEl.style.width = `${revealWidth}px`;
    revealEl.style.height = `${revealHeight}px`;

    if (presentationReveal) {
        // Don't reconfigure width/height - keep 1280x720 as internal slide content dimensions.
        // Reveal.js scales the content to fit the container element size automatically.
        presentationReveal.layout();
    }
}

export function resizePresentationLayoutInterop() {
    resizePresentationLayout();
}

/**
 * Notifies JS that the user clicked a note, suppressing timeupdate-driven
 * note updates until the media catches up to the clicked note's position.
 */
export function notifyUserNoteSeek(noteId, slideIndex) {
    userSeekNoteId = noteId;
    userSeekSlideIndex = slideIndex;
    userSeekTimestamp = Date.now();
    lastHighlightedNoteId = noteId;
    lastHighlightedSlideIndex = slideIndex;
    logDebug(`notifyUserNoteSeek: suppressing updates until note '${noteId}' slide ${slideIndex}`);
}

function invokeCurrentNoteUpdate(noteId, slideIndex, timestamp) {
    if (!dotNetRef || !noteId || !slideIndex) {
        return;
    }

    // Suppress ALL note updates during view switch / mode toggle transitions
    if (Date.now() < noteUpdateSuppressedUntil) {
        return;
    }

    // Suppress timeupdate-driven updates during user-initiated seek
    if (userSeekNoteId !== null) {
        if (noteId === userSeekNoteId && slideIndex === userSeekSlideIndex) {
            // Media caught up to user's click target — clear override
            userSeekNoteId = null;
            userSeekSlideIndex = null;
        } else if (Date.now() - userSeekTimestamp < USER_SEEK_TIMEOUT_MS) {
            // Override active and note doesn't match — suppress
            return;
        } else {
            // Timeout expired — clear override
            userSeekNoteId = null;
            userSeekSlideIndex = null;
        }
    }

    if (noteId === lastHighlightedNoteId && slideIndex === lastHighlightedSlideIndex) {
        return;
    }

    lastHighlightedNoteId = noteId;
    lastHighlightedSlideIndex = slideIndex;
    highlightNoteLines(noteId);
    dotNetRef.invokeMethodAsync('UpdateCurrentNote', noteId, slideIndex, timestamp);
}

// Load Reveal.js and its plugins dynamically
async function loadRevealJs() {
    if (revealLoaded) {
        logDebug('Reveal.js already loaded');
        return true;
    }
    
    if (revealLoadPromise) {
        logDebug('Reveal.js load already in progress, waiting...');
        return revealLoadPromise;
    }
    
    logInfo('Loading Reveal.js and plugins...');
    
    revealLoadPromise = (async () => {
        const revealBasePaths = [
            'libs/revealjs/5.0.5',
            'revealjs',
            '_content/JBC.DoMoreWithSlides.AL.BlazorLib/revealjs'
        ];

        for (const basePath of revealBasePaths) {
            try {
                const trimmedBasePath = basePath.replace(/^\/+|\/+$/g, '');
                const moduleRoot = `/${trimmedBasePath}`;

                logDebug(`Loading Reveal.js from: ${moduleRoot}/reveal.esm.js`);

                const [revealModule, zoomModule, notesModule, searchModule, highlightModule] = await Promise.all([
                    import(`${moduleRoot}/reveal.esm.js`),
                    import(`${moduleRoot}/plugin/zoom/zoom.esm.js`),
                    import(`${moduleRoot}/plugin/notes/notes.esm.js`),
                    import(`${moduleRoot}/plugin/search/search.esm.js`),
                    import(`${moduleRoot}/plugin/highlight/highlight.esm.js`)
                ]);

                Reveal = revealModule.default;
                RevealZoom = zoomModule.default;
                RevealNotes = notesModule.default;
                RevealSearch = searchModule.default;
                RevealHighlight = highlightModule.default;

                revealLoaded = true;
                logInfo(`Reveal.js and plugins loaded successfully from: ${moduleRoot}`);
                logDebug(`Reveal constructor: ${typeof Reveal}`);
                logDebug(`RevealZoom: ${typeof RevealZoom}`);
                logDebug(`RevealNotes: ${typeof RevealNotes}`);
                logDebug(`RevealSearch: ${typeof RevealSearch}`);
                logDebug(`RevealHighlight: ${typeof RevealHighlight}`);

                return true;
            } catch (error) {
                logDebug(`Reveal.js load attempt failed for base path: ${basePath}`, error);
            }
        }

        logError('Failed to load Reveal.js from all known base paths');
        revealLoadPromise = null;
        return false;
    })();
    
    return revealLoadPromise;
}

// Check if Reveal.js is loaded
export function isRevealLoaded() {
    logDebug(`isRevealLoaded: ${revealLoaded}`);
    return revealLoaded;
}

export async function initialize(dotNet, level = 3) {
    logLevel = level;
    logInfo(`Initializing ArticleViewer JS module (log level: ${level})`);
    dotNetRef = dotNet;
    
    // Load Reveal.js
    const revealSuccess = await loadRevealJs();
    if (revealSuccess) {
        logInfo('Reveal.js is available for use');
    } else {
        logError('Reveal.js failed to load - presentation features will be disabled');
    }
    
    // Load YouTube API
    if (!document.getElementById('youtube-api-script')) {
        logDebug('Loading YouTube API script');
        const tag = document.createElement('script');
        tag.id = 'youtube-api-script';
        tag.src = 'https://www.youtube.com/iframe_api';
        document.head.appendChild(tag);
    }
    
    // Setup audio timeupdate listeners
    setupAudioListeners();

    // A slide's notes drawn from its notes file are markup Blazor does not own, so no @onclick can be
    // put on a line of them — listeners on the document find the line. A click marks it; hovering or
    // marking a line puts its Play and Copy buttons beside it (see showNoteLineActions).
    document.removeEventListener('click', onNotesHtmlClick);
    document.addEventListener('click', onNotesHtmlClick);
    document.removeEventListener('mouseover', onNoteLineHover);
    document.addEventListener('mouseover', onNoteLineHover);
    document.removeEventListener('scroll', onNoteLineScroll, true);
    document.addEventListener('scroll', onNoteLineScroll, true);
    document.removeEventListener('play', onNarrationPlayOrPause, true);
    document.addEventListener('play', onNarrationPlayOrPause, true);
    document.removeEventListener('pause', onNarrationPlayOrPause, true);
    document.addEventListener('pause', onNarrationPlayOrPause, true);

    logInfo('ArticleViewer JS module initialized');
}

// -- A line of a slide's notes: marked by a click, played and copied by its own two buttons ------------
//
// Clicking a line used to seek and play at once. Now a click only marks it — highlighted, nothing plays
// (the user's ask, 2026-09-30) — and a small Play and Copy button sit beside the line under the pointer,
// or beside the marked line when the pointer is elsewhere. One pair of buttons for the whole document,
// fixed to the window rather than put inside the notes: the notes are Blazor's and their markup is
// replaced whenever the line being said moves, which would take a child of theirs with it.

const NOTE_LINE_SELECTOR = '.slide-notes .note-line[data-note-id]';

/** The line a click marked — { slideIndex, noteId } — or null. */
let markedNoteLine = null;

/** The line the buttons are beside now — { slideIndex, noteId } — or null. */
let shownNoteLine = null;

let noteLineActions = null;

/**
 * The slide whose animation step the pointer is holding — a line of it is hovered — or null. While it is
 * set the narration leaves that slide's step alone, so hovering works while the audio plays too (the
 * user's ask, 2026-09-30); moving off every line hands the step back to the narration.
 */
let hoverFragmentSlide = null;

function cssEscape(value) {
    return window.CSS && CSS.escape ? CSS.escape(value) : String(value).replace(/["\\]/g, '');
}

function noteIdOf(line) {
    return line.getAttribute('data-note-id') || line.id;
}

/** Every piece of a line — a line written over two paragraphs is two spans — in its own slide's notes. */
function noteLinePieces(slideIndex, noteId) {
    const notes = document.getElementById(`slideNotes-${slideIndex}`);

    return notes
        ? Array.from(notes.querySelectorAll(`.note-line[data-note-id="${cssEscape(noteId)}"]`))
        : [];
}

function ensureNoteLineActions() {
    if (noteLineActions && noteLineActions.isConnected) return noteLineActions;

    noteLineActions = document.createElement('div');
    noteLineActions.id = 'noteLineActions';
    noteLineActions.className = 'note-line-actions';
    noteLineActions.innerHTML =
        '<button type="button" id="noteLinePlayButton" data-action="play" title="Play from this line" aria-label="Play from this line">play_arrow</button>' +
        '<button type="button" id="noteLineCopyButton" data-action="copy" title="Copy this line" aria-label="Copy this line">content_copy</button>';

    // Leaving the buttons for anywhere but the line they belong to goes back to the marked line — after the
    // same grace a line gets, so a pointer that slips off the edge and back does not lose them.
    noteLineActions.addEventListener('mouseleave', event => {
        const into = event.relatedTarget && event.relatedTarget.closest
            ? event.relatedTarget.closest(NOTE_LINE_SELECTOR)
            : null;

        if (!into) scheduleLeaveNoteLines();
    });

    noteLineActions.addEventListener('mouseenter', cancelLeaveNoteLines);

    document.body.appendChild(noteLineActions);

    return noteLineActions;
}

/**
 * Puts the two buttons above the start of a line (the user's ask, 2026-09-30), or hides them for null or
 * a line no longer drawn. Below the line instead when above would leave the window.
 */
function showNoteLineActions(target) {
    const pieces = target ? noteLinePieces(target.slideIndex, target.noteId) : [];
    const first = pieces.length > 0 ? pieces[0] : null;
    const rects = first ? first.getClientRects() : [];

    if (!first || rects.length === 0 || first.offsetParent === null) {
        shownNoteLine = null;
        if (noteLineActions) noteLineActions.classList.remove('visible');
        return;
    }

    const actions = ensureNoteLineActions();
    const rect = rects[0];

    shownNoteLine = target;
    actions.classList.add('visible');
    refreshNoteLinePlayButton();

    const width = actions.offsetWidth || 56;
    const height = actions.offsetHeight || 24;
    let left = rect.left;
    let top = rect.top - height - 2;

    if (top < 4) top = rect.bottom + 2;

    left = Math.min(Math.max(4, left), window.innerWidth - width - 4);

    actions.style.left = `${Math.round(left)}px`;
    actions.style.top = `${Math.round(top)}px`;
}

/** The narration element that is playing, or null. */
function playingNarration() {
    return ['consolidated-audio-player', 'consolidated-video-player']
        .map(id => document.getElementById(id))
        .find(media => media && !media.paused) || null;
}

/** True while a narration is playing. */
function isNarrationPlaying() {
    return playingNarration() !== null;
}

/** The Play button shows Pause while a narration plays — pressing it then pauses (the user's ask). */
function refreshNoteLinePlayButton() {
    const button = noteLineActions ? noteLineActions.querySelector('#noteLinePlayButton') : null;

    if (!button) return;

    const playing = isNarrationPlaying();
    const label = playing ? 'Pause' : 'Play from this line';

    button.textContent = playing ? 'pause' : 'play_arrow';
    button.title = label;
    button.setAttribute('aria-label', label);
}

/** A media element started or stopped somewhere — media events do not bubble, so this listens capturing. */
function onNarrationPlayOrPause(event) {
    if (event.target && (event.target.tagName === 'AUDIO' || event.target.tagName === 'VIDEO'))
        refreshNoteLinePlayButton();
}

/** The pointer is over a line: its buttons, and its animation step held until the pointer leaves the lines. */
function hoverNoteLine(target) {
    showNoteLineActions(target);

    if (showFragmentForNoteLine(target.slideIndex, target.noteId)) hoverFragmentSlide = target.slideIndex;
}

/**
 * How long the buttons stay after the pointer leaves a line, in milliseconds. The buttons sit just above
 * the line, and on the way up to them the pointer crosses the gap between — which is off every line —
 * so without a grace they moved away the moment the user reached for them (the user's report, 2026-10-02).
 */
const NOTE_LINE_LEAVE_GRACE_MS = 400;

let noteLineLeaveTimer = null;

function scheduleLeaveNoteLines() {
    if (noteLineLeaveTimer) return;

    noteLineLeaveTimer = setTimeout(() => {
        noteLineLeaveTimer = null;
        leaveNoteLines();
    }, NOTE_LINE_LEAVE_GRACE_MS);
}

function cancelLeaveNoteLines() {
    if (!noteLineLeaveTimer) return;

    clearTimeout(noteLineLeaveTimer);
    noteLineLeaveTimer = null;
}

/** True when a line's buttons still have a line on screen to sit beside. */
function isNoteLineShowing(target) {
    const pieces = target ? noteLinePieces(target.slideIndex, target.noteId) : [];

    return pieces.length > 0 && pieces[0].offsetParent !== null && pieces[0].getClientRects().length > 0;
}

/**
 * Takes the buttons away and forgets the line they were beside — called by .NET when the Main View changes
 * view, because the buttons are fixed to the window and outlive the article they were drawn for (the user's
 * report, 2026-10-02: they stayed over the Slides view). The marked line is forgotten too, so they do not
 * come back beside it until the article is shown and a line is hovered or clicked again.
 */
export function hideNoteLineActions() {
    cancelLeaveNoteLines();
    hoverFragmentSlide = null;
    markedNoteLine = null;
    shownNoteLine = null;

    document.querySelectorAll('.slide-notes .note-line.selected')
        .forEach(element => element.classList.remove('selected'));

    if (noteLineActions) noteLineActions.classList.remove('visible');
}

/**
 * The pointer left every line: the buttons go back to the marked line, or away, and the narration has
 * its step back. With nothing playing, the marked line's step is shown again.
 */
function leaveNoteLines() {
    cancelLeaveNoteLines();
    hoverFragmentSlide = null;

    showNoteLineActions(markedNoteLine);

    if (markedNoteLine && !isNarrationPlaying())
        showFragmentForNoteLine(markedNoteLine.slideIndex, markedNoteLine.noteId);
}

/**
 * Shows the slide's animation step for a line of its notes — the step the narration shows when it says
 * that line (the user's ask, 2026-09-30: "on hover or on click, also trigger the fragment/animation").
 * The same mapping the narration uses: the n-th spoken line is step n-1, the first is before any step.
 * A slide that has never played still has its fragments switched off (drawn all at once), so they are
 * switched on first, as pressing Play does. Works while a narration plays as well: the hover holds the
 * slide's step (hoverFragmentSlide) so the narration's time updates do not take it straight back.
 * Answers true when there was a step to show.
 */
function showFragmentForNoteLine(slideIndex, noteId) {
    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
    const slideCard = document.getElementById(`slide-${slideIndex}`);

    if (!slideData || !slideData.instance || !slideCard) return false;

    const lines = slideCard.querySelectorAll('[data-start][data-end]');
    let position = -1;

    for (let i = 0; i < lines.length; i++) {
        if ((lines[i].getAttribute('data-note-id') || lines[i].id) === noteId) {
            position = i;
            break;
        }
    }

    if (position < 0) return false;

    const section = slideData.element.querySelector('section');
    const fragmentsOff = section ? section.querySelectorAll('.fragment-off') : [];

    if (fragmentsOff.length > 0) {
        fragmentsOff.forEach(f => {
            f.classList.remove('fragment-off');
            f.classList.remove('visible');
            f.classList.add('fragment');
        });

        // The step it was on meant "everything drawn"; forget it so the step below is applied.
        slideData.appliedFragment = undefined;

        if (dotNetRef) dotNetRef.invokeMethodAsync('ShowArticleSlideFragmentControls', slideIndex);
    }

    applyArticleFragment(slideData, slideIndex, spokenLineIndex(lines, position));

    return true;
}

function markNoteLine(slideIndex, noteId) {
    document.querySelectorAll('.slide-notes .note-line.selected')
        .forEach(element => element.classList.remove('selected'));

    markedNoteLine = { slideIndex, noteId };

    noteLinePieces(slideIndex, noteId).forEach(element => element.classList.add('selected'));
    hoverNoteLine(markedNoteLine);
}

/**
 * A click in a slide's notes: on a line's Play or Copy button, that; on a line, marks it. Nothing plays
 * until Play is pressed. A comment's icon inside a line is left to show its comment.
 */
function onNotesHtmlClick(event) {
    const target = event.target;

    if (!target || !target.closest) return;

    const button = target.closest('#noteLineActions button');

    if (button) {
        event.preventDefault();
        onNoteLineAction(button);
        return;
    }

    const line = target.closest(NOTE_LINE_SELECTOR);

    if (!line || target.closest('.dmws-comment')) return;

    const noteId = noteIdOf(line);
    const slideIndex = parseInt(line.getAttribute('data-slide-index'));

    if (!noteId || !slideIndex) return;

    markNoteLine(slideIndex, noteId);
}

function onNoteLineHover(event) {
    const target = event.target;

    if (!target || !target.closest) return;

    if (target.closest('#noteLineActions')) {
        cancelLeaveNoteLines();
        return;
    }

    // Buttons left over a line that is no longer drawn — its view was swapped out — go at once.
    if (shownNoteLine && !isNoteLineShowing(shownNoteLine)) {
        if (markedNoteLine && !isNoteLineShowing(markedNoteLine)) markedNoteLine = null;
        leaveNoteLines();
    }

    const line = target.closest(NOTE_LINE_SELECTOR);

    if (line) {
        cancelLeaveNoteLines();

        const noteId = noteIdOf(line);
        const slideIndex = parseInt(line.getAttribute('data-slide-index'));

        if (noteId && slideIndex
            && !(shownNoteLine && shownNoteLine.noteId === noteId && shownNoteLine.slideIndex === slideIndex
                 && hoverFragmentSlide === slideIndex)) {
            hoverNoteLine({ slideIndex, noteId });
        }
        return;
    }

    // Off every line — once, not on every move over the rest of the page, and after a grace (see
    // NOTE_LINE_LEAVE_GRACE_MS) so the pointer can cross to the buttons.
    if (hoverFragmentSlide !== null || (shownNoteLine && (!markedNoteLine
        || shownNoteLine.noteId !== markedNoteLine.noteId
        || shownNoteLine.slideIndex !== markedNoteLine.slideIndex))) {
        scheduleLeaveNoteLines();
    }
}

/** The buttons are fixed to the window, so they follow their line when anything scrolls. */
function onNoteLineScroll() {
    if (shownNoteLine) showNoteLineActions(shownNoteLine);
}

function onNoteLineAction(button) {
    const target = shownNoteLine;

    if (!target) return;

    if (button.getAttribute('data-action') === 'play') {
        // A toggle: Pause while a narration plays, Play from this line otherwise. The media's own pause
        // event tells .NET, as the toolbar's pause does.
        const playing = playingNarration();

        if (playing) playing.pause();
        else if (dotNetRef) dotNetRef.invokeMethodAsync('OnNoteLinePlay', target.slideIndex, target.noteId);

        return;
    }

    copyNoteLine(target, button);
}

/** Copies a line's words — every piece of it, without the author's comments — to the clipboard. */
function copyNoteLine(target, button) {
    const text = noteLinePieces(target.slideIndex, target.noteId)
        .map(piece => {
            const copy = piece.cloneNode(true);
            copy.querySelectorAll('.dmws-comment').forEach(comment => comment.remove());
            return copy.textContent;
        })
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();

    if (!text) return;

    const done = () => {
        button.classList.add('copied');
        button.textContent = 'check';
        setTimeout(() => {
            button.classList.remove('copied');
            button.textContent = 'content_copy';
        }, 1200);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(done, () => copyWithSelection(text) && done());
    } else if (copyWithSelection(text)) {
        done();
    }
}

/** The clipboard for a page the Clipboard API is not allowed on — an http:// preview, an old WebView. */
function copyWithSelection(text) {
    const area = document.createElement('textarea');
    area.value = text;
    area.setAttribute('readonly', '');
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.select();

    let copied = false;
    try { copied = document.execCommand('copy'); } catch (e) { copied = false; }

    area.remove();
    return copied;
}

/**
 * Lights every stretch of the line being said in notes drawn as markup — every span carrying its
 * `data-note-id`, because a line written over two paragraphs is two spans. Blazor redraws the class
 * on its next render; this is what moves it between renders, several times a second.
 */
function highlightNoteLines(noteId) {
    document.querySelectorAll('.slide-notes-html .note-line.current')
        .forEach(element => element.classList.remove('current'));

    if (!noteId) return;

    const selector = `.slide-notes-html .note-line[data-note-id="${window.CSS && CSS.escape ? CSS.escape(noteId) : noteId}"]`;

    document.querySelectorAll(selector).forEach(element => element.classList.add('current'));
}

export function dispose() {
    logInfo('Disposing ArticleViewer JS module');
    stopAllMedia();
    if (videoTimeTracker) {
        clearInterval(videoTimeTracker);
        videoTimeTracker = null;
    }
    if (_articleRetryTimer) {
        clearInterval(_articleRetryTimer);
        _articleRetryTimer = null;
    }
    if (_videoPlayerRetryTimer) {
        clearInterval(_videoPlayerRetryTimer);
        _videoPlayerRetryTimer = null;
    }
    
    // Destroy all article reveal instances
    logDebug(`Disposing ${articleRevealInstances.length} article reveal instances`);
    for (const slideData of articleRevealInstances) {
        try {
            if (slideData.instance && typeof slideData.instance.destroy === 'function') {
                slideData.instance.destroy();
            }
        } catch (e) {
            logWarning(`Error destroying article reveal instance ${slideData.index}`, e);
        }
    }
    articleRevealInstances.length = 0;
    
    // Destroy presentation reveal instance
    if (presentationReveal && typeof presentationReveal.destroy === 'function') {
        try {
            presentationReveal.destroy();
        } catch (e) {
            logWarning('Error destroying presentation reveal instance', e);
        }
    }

    if (presentationResizeObserver) {
        presentationResizeObserver.disconnect();
        presentationResizeObserver = null;
    }
    
    document.removeEventListener('click', onNotesHtmlClick);
    document.removeEventListener('mouseover', onNoteLineHover);
    document.removeEventListener('scroll', onNoteLineScroll, true);
    document.removeEventListener('play', onNarrationPlayOrPause, true);
    document.removeEventListener('pause', onNarrationPlayOrPause, true);

    cancelLeaveNoteLines();
    if (noteLineActions) noteLineActions.remove();
    hoverFragmentSlide = null;
    noteLineActions = null;
    markedNoteLine = null;
    shownNoteLine = null;

    dotNetRef = null;
    youtubePlayer = null;
    presentationReveal = null;
    filmPlayerBound = null;
    // Taken off, not just forgotten: the next Main View may listen to the same element (see unbindMediaListeners).
    unbindMediaListeners(consolidatedAudioListenersAttached);
    unbindMediaListeners(consolidatedVideoListenersAttached);
    consolidatedAudioListenersAttached = null;
    consolidatedVideoListenersAttached = null;
    logDebug('ArticleViewer JS module disposed');
}

function setupAudioListeners() {
    logDebug('Setting up audio listeners');
    // Re-run periodically to catch dynamically added audio elements
    const audioElements = document.querySelectorAll('audio');
    logTrace(`Found ${audioElements.length} audio elements`);
    
    audioElements.forEach(audio => {
        if (audio.dataset.listenerAttached) {
            return;
        }

        audio.dataset.listenerAttached = 'true';
        logTrace(`Attaching listeners to audio element: ${audio.id}`);

        if (audio.id === 'pres-audio-player') {
            attachPresentationAudioListeners(audio);
            return;
        }

        audio.addEventListener('timeupdate', () => {
            if (dotNetRef && !audio.paused) {
                const slideMatch = audio.id?.match(/audio-player-(\d+)/);
                if (slideMatch) {
                    const slideIndex = parseInt(slideMatch[1]);
                    const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
                    if (notesEl) {
                        const lines = notesEl.querySelectorAll('[data-start]');
                        const now = audio.currentTime;
                        
                        for (const line of lines) {
                            const start = parseFloat(line.getAttribute('data-start'));
                            const end = parseFloat(line.getAttribute('data-end'));
                            const startVideo = parseFloat(line.getAttribute('data-start-video'));
                            
                            if (now >= start && now <= end) {
                                const noteId = line.getAttribute('data-note-id') || line.id;
                                const absoluteTime = startVideo + (now - start);
                                invokeCurrentNoteUpdate(noteId, slideIndex, absoluteTime);
                                break;
                            }
                        }
                    }
                    
                    // Update timestamp in bottom toolbar via Blazor
                    // Calculate absolute video time from audio currentTime
                    const notesContainer = document.getElementById(`slideNotes-${slideIndex}`);
                    if (notesContainer) {
                        const firstLine = notesContainer.querySelector('[data-start-video]');
                        if (firstLine) {
                            const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                            const absoluteTime = slideStartVideo + audio.currentTime;
                            
                            // Update progress bar directly in JS (no Blazor round-trip)
                            updateProgressBarDirect(absoluteTime);
                            
                            // Throttle Blazor timestamp updates
                            const now = Date.now();
                            if (now - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
                                lastTimestampUpdate = now;
                                dotNetRef.invokeMethodAsync('UpdateTimestamp', absoluteTime);
                            }
                        }
                    }
                    
                    // Navigate to matching fragment based on current time
                    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
                    if (slideData && slideData.instance) {
                        const notesForFragment = document.getElementById(`slideNotes-${slideIndex}`);
                        if (notesForFragment) {
                            const lines = notesForFragment.querySelectorAll('[data-start]');
                            const now = audio.currentTime;
                            let lineIndex = -1;
                            
                            for (let i = 0; i < lines.length; i++) {
                                const start = parseFloat(lines[i].getAttribute('data-start'));
                                const end = parseFloat(lines[i].getAttribute('data-end'));
                                if (now >= start && now <= end) {
                                    lineIndex = i;
                                    break;
                                }
                            }
                            
                            if (lineIndex >= 0) {
                                applyArticleFragment(slideData, slideIndex, spokenLineIndex(lines, lineIndex));
                            }
                        }
                    }
                }
            }
        });
        
        audio.addEventListener('play', () => {
            logDebug(`Audio play: ${audio.id}`);
            // Pause all other audio
            document.querySelectorAll('audio').forEach(a => {
                if (a !== audio && !a.paused) a.pause();
            });
            
            // Get slide index from audio id
            const slideMatch = audio.id?.match(/audio-player-(\d+)/);
            if (slideMatch) {
                const slideIndex = parseInt(slideMatch[1]);
                
                // Enable fragments on the slide's section (convert fragment-off to fragment)
                const revealEl = document.getElementById(`reveal-${slideIndex}`);
                if (revealEl) {
                    const fragmentsOff = revealEl.querySelectorAll('.fragment-off');
                    logDebug(`Enabling ${fragmentsOff.length} fragments for slide ${slideIndex}`);
                    fragmentsOff.forEach(f => {
                        f.classList.remove('fragment-off');
                        f.classList.remove('visible');
                        f.classList.add('fragment');
                    });
                }
                
                // Expand the slideNavCard in the sidebar and show fragment controls
                if (dotNetRef) {
                    dotNetRef.invokeMethodAsync('ExpandSlideNavCard', slideIndex);
                    dotNetRef.invokeMethodAsync('ShowArticleSlideFragmentControls', slideIndex);
                }
            }
            
            if (dotNetRef) {
                dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
            }
        });
        
        audio.addEventListener('pause', () => {
            logDebug(`Audio pause: ${audio.id}`);
            if (dotNetRef) {
                dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
            }
        });
        
        audio.addEventListener('ended', () => {
            logDebug(`Audio ended: ${audio.id}`);
            if (dotNetRef) {
                dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
            }
            
            // Auto-advance to next slide
            const slideMatch = audio.id?.match(/audio-player-(\d+)/);
            if (slideMatch) {
                const slideIndex = parseInt(slideMatch[1]);
                const nextSlideIndex = slideIndex + 1;
                const nextAudio = document.getElementById(`audio-player-${nextSlideIndex}`);
                if (nextAudio) {
                    // Scroll to next slide and play
                    const nextSlideEl = document.getElementById(`slide-${nextSlideIndex}`);
                    if (nextSlideEl) {
                        nextSlideEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                    setTimeout(() => {
                        nextAudio.play();
                    }, 500);
                }
            }
        });
    });
}

function attachPresentationAudioListeners(audio) {
    audio.addEventListener('timeupdate', () => {
        if (!dotNetRef || audio.paused) {
            return;
        }

        if (!presentationReveal) {
            return;
        }

        const state = presentationReveal.getState();
        const slideIndex = state.indexh + 1;
        const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
        if (!notesEl) {
            return;
        }

        const now = audio.currentTime;
        const lines = notesEl.querySelectorAll('[data-start]');
        let matched = false;
        let matchedLineIndex = -1;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const start = parseFloat(line.getAttribute('data-start'));
            const end = parseFloat(line.getAttribute('data-end'));
            const startVideo = parseFloat(line.getAttribute('data-start-video')) || 0;

            if (now >= start && now <= end) {
                const noteId = line.getAttribute('data-note-id') || line.id;
                const absoluteTime = startVideo + (now - start);
                updateProgressBarDirect(absoluteTime);

                const timestampNow = Date.now();
                if (timestampNow - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
                    lastTimestampUpdate = timestampNow;
                    dotNetRef.invokeMethodAsync('UpdateTimestamp', absoluteTime);
                }

                invokeCurrentNoteUpdate(noteId, slideIndex, absoluteTime);
                matched = true;
                matchedLineIndex = i;
                break;
            }
        }

        // Navigate to the correct fragment based on the matched line index
        // (matches static site: line 0 -> fragment -1, line 1 -> fragment 0, etc.)
        if (matchedLineIndex >= 0) {
            applyPresentationFragment(spokenLineIndex(lines, matchedLineIndex));
        }

        if (!matched) {
            const firstLine = notesEl.querySelector('[data-start-video]');
            if (firstLine) {
                const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                const absoluteTime = slideStartVideo + now;
                updateProgressBarDirect(absoluteTime);

                const timestampNow = Date.now();
                if (timestampNow - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
                    lastTimestampUpdate = timestampNow;
                    dotNetRef.invokeMethodAsync('UpdateTimestamp', absoluteTime);
                }
            }
        }
    });

    audio.addEventListener('play', () => {
        logDebug('Presentation audio play');
        enablePresentationFragmentsForSlide();
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
            // Show fragment controls when audio starts playing
            dotNetRef.invokeMethodAsync('ShowPresentationFragmentControls');
        }
    });

    audio.addEventListener('pause', () => {
        logDebug('Presentation audio pause');
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
        }
    });

    audio.addEventListener('ended', () => {
        logDebug('Presentation audio ended');
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
        }
    });
}

function enablePresentationFragmentsForSlide() {
    if (!presentationReveal) {
        return;
    }

    const currentSlide = presentationReveal.getCurrentSlide();
    if (!currentSlide) {
        return;
    }

    const fragmentsOff = currentSlide.querySelectorAll('.fragment-off');
    fragmentsOff.forEach(fragment => {
        fragment.classList.remove('fragment-off');
        fragment.classList.add('fragment');
    });
}

// Enable fragments for all slides in presentation view
export function enablePresentationFragments() {
    if (!presentationReveal) return;
    const container = document.getElementById('presentationReveal');
    if (!container) return;
    const fragmentsOff = container.querySelectorAll('.fragment-off');
    logDebug(`enablePresentationFragments: enabling ${fragmentsOff.length} fragments`);
    fragmentsOff.forEach(f => {
        f.classList.remove('fragment-off');
        f.classList.remove('visible');
        f.classList.add('fragment');
    });
    // Reset to beginning of fragments for current slide
    const state = presentationReveal.getState();
    presentationReveal.slide(state.indexh, state.indexv, -1);
    // Update Blazor with new fragment count
    if (dotNetRef) {
        const currentSlide = presentationReveal.getCurrentSlide();
        const total = countUniqueFragments(currentSlide);
        dotNetRef.invokeMethodAsync('UpdatePresentationState', state.indexh, -1, total);
    }
}

// Disable fragments for all slides in presentation view
export function disablePresentationFragments() {
    if (!presentationReveal) return;
    const container = document.getElementById('presentationReveal');
    if (!container) return;
    const fragments = container.querySelectorAll('.fragment');
    logDebug(`disablePresentationFragments: disabling ${fragments.length} fragments`);
    fragments.forEach(f => {
        f.classList.remove('fragment');
        f.classList.add('fragment-off');
        f.classList.add('visible');
    });
    // Update Blazor
    if (dotNetRef) {
        const state = presentationReveal.getState();
        const currentSlide = presentationReveal.getCurrentSlide();
        const total = countUniqueFragments(currentSlide);
        dotNetRef.invokeMethodAsync('UpdatePresentationState', state.indexh, -1, total);
    }
}

// Enable fragments for a specific article slide
export function enableArticleSlideFragments(slideIndex) {
    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
    if (!slideData) return;
    const section = slideData.element.querySelector('section');
    if (!section) return;
    const fragmentsOff = section.querySelectorAll('.fragment-off');
    logDebug(`enableArticleSlideFragments: slide ${slideIndex}, enabling ${fragmentsOff.length} fragments`);
    fragmentsOff.forEach(f => {
        f.classList.remove('fragment-off');
        f.classList.remove('visible');
        f.classList.add('fragment');
    });
    // Reset to beginning
    slideData.instance.slide(0, 0, -1);
    // Update Blazor
    if (dotNetRef) {
        const total = countUniqueFragments(section);
        reportArticleFragment(slideIndex, -1, total);
    }
}

// Disable fragments for a specific article slide
export function disableArticleSlideFragments(slideIndex) {
    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
    if (!slideData) return;
    const section = slideData.element.querySelector('section');
    if (!section) return;
    const fragments = section.querySelectorAll('.fragment');
    logDebug(`disableArticleSlideFragments: slide ${slideIndex}, disabling ${fragments.length} fragments`);
    fragments.forEach(f => {
        f.classList.remove('fragment');
        f.classList.add('fragment-off');
        f.classList.add('visible');
    });
    // Update Blazor
    if (dotNetRef) {
        const total = countUniqueFragments(section);
        reportArticleFragment(slideIndex, -1, total);
    }
}

function syncPresentationAudioToSlide(slideIndex) {
    const audio = document.getElementById('pres-audio-player');
    if (!audio) {
        return;
    }

    const source = audio.querySelector('source');
    const newSource = getSlideMediaUrl(slideIndex, false);
    if (!newSource) return;

    if (source && source.src.endsWith(newSource)) {
        return;
    }

    const wasPlaying = !audio.paused;
    audio.pause();
    audio.currentTime = 0;

    if (source) {
        source.src = newSource;
    } else {
        audio.src = newSource;
    }

    audio.load();

    if (wasPlaying) {
        audio.play();
    }

    if (dotNetRef) {
        dotNetRef.invokeMethodAsync('UpdatePlayingState', wasPlaying);
    }
}

// Audio Player Functions
export function playAudio(slideIndex) {
    logDebug(`playAudio: slide ${slideIndex}`);
    setupAudioListeners(); // Ensure listeners are attached
    
    const audio = document.getElementById(`audio-player-${slideIndex}`);
    if (audio) {
        // Pause all other audio
        document.querySelectorAll('audio').forEach(a => {
            if (a !== audio) a.pause();
        });
        audio.play();
        logTrace(`Started playing audio for slide ${slideIndex}`);
    } else {
        logWarning(`Audio element not found for slide ${slideIndex}`);
    }
}

export function pauseAudio(slideIndex) {
    logDebug(`pauseAudio: slide ${slideIndex}`);
    const audio = document.getElementById(`audio-player-${slideIndex}`);
    if (audio) {
        audio.pause();
    }
}

export function seekAudio(slideIndex, time) {
    logDebug(`seekAudio: slide ${slideIndex}, time ${time.toFixed(2)}s`);
    const audio = document.getElementById(`audio-player-${slideIndex}`);
    if (audio) {
        audio.currentTime = time;
    }
}

// Pause all article Reveal.js instances to prevent layout work while hidden
function pauseArticleRevealInstances() {
    for (const slideData of articleRevealInstances) {
        try {
            if (slideData.instance && typeof slideData.instance.destroy === 'function') {
                // Remove Reveal's internal resize/mutation observers by destroying
                // We'll re-init when switching back to article view
                slideData.instance.destroy();
            }
        } catch (e) {
            logWarning(`Error destroying article reveal instance ${slideData.index}`, e);
        }
    }
    articleRevealInstances.length = 0;
    logDebug('Article reveal instances destroyed for view switch');
}

export function stopAllMedia() {
    logInfo('Stopping all media');
    // Stop all audio
    document.querySelectorAll('audio').forEach(audio => {
        audio.pause();
        audio.currentTime = 0;
    });

    // Stop all video
    document.querySelectorAll('video').forEach(video => {
        video.pause();
        video.currentTime = 0;
    });
    
    // Stop YouTube
    if (youtubePlayer && youtubePlayerReady) {
        try {
            youtubePlayer.pauseVideo();
            logDebug('YouTube player paused');
        } catch (e) {
            logError('Error pausing YouTube player', e);
        }
    }

    // Stop YouTube time tracking interval to prevent stale timestamp updates
    if (videoTimeTracker) {
        clearInterval(videoTimeTracker);
        videoTimeTracker = null;
    }

    // Destroy article Reveal instances so they don't fire layout events while hidden
    pauseArticleRevealInstances();
    
    if (dotNetRef) {
        dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
    }
}

// Soft stop: pause all media WITHOUT resetting currentTime (used during view switch)
// This preserves playback positions so captureCurrentMediaState can read them.
export function softStopAllMedia() {
    logInfo('Soft-stopping all media (preserving positions)');
    document.querySelectorAll('audio').forEach(audio => {
        audio.pause();
    });
    document.querySelectorAll('video').forEach(video => {
        video.pause();
    });
    if (youtubePlayer && youtubePlayerReady) {
        try {
            youtubePlayer.pauseVideo();
        } catch (e) {
            logError('Error pausing YouTube player', e);
        }
    }

    // Stop YouTube time tracking interval to prevent stale timestamp updates during view switch
    if (videoTimeTracker) {
        clearInterval(videoTimeTracker);
        videoTimeTracker = null;
    }
}

/**
 * Captures the current media playback state across all views.
 * Call this BEFORE stopping media so we can restore position in the new view.
 * Returns { slideIndex, audioTimestamp, videoTimestamp, wasPlaying }
 */
export function captureCurrentMediaState() {
    let slideIndex = 0;
    let audioTimestamp = 0;
    let videoTimestamp = 0;
    let wasPlaying = false;

    // Check if any audio is playing or has a position
    const allAudio = Array.from(document.querySelectorAll('audio'));
    const playingAudio = allAudio.find(a => !a.paused);
    const audioWithPosition = allAudio.find(a => a.currentTime > 0);
    const activeAudio = playingAudio || audioWithPosition;

    if (playingAudio) {
        wasPlaying = true;
    }

    // Check YouTube player state
    if (youtubePlayer && youtubePlayerReady) {
        try {
            const state = youtubePlayer.getPlayerState();
            if (state === YT.PlayerState.PLAYING) {
                wasPlaying = true;
            }
            const ytTime = youtubePlayer.getCurrentTime();
            if (ytTime > 0) {
                videoTimestamp = ytTime;
            }
        } catch (e) {
            logWarning('Error reading YouTube state', e);
        }
    }

    // The page's own film is on the same clock as the YouTube upload of it.
    const film = getFilmPlayer();
    if (film) {
        if (!film.paused) {
            wasPlaying = true;
        }
        if (film.currentTime > 0) {
            videoTimestamp = film.currentTime;
        }
    }

    if (activeAudio) {
        const slideMatch = activeAudio.id?.match(/audio-player-(\d+)/);
        const isPresAudio = activeAudio.id === 'pres-audio-player';
        const isConsolidatedAudio = activeAudio.id === 'consolidated-audio-player';

        if (isConsolidatedAudio) {
            // Consolidated audio player (article/presentation views)
            slideIndex = consolidatedAudioSlideIndex;
            audioTimestamp = activeAudio.currentTime;
            const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
            if (notesEl) {
                const firstLine = notesEl.querySelector('[data-start-video]');
                if (firstLine) {
                    const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                    videoTimestamp = slideStartVideo + activeAudio.currentTime;
                }
            }
        } else if (slideMatch) {
            // Article view audio
            slideIndex = parseInt(slideMatch[1]);
            audioTimestamp = activeAudio.currentTime;
            // Calculate global video timestamp from slide-local audio time
            const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
            if (notesEl) {
                const firstLine = notesEl.querySelector('[data-start-video]');
                if (firstLine) {
                    const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                    videoTimestamp = slideStartVideo + activeAudio.currentTime;
                }
            }
        } else if (isPresAudio && presentationReveal) {
            // Presentation view audio
            const state = presentationReveal.getState();
            slideIndex = state.indexh + 1;
            audioTimestamp = activeAudio.currentTime;
            const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
            if (notesEl) {
                const firstLine = notesEl.querySelector('[data-start-video]');
                if (firstLine) {
                    const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                    videoTimestamp = slideStartVideo + activeAudio.currentTime;
                }
            }
        }
    }

    // Check consolidated video player state
    const consolidatedVideo = getConsolidatedVideo();
    if (consolidatedVideo && (slideIndex === 0 || playerVideoMode)) {
        if (!consolidatedVideo.paused) {
            wasPlaying = true;
        }
        if (consolidatedVideo.currentTime > 0 && playerVideoMode) {
            slideIndex = consolidatedVideoSlideIndex;
            audioTimestamp = consolidatedVideo.currentTime;
            const notesEl = document.getElementById(`slideNotes-${slideIndex}`);
            if (notesEl) {
                const firstLine = notesEl.querySelector('[data-start-video]');
                if (firstLine) {
                    const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
                    videoTimestamp = slideStartVideo + consolidatedVideo.currentTime;
                }
            }
        }
    }

    // If we got a video timestamp from YouTube but no slide info, find the slide from notes
    if (videoTimestamp > 0 && slideIndex === 0) {
        const noteLines = document.querySelectorAll('[data-start-video][data-end-video][data-slide-index]');
        for (const line of noteLines) {
            const startV = parseFloat(line.getAttribute('data-start-video'));
            const endV = parseFloat(line.getAttribute('data-end-video'));
            const si = parseInt(line.getAttribute('data-slide-index'));
            if (videoTimestamp >= startV && videoTimestamp <= endV) {
                slideIndex = si;
                // Derive audio timestamp: video time minus the first note's video start for that slide
                const slideNotes = document.getElementById(`slideNotes-${si}`);
                if (slideNotes) {
                    const fl = slideNotes.querySelector('[data-start-video]');
                    if (fl) {
                        audioTimestamp = videoTimestamp - (parseFloat(fl.getAttribute('data-start-video')) || 0);
                    }
                }
                break;
            }
        }
    }

    logInfo(`captureCurrentMediaState: slide=${slideIndex}, audio=${audioTimestamp.toFixed(2)}, video=${videoTimestamp.toFixed(2)}, playing=${wasPlaying}`);
    return { slideIndex, audioTimestamp, videoTimestamp, wasPlaying };
}

/**
 * Synchronizes a view to a specific slide and timestamp.
 * Called after switching views to restore playback position.
 * Mirrors the static site's syncViewToCurrentSlideWithData function.
 */
export function syncToView(viewId, slideIndex, audioTimestamp, videoTimestamp, autoPlay) {
    logInfo(`syncToView: view=${viewId}, slide=${slideIndex}, audio=${audioTimestamp.toFixed(2)}, video=${videoTimestamp.toFixed(2)}, autoPlay=${autoPlay}`);

    // Suppress note updates during the transition so stale timeupdate events don't cause flashing
    noteUpdateSuppressedUntil = Date.now() + NOTE_SUPPRESS_TRANSITION_MS;

    if (viewId === 'article') {
        // Scroll to the slide
        const slideEl = document.getElementById(`slide-${slideIndex}`);
        if (slideEl) {
            slideEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        // Seek consolidated media after a delay to let Reveal.js instances initialize
        setTimeout(() => {
            if (playerVideoMode) {
                const video = getConsolidatedVideo();
                if (video) {
                    syncConsolidatedVideoSource(slideIndex);
                    const trySeek = () => {
                        if (video.readyState >= 2) {
                            if (audioTimestamp > 0) {
                                video.currentTime = audioTimestamp;
                            }
                            noteUpdateSuppressedUntil = 0;
                            if (autoPlay) {
                                video.play().catch(e => logWarning(`syncToView article video play failed: ${e.message}`));
                            }
                        } else {
                            setTimeout(trySeek, 50);
                        }
                    };
                    trySeek();
                }
            } else {
                const audio = getConsolidatedAudio();
                if (audio) {
                    syncConsolidatedAudioSource(slideIndex);
                    const trySeek = () => {
                        if (audio.readyState >= 2) {
                            if (audioTimestamp > 0) {
                                audio.currentTime = audioTimestamp;
                            }
                            noteUpdateSuppressedUntil = 0;
                            if (autoPlay) {
                                audio.play().catch(e => logWarning(`syncToView article audio play failed: ${e.message}`));
                            }
                        } else {
                            setTimeout(trySeek, 50);
                        }
                    };
                    trySeek();
                }
            }
        }, 500);
    } else if (viewId === 'presentation') {
        if (presentationReveal) {
            presentationReveal.slide(slideIndex - 1);
            // Seek consolidated media after presentation navigation
            setTimeout(() => {
                if (playerVideoMode) {
                    const video = getConsolidatedVideo();
                    if (video) {
                        syncConsolidatedVideoSource(slideIndex);
                        const trySeek = () => {
                            if (video.readyState >= 2) {
                                logDebug(`syncToView: seeking consolidated video to ${audioTimestamp.toFixed(2)}s`);
                                if (audioTimestamp > 0) {
                                    video.currentTime = audioTimestamp;
                                }
                                noteUpdateSuppressedUntil = 0;
                                if (autoPlay) {
                                    video.play().catch(e => logWarning(`syncToView presentation video play failed: ${e.message}`));
                                }
                            } else {
                                setTimeout(trySeek, 50);
                            }
                        };
                        trySeek();
                    }
                } else {
                    const audio = getConsolidatedAudio();
                    if (audio) {
                        syncConsolidatedAudioSource(slideIndex);
                        const trySeek = () => {
                            if (audio.readyState >= 2) {
                                logDebug(`syncToView: seeking consolidated audio to ${audioTimestamp.toFixed(2)}s`);
                                if (audioTimestamp > 0) {
                                    audio.currentTime = audioTimestamp;
                                }
                                noteUpdateSuppressedUntil = 0;
                                if (autoPlay) {
                                    audio.play().catch(e => logWarning(`syncToView presentation audio play failed: ${e.message}`));
                                }
                            } else {
                                setTimeout(trySeek, 50);
                            }
                        };
                        trySeek();
                    }
                }
            }, 400);
        } else {
            logWarning('syncToView: presentation not initialized');
        }
    } else if (viewId === 'video') {
        // The page's own film: its clock is the page's film clock, so the video timestamp is the seek.
        const film = getFilmPlayer();
        if (film) {
            const trySeek = () => {
                if (film.readyState >= 1) {
                    film.currentTime = Math.max(0, videoTimestamp);
                    noteUpdateSuppressedUntil = 0;
                    if (autoPlay) {
                        film.play().catch(e => logWarning(`syncToView film play failed: ${e.message}`));
                    }
                    updateCurrentNoteFromVideoTime(film.currentTime);
                } else {
                    setTimeout(trySeek, 100);
                }
            };
            trySeek();
        } else {
            logWarning('syncToView: film player not found');
        }
    } else if (viewId === 'youtube') {
        const trySeek = () => {
            if (youtubePlayer && youtubePlayerReady) {
                logDebug(`syncToView: seeking YouTube to ${videoTimestamp.toFixed(2)}s`);
                youtubePlayer.seekTo(videoTimestamp, true);
                if (autoPlay) {
                    youtubePlayer.playVideo();
                } else {
                    // Explicitly pause to prevent auto-play after seek
                    setTimeout(() => {
                        try { youtubePlayer.pauseVideo(); } catch (e) {}
                    }, 100);
                }
                // Force highlight update after seek completes
                setTimeout(() => {
                    noteUpdateSuppressedUntil = 0;
                    try {
                        const actualTime = youtubePlayer.getCurrentTime();
                        updateCurrentNoteFromVideoTime(actualTime);
                    } catch (e) {}
                }, 300);
            } else {
                setTimeout(trySeek, 100);
            }
        };
        trySeek();
    }
}

// Scroll Functions
export function scrollToSlide(slideIndex) {
    logDebug(`scrollToSlide: ${slideIndex}`);
    const element = document.getElementById(`slide-${slideIndex}`);
    if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'start' });
        logTrace(`Scrolled to slide ${slideIndex}`);
    } else {
        logWarning(`Slide element not found: slide-${slideIndex}`);
    }
}

// Presentation View Functions
export async function initializePresentation(container, svgSources) {
    logInfo(`initializePresentation: ${svgSources.length} slides`);
    if (!container) {
        logError('Presentation container is null');
        return;
    }

    // Guard: if already initialized, just do a layout and return
    if (presentationReveal) {
        logDebug('Presentation already initialized, re-laying out');
        resizePresentationLayout();
        return;
    }
    
    // Ensure Reveal.js is loaded
    if (!revealLoaded) {
        logDebug('Reveal.js not yet loaded, loading now...');
        const loaded = await loadRevealJs();
        if (!loaded) {
            logError('Cannot initialize presentation - Reveal.js failed to load');
            return;
        }
    }
    
    // Initialize Reveal.js
    logDebug('Creating Reveal.js instance for presentation view');
    try {
        presentationReveal = new Reveal(container, {
            width: 1280,
            height: 720,
            controls: false,
            progress: true,
            hash: false,
            mouseWheel: true,
            embedded: true,
            // Never Reveal's scroll view. From 5.0 it swaps to it on its own whenever the slide's box is
            // narrower than 435px, and there a fragment is shown by scrolling rather than by slide() — so
            // a slide drawn narrow (the Slide Site Designer's center, a phone) lost every animation the
            // narration steps through.
            scrollActivationWidth: null,
            center: false,
            margin: 0,
            minScale: 0.1,
            maxScale: 2.0,
            plugins: [RevealZoom, RevealNotes, RevealSearch, RevealHighlight]
        });
        
        await presentationReveal.initialize();
        logInfo('Reveal.js presentation view initialized successfully');

        syncPresentationAudioToSlide(1);
        
        // Notify Blazor of initial fragment state for the first slide
        if (dotNetRef) {
            const firstSlide = presentationReveal.getCurrentSlide();
            const initialFragments = countUniqueFragments(firstSlide);
            dotNetRef.invokeMethodAsync('UpdatePresentationState', 0, -1, initialFragments);
        }
        
        const wrapper = document.getElementById('presentationViewWrapper');
        if (wrapper && !presentationResizeObserver) {
            presentationResizeObserver = new ResizeObserver(() => {
                resizePresentationLayout();
            });
            presentationResizeObserver.observe(wrapper);
        }

        resizePresentationLayout();
        if (!presentationResizeBound) {
            presentationResizeBound = true;
            window.addEventListener('resize', () => {
                resizePresentationLayout();
            });
        }
        
        presentationReveal.on('slidechanged', (event) => {
            logDebug(`Reveal slidechanged: index ${event.indexh}`);
            syncPresentationAudioToSlide(event.indexh + 1);
            if (dotNetRef) {
                const totalFragments = countUniqueFragments(event.currentSlide);
                dotNetRef.invokeMethodAsync('UpdatePresentationState', event.indexh, -1, totalFragments);
            }

            if (pendingPresentationSeek && pendingPresentationSeek.slideIndex === event.indexh + 1) {
                const seekTime = pendingPresentationSeek.time;
                pendingPresentationSeek = null;
                setTimeout(() => {
                    seekPresentationAudio(seekTime);
                }, 50);
            }
        });
        
        presentationReveal.on('fragmentshown', (event) => {
            if (dotNetRef) {
                const state = presentationReveal.getState();
                const currentSlide = presentationReveal.getCurrentSlide();
                const totalFragments = countUniqueFragments(currentSlide);
                logTrace(`Fragment shown: ${state.indexf}/${totalFragments}`);
                dotNetRef.invokeMethodAsync('UpdatePresentationState', state.indexh, state.indexf, totalFragments);
            }
        });
        
        presentationReveal.on('fragmenthidden', (event) => {
            if (dotNetRef) {
                const state = presentationReveal.getState();
                const currentSlide = presentationReveal.getCurrentSlide();
                const totalFragments = countUniqueFragments(currentSlide);
                logTrace(`Fragment hidden: ${state.indexf}/${totalFragments}`);
                dotNetRef.invokeMethodAsync('UpdatePresentationState', state.indexh, state.indexf, totalFragments);
            }
        });
    } catch (error) {
        logError('Error initializing Reveal.js presentation', error);
    }
}

export function navigatePresentationToSlide(index) {
    logDebug(`navigatePresentationToSlide: ${index}`);
    if (presentationReveal) {
        presentationReveal.slide(index);
    } else {
        logWarning('Presentation not initialized');
    }
}

export function navigatePresentationToSlideWithSeek(slideIndex, time) {
    pendingPresentationSeek = { slideIndex: slideIndex, time: time };
    navigatePresentationToSlide(slideIndex - 1);
}

export function navigatePresentationFragment(direction) {
    logDebug(`navigatePresentationFragment: ${direction > 0 ? 'next' : 'previous'}`);
    if (presentationReveal) {
        if (direction > 0) {
            presentationReveal.nextFragment();
        } else {
            presentationReveal.prevFragment();
        }
    } else {
        logWarning('Presentation not initialized');
    }
}

export function seekPresentationAudio(time) {
    logDebug(`seekPresentationAudio: ${time.toFixed(2)}s`);
    const audio = document.getElementById('pres-audio-player');
    if (audio) {
        setupAudioListeners();
        audio.currentTime = time;
        audio.play();
    } else {
        logWarning('Presentation audio player not found');
    }
}

export function togglePresentationAudio() {
    logDebug('togglePresentationAudio');
    const audio = document.getElementById('pres-audio-player');
    if (audio) {
        setupAudioListeners();
        if (audio.paused) {
            audio.play();
            if (dotNetRef) dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
        } else {
            audio.pause();
            if (dotNetRef) dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
        }
    } else {
        logWarning('Presentation audio player not found');
    }
}

// YouTube Player Functions
export function initializeYouTubePlayer(videoId) {
    logInfo(`initializeYouTubePlayer: ${videoId}`);
    const container = document.getElementById('video-player');
    if (!container) {
        logError('Video player container not found');
        return;
    }

    try {
        if (typeof YT === 'undefined' || typeof YT.Player === 'undefined') {
            logDebug('Waiting for YouTube API to load');
            // Wait for API to load
            window.onYouTubeIframeAPIReady = () => {
                logDebug('YouTube API ready');
                createYouTubePlayer(container, videoId);
            };
        } else {
            createYouTubePlayer(container, videoId);
        }
    } catch (error) {
        logError('Error initializing YouTube player', error);
    }
}

function createYouTubePlayer(container, videoId) {
    logDebug(`Creating YouTube player for video: ${videoId}`);
    container.innerHTML = '<div id="yt-player"></div>';

    try {
        youtubePlayer = new YT.Player('yt-player', {
            videoId: videoId,
            width: '100%',
            height: '100%',
            playerVars: {
                'autoplay': 0,
                'playsinline': 1,
                'enablejsapi': 1,
                'origin': window.location.origin
            },
            events: {
                'onReady': () => {
                    youtubePlayerReady = true;
                    logInfo('YouTube player ready');
                    applyYouTubeSettings();
                },
                'onStateChange': (event) => {
                    const stateNames = {
                        [-1]: 'unstarted',
                        [0]: 'ended',
                        [1]: 'playing',
                        [2]: 'paused',
                        [3]: 'buffering',
                        [5]: 'cued'
                    };
                    logDebug(`YouTube state changed: ${stateNames[event.data] || event.data}`);
                    
                    if (dotNetRef) {
                        const isPlaying = event.data === YT.PlayerState.PLAYING;
                        dotNetRef.invokeMethodAsync('UpdatePlayingState', isPlaying);
                        
                        if (isPlaying) {
                            startVideoTimeTracking();
                        }
                    }
                },
                'onError': (event) => {
                    logError(`YouTube player error: ${event.data}`);
                }
            }
        });
    } catch (error) {
        logError('Error creating YouTube player', error);
    }
}

function startVideoTimeTracking() {
    logTrace('Starting video time tracking');
    if (videoTimeTracker) {
        clearInterval(videoTimeTracker);
    }
    
    videoTimeTracker = setInterval(() => {
        if (youtubePlayer && youtubePlayerReady && dotNetRef) {
            try {
                // Guard: only update when YouTube is actively playing to prevent
                // stale timestamp updates if the interval survives a stop
                const state = youtubePlayer.getPlayerState();
                if (state !== YT.PlayerState.PLAYING) return;

                const currentTime = youtubePlayer.getCurrentTime();
                updateProgressBarDirect(currentTime);

                const now = Date.now();
                if (now - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
                    lastTimestampUpdate = now;
                    dotNetRef.invokeMethodAsync('UpdateTimestamp', currentTime);
                }

                updateCurrentNoteFromVideoTime(currentTime);
            } catch (e) {
                logError('Error getting YouTube current time', e);
            }
        }
    }, 100);
}

function updateCurrentNoteFromVideoTime(currentTime) {
    const lines = document.querySelectorAll('[data-start-video][data-end-video]');
    for (const line of lines) {
        const startVideo = parseFloat(line.getAttribute('data-start-video'));
        const endVideo = parseFloat(line.getAttribute('data-end-video'));
        if (currentTime >= startVideo && currentTime <= endVideo) {
            const noteId = line.getAttribute('data-note-id') || line.id;
            const slideIndex = parseInt(line.getAttribute('data-slide-index'));
            invokeCurrentNoteUpdate(noteId, slideIndex, currentTime);
            break;
        }
    }
}

export function seekYouTubeVideo(time) {
    logDebug(`seekYouTubeVideo: ${time.toFixed(2)}s`);
    if (youtubePlayer && youtubePlayerReady) {
        youtubePlayer.seekTo(time, true);
        youtubePlayer.playVideo();
    } else {
        logWarning('YouTube player not ready');
    }
}

// Film Player Functions (the Video view: the page's own film, played in the page)

function getFilmPlayer() {
    return document.getElementById('film-player');
}

/**
 * Follows the page's film: the progress bar, the timestamp .NET shows and the line being said, all on
 * the film's clock — the same three things the YouTube tracker drives, from the element's own events
 * rather than a timer, because a <video> says when its time moves.
 */
export function initializeFilmPlayer() {
    const film = getFilmPlayer();
    if (!film) {
        logWarning('initializeFilmPlayer: film player not found');
        return;
    }

    if (filmPlayerBound === film) return;
    filmPlayerBound = film;

    // The page's volume and speed, and the film's own controls reported back - see onMediaSettingChanged.
    Object.entries(mediaSettingHandlers).forEach(([name, handler]) => film.addEventListener(name, handler));
    applyMediaSettings(film);

    film.addEventListener('timeupdate', () => {
        const currentTime = film.currentTime;
        updateProgressBarDirect(currentTime);

        const now = Date.now();
        if (dotNetRef && now - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
            lastTimestampUpdate = now;
            dotNetRef.invokeMethodAsync('UpdateTimestamp', currentTime);
        }

        updateCurrentNoteFromVideoTime(currentTime);
    });

    film.addEventListener('play', () => {
        if (dotNetRef) dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
    });

    film.addEventListener('pause', () => {
        if (dotNetRef) dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
    });

    logInfo('Film player initialized');
}

export function seekFilm(time) {
    const film = getFilmPlayer();
    if (!film || !Number.isFinite(time)) return;

    film.currentTime = Math.max(0, time);
    const played = film.play();
    if (played && played.catch) played.catch(e => logWarning(`seekFilm play failed: ${e.message}`));
}

export function toggleFilmPlayback() {
    const film = getFilmPlayer();
    if (!film) return;

    if (film.paused) {
        const played = film.play();
        if (played && played.catch) played.catch(e => logWarning(`toggleFilmPlayback play failed: ${e.message}`));
    } else {
        film.pause();
    }
}

export function toggleYouTubePlayback() {
    logDebug('toggleYouTubePlayback');
    if (youtubePlayer && youtubePlayerReady) {
        const state = youtubePlayer.getPlayerState();
        if (state === YT.PlayerState.PLAYING) {
            youtubePlayer.pauseVideo();
        } else {
            youtubePlayer.playVideo();
        }
    } else {
        logWarning('YouTube player not ready');
    }
}

export function scrollSidebarNoteIntoView(slideIndex, noteId) {
    requestAnimationFrame(() => {
        const noteElement = document.getElementById(`slideNoteStack-${slideIndex}-${noteId}`);
        if (!noteElement) {
            return;
        }

        const scrollable = document.querySelector('.mainview-content-nav-scrollable');
        if (scrollable) {
            const slideNavBlock = noteElement.closest('.slide-nav-block');
            const stickyHeader = slideNavBlock ? slideNavBlock.querySelector('.slide-nav-block-header') : null;
            const stickyHeaderHeight = stickyHeader ? stickyHeader.offsetHeight : 0;
            const noteTop = noteElement.getBoundingClientRect().top;
            const scrollableTop = scrollable.getBoundingClientRect().top;
            const offset = noteTop - scrollableTop + scrollable.scrollTop - stickyHeaderHeight;
            scrollable.scrollTo({ top: offset, behavior: 'smooth' });
        } else {
            noteElement.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    });
}

// Article View Functions - Initialize Reveal.js for individual slide cards
// Note: This function expects the reveal structure to already exist in the DOM (created by Razor)
// It does NOT recreate the structure - it only initializes Reveal.js on the existing elements
export async function initializeArticleSlide(slideIndex, containerId, svgContent, customWidth, customHeight) {
    logDebug(`initializeArticleSlide: slide ${slideIndex}, container ${containerId}`);
    
    // Ensure Reveal.js is loaded
    if (!revealLoaded) {
        logDebug('Reveal.js not yet loaded, loading now...');
        const loaded = await loadRevealJs();
        if (!loaded) {
            logError(`Cannot initialize article slide ${slideIndex} - Reveal.js failed to load`);
            return null;
        }
    }
    
    const element = document.getElementById(containerId);
    if (!element) {
        logWarning(`Article slide container not found: ${containerId}`);
        return null;
    }
    
    logTrace(`Found container element for slide ${slideIndex}`);
    
    // The element should already have the reveal class, slides structure, and inline styles from Razor
    // We only set position relative for Reveal.js absolute positioning of children
    // DO NOT override width/height - those are controlled by CSS classes (.slide-size-small, etc.)
    element.style.position = 'relative';
    element.style.overflow = 'hidden';
    // Match static site behavior: use zIndex for proper layering
    element.style.zIndex = '1100';
    element.style.marginTop = '0';
    
    // Get custom width/height from section data attributes (matching static site)
    const section = element.querySelector('section');
    const sectionWidth = section?.getAttribute('data-width');
    const sectionHeight = section?.getAttribute('data-height');
    
    try {
        const revealInstance = new Reveal(element, {
            embedded: true,
            // Never Reveal's scroll view. From 5.0 it swaps to it on its own whenever the slide's box is
            // narrower than 435px, and there a fragment is shown by scrolling rather than by slide() — so
            // a slide drawn narrow (the Slide Site Designer's center, a phone) lost every animation the
            // narration steps through.
            scrollActivationWidth: null,
            // Reveal lays nothing out here: the article sizes its slides with CSS (see the overrides
            // below and layoutAllArticleSlides). Left on, Reveal's own layout ran inside every slide()
            // call — which the narration makes several times a second — and put back its fixed-size,
            // scaled `.slides` each time, so the picture jumped between the two sizes while it played.
            disableLayout: true,
            width: sectionWidth ? parseInt(sectionWidth) : (customWidth || 1280),
            height: sectionHeight ? parseInt(sectionHeight) : (customHeight || 720),
            margin: 0,
            minScale: 0.1,
            maxScale: 2.0,
            keyboardCondition: 'focused',
            controls: false,
            progress: false,
            hash: false,
            showNotes: false,
            slideNumber: false,
            mouseWheel: false,
            keyboard: true,
            plugins: [RevealZoom, RevealNotes, RevealSearch, RevealHighlight]
        });
        
        await revealInstance.initialize();
        logInfo(`Article slide ${slideIndex} Reveal.js initialized successfully`);
        
        // Reveal.js sets inline width/height on the .reveal element during layout.
        // In embedded article mode the parent has no fixed height, so Reveal computes 0.
        // Remove the inline height and let CSS (height: auto + aspect-ratio) size the container.
        element.style.height = 'auto';
        element.style.width = '';
        
        // Also fix the .slides wrapper - Reveal.js sets inline styles on it
        const slidesWrapper = element.querySelector('.slides');
        if (slidesWrapper) {
            slidesWrapper.style.width = '100%';
            slidesWrapper.style.height = '100%';
            slidesWrapper.style.left = '0';
            slidesWrapper.style.top = '0';
            slidesWrapper.style.transform = 'none';
        }
        
        // Force visibility of slides after Reveal.js initialization
        // Reveal.js applies inline styles that hide non-present slides
        const sections = element.querySelectorAll('.slides > section');
        sections.forEach(section => {
            section.style.display = 'block';
            section.style.opacity = '1';
            section.style.visibility = 'visible';
            section.style.position = 'absolute';
            section.style.top = '0';
            section.style.left = '0';
            section.style.width = '100%';
            section.style.height = '100%';
            section.style.transform = 'none';
        });
        
        // Clear any inline styles Reveal.js set on r-stretch elements (SVGs/images)
        element.querySelectorAll('.r-stretch').forEach(el => {
            el.style.width = '';
            el.style.height = '';
            el.style.maxWidth = '';
            el.style.maxHeight = '';
        });
        
        // Store the instance for later reference
        articleRevealInstances.push({ index: slideIndex, instance: revealInstance, element });
        
        // Add fragment event listeners to update Blazor state
        const sectionForFragments = element.querySelector('section');
        const totalFragments = countUniqueFragments(sectionForFragments);
        // Notify Blazor of initial fragment count
        if (dotNetRef && totalFragments > 0) {
            reportArticleFragment(slideIndex, -1, totalFragments);
        }
        
        revealInstance.on('fragmentshown', () => {
            if (dotNetRef) {
                const state = revealInstance.getState();
                const frags = countUniqueFragments(element.querySelector('section'));
                logTrace(`Article slide ${slideIndex} fragment shown: ${state.indexf}/${frags}`);
                reportArticleFragment(slideIndex, state.indexf, frags);
            }
        });
        
        revealInstance.on('fragmenthidden', () => {
            if (dotNetRef) {
                const state = revealInstance.getState();
                const frags = countUniqueFragments(element.querySelector('section'));
                logTrace(`Article slide ${slideIndex} fragment hidden: ${state.indexf}/${frags}`);
                reportArticleFragment(slideIndex, state.indexf, frags);
            }
        });
        
        // Return the instance info
        return {
            index: slideIndex,
            totalFragments: totalFragments
        };
    } catch (error) {
        logError(`Error initializing Reveal.js for article slide ${slideIndex}`, error);
        return null;
    }
}

// Initialize all article slides from an array of slide data
export async function initializeArticleView(slides) {
    logInfo(`initializeArticleView: ${slides.length} slides`);
    
    // Ensure Reveal.js is loaded
    if (!revealLoaded) {
        logDebug('Reveal.js not yet loaded, loading now...');
        const loaded = await loadRevealJs();
        if (!loaded) {
            logError('Cannot initialize article view - Reveal.js failed to load');
            return [];
        }
    }
    
    // Clear existing instances
    articleRevealInstances.length = 0;

    if (!articleResizeBound) {
        articleResizeBound = true;
        window.addEventListener('resize', scheduleArticleRelayout);
    }

    const results = [];
    for (const slide of slides) {
        // Use the containerId from the slide data (matches what Razor generates)
        const containerId = slide.containerId;
        const element = document.getElementById(containerId);
        
        if (!element) {
            logDebug(`Article slide container not yet in DOM: ${containerId}`);
            continue;
        }
        
        logTrace(`Initializing article slide ${slide.index}`);
        
        // The element already has the reveal class, slides structure, and inline styles from Razor
        // We only set position relative for Reveal.js absolute positioning of children
        // DO NOT override width/height - those are controlled by CSS classes (.slide-size-small, etc.)
        element.style.position = 'relative';
        element.style.overflow = 'hidden';
        // Match static site behavior: use zIndex for proper layering
        element.style.zIndex = '1100';
        element.style.marginTop = '0';
        
        // Get custom width/height from section data attributes (matching static site)
        const section = element.querySelector('section');

        try {
            const revealInstance = new Reveal(element, {
                embedded: true,
                // Never Reveal's scroll view. From 5.0 it swaps to it on its own whenever the slide's box is
                // narrower than 435px, and there a fragment is shown by scrolling rather than by slide() — so
                // a slide drawn narrow (the Slide Site Designer's center, a phone) lost every animation the
                // narration steps through.
                scrollActivationWidth: null,
                // Reveal lays nothing out here: the article sizes its slides with CSS (see the overrides
                // below and layoutAllArticleSlides). Left on, Reveal's own layout ran inside every slide()
                // call — which the narration makes several times a second — and put back its fixed-size,
                // scaled `.slides` each time, so the picture jumped between the two sizes while it played.
                disableLayout: true,
                width: slide.width,
                height: slide.height,
                margin: 0,
                minScale: 0.1,
                maxScale: 2.0,
                keyboardCondition: 'focused',
                controls: false,
                progress: false,
                hash: false,
                showNotes: false,
                slideNumber: false,
                mouseWheel: false,
                keyboard: true,
                plugins: [RevealZoom, RevealNotes, RevealSearch, RevealHighlight]
            });

            await revealInstance.initialize();
            logDebug(`Article slide ${slide.index} initialized`);
            
            // Reveal.js sets inline width/height on the .reveal element during layout.
            // In embedded article mode the parent has no fixed height, so Reveal computes 0.
            // Remove the inline height and let CSS (height: auto + aspect-ratio) size the container.
            element.style.height = 'auto';
            element.style.width = '';
            
            // Also fix the .slides wrapper - Reveal.js sets inline styles on it
            const slidesWrapper = element.querySelector('.slides');
            if (slidesWrapper) {
                slidesWrapper.style.width = '100%';
                slidesWrapper.style.height = '100%';
                slidesWrapper.style.left = '0';
                slidesWrapper.style.top = '0';
                slidesWrapper.style.transform = 'none';
            }
            
            // Force visibility of slides after Reveal.js initialization
            // Reveal.js applies inline styles that hide non-present slides
            const sections = element.querySelectorAll('.slides > section');
            sections.forEach(section => {
                section.style.display = 'block';
                section.style.opacity = '1';
                section.style.visibility = 'visible';
                section.style.position = 'absolute';
                section.style.top = '0';
                section.style.left = '0';
                section.style.width = '100%';
                section.style.height = '100%';
                section.style.transform = 'none';
            });
            
            // Clear any inline styles Reveal.js set on r-stretch elements (SVGs/images)
            element.querySelectorAll('.r-stretch').forEach(el => {
                el.style.width = '';
                el.style.height = '';
                el.style.maxWidth = '';
                el.style.maxHeight = '';
            });
            
            articleRevealInstances.push({ index: slide.index, instance: revealInstance, element });
            
            // Add fragment event listeners to update Blazor state
            const slideIndex = slide.index;
            const section2 = element.querySelector('section');
            const totalFragments = countUniqueFragments(section2);
            // Notify Blazor of initial fragment count
            if (dotNetRef && totalFragments > 0) {
                reportArticleFragment(slideIndex, -1, totalFragments);
            }
            
            revealInstance.on('fragmentshown', () => {
                if (dotNetRef) {
                    const state = revealInstance.getState();
                    const frags = countUniqueFragments(element.querySelector('section'));
                    logTrace(`Article slide ${slideIndex} fragment shown: ${state.indexf}/${frags}`);
                    reportArticleFragment(slideIndex, state.indexf, frags);
                }
            });
            
            revealInstance.on('fragmenthidden', () => {
                if (dotNetRef) {
                    const state = revealInstance.getState();
                    const frags = countUniqueFragments(element.querySelector('section'));
                    logTrace(`Article slide ${slideIndex} fragment hidden: ${state.indexf}/${frags}`);
                    reportArticleFragment(slideIndex, state.indexf, frags);
                }
            });
            
            results.push({ index: slide.index, success: true });
        } catch (error) {
            logError(`Error initializing article slide ${slide.index}`, error);
            results.push({ index: slide.index, success: false, error: error.message });
        }
    }
    
    logInfo(`Article view initialized: ${results.filter(r => r.success).length}/${slides.length} slides`);

    // If no elements were found, the Blazor render batch may not have been
    // flushed to the DOM yet. Schedule a non-blocking background retry so
    // the C# caller is not blocked and slides initialize when they appear.
    if (results.length === 0 && slides.length > 0) {
        if (_articleRetryTimer) clearInterval(_articleRetryTimer);
        let attempt = 0;
        _articleRetryTimer = setInterval(async () => {
            attempt++;
            const firstEl = document.getElementById(slides[0].containerId);
            if (firstEl) {
                clearInterval(_articleRetryTimer);
                _articleRetryTimer = null;
                logInfo(`Article slide elements appeared on retry #${attempt}, initializing`);
                await initializeArticleView(slides);
            } else if (attempt >= 40) {
                clearInterval(_articleRetryTimer);
                _articleRetryTimer = null;
                logWarning(`Article slide elements not found after ${attempt} retries`);
            }
        }, 250);
    }

    return results;
}

// Navigate fragment in a specific article slide
export function navigateArticleSlideFragment(slideIndex, direction) {
    logDebug(`navigateArticleSlideFragment: slide ${slideIndex}, direction ${direction > 0 ? 'next' : 'previous'}`);
    
    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
    if (slideData && slideData.instance) {
        if (direction > 0) {
            slideData.instance.nextFragment();
        } else {
            slideData.instance.prevFragment();
        }
        
        const state = slideData.instance.getState();
        const totalFragments = countUniqueFragments(slideData.element.querySelector('section'));
        logTrace(`Article slide ${slideIndex} fragment: ${state.indexf}/${totalFragments}`);
        
        // Also notify Blazor to update the display
        if (dotNetRef) {
            reportArticleFragment(slideIndex, state.indexf, totalFragments);
        }
        
        return { fragmentIndex: state.indexf, totalFragments };
    } else {
        logWarning(`Article slide instance not found: ${slideIndex}`);
        return null;
    }
}

// Get article slide fragment state
export function getArticleSlideFragmentState(slideIndex) {
    const slideData = articleRevealInstances.find(s => s.index === slideIndex);
    if (slideData && slideData.instance) {
        const state = slideData.instance.getState();
        const totalFragments = slideData.element.querySelectorAll('.fragment').length;
        return { fragmentIndex: state.indexf, totalFragments };
    }
    return null;
}

// Re-layout all article slide Reveal.js instances (call after size/width changes)
export function layoutAllArticleSlides() {
    logDebug(`layoutAllArticleSlides: ${articleRevealInstances.length} instances`);
    for (const slideData of articleRevealInstances) {
        try {
            // Reset inline width/height that Reveal.js may have set
            slideData.element.style.height = 'auto';
            slideData.element.style.width = '';
            
            slideData.instance.layout();
            
            // Re-apply the CSS overrides after layout
            slideData.element.style.height = 'auto';
            slideData.element.style.width = '';
            
            const slidesWrapper = slideData.element.querySelector('.slides');
            if (slidesWrapper) {
                slidesWrapper.style.width = '100%';
                slidesWrapper.style.height = '100%';
                slidesWrapper.style.left = '0';
                slidesWrapper.style.top = '0';
                slidesWrapper.style.transform = 'none';
            }
            
            // Force visibility of sections
            const sections = slideData.element.querySelectorAll('.slides > section');
            sections.forEach(section => {
                section.style.display = 'block';
                section.style.opacity = '1';
                section.style.visibility = 'visible';
                section.style.position = 'absolute';
                section.style.top = '0';
                section.style.left = '0';
                section.style.width = '100%';
                section.style.height = '100%';
                section.style.transform = 'none';
            });
            
            logTrace(`Relayout complete for slide ${slideData.index}`);
        } catch (e) {
            logWarning(`Failed to layout slide ${slideData.index}`, e);
        }
    }
}

// Global callback for YouTube API
window.onYouTubeIframeAPIReady = window.onYouTubeIframeAPIReady || function() {
    // API loaded, player will be created when initializeYouTubePlayer is called
};

// ===== Progress Bar Direct Update (from JS, no Blazor round-trip) =====
function updateProgressBarDirect(currentTime) {
    if (cachedTotalDuration <= 0) return;
    
    const percentage = (currentTime / cachedTotalDuration) * 100;
    const fill = document.getElementById('progressBarFill');
    const handle = document.getElementById('progressBarHandle');
    
    if (fill) fill.style.width = `${percentage}%`;
    if (handle) handle.style.left = `${percentage}%`;
}

// ===== Progress Bar Thumbnail Preview =====
// Set up hover interactions on the custom progress bar to show thumbnail + tooltip
export function setupProgressBarInteractions(totalDurationMs, slideCount, siteBaseUrl) {
    logDebug(`setupProgressBarInteractions: totalDuration=${totalDurationMs}, slides=${slideCount}, siteBaseUrl=${siteBaseUrl}`);
    
    // Cache for JS-side progress updates
    cachedTotalDuration = totalDurationMs;
    
    const container = document.getElementById('progressBarContainer');
    const fill = document.getElementById('progressBarFill');
    const handle = document.getElementById('progressBarHandle');
    const thumbnailPreview = document.getElementById('progressThumbnailPreview');
    const thumbnailImage = document.getElementById('progressThumbnailImage');
    const hoverTooltip = document.getElementById('progressHoverTooltip');
    
    if (!container) {
        logWarning('Progress bar container not found');
        return;
    }
    
    const totalDuration = totalDurationMs;
    let isDragging = false;
    let wasPlayingBeforeDrag = false;
    let lastDragSeekTime = 0;
    let lastDragSlideIndex = 1;
    let lastDragAudioTime = 0;
    
    // Collect slide timing data from the DOM
    function getSlideForTime(hoverTime) {
        const allLines = document.querySelectorAll('[data-start-video][data-slide-index]');
        let thumbnailSlideIndex = 1;
        for (let i = allLines.length - 1; i >= 0; i--) {
            const startVideo = parseFloat(allLines[i].getAttribute('data-start-video'));
            const slideIndex = parseInt(allLines[i].getAttribute('data-slide-index'));
            if (hoverTime >= startVideo && slideIndex) {
                thumbnailSlideIndex = slideIndex;
                break;
            }
        }
        return thumbnailSlideIndex;
    }
    
    function formatTimestampLocal(seconds) {
        if (isNaN(seconds) || seconds < 0) return '0:00';
        const hours = Math.floor(seconds / 3600);
        const minutes = Math.floor((seconds % 3600) / 60);
        const secs = Math.floor(seconds % 60);
        return hours > 0
            ? `${hours}:${String(minutes).padStart(2, '0')}:${String(secs).padStart(2, '0')}`
            : `${minutes}:${String(secs).padStart(2, '0')}`;
    }

    function resolveSeekTarget(e) {
        const rect = container.getBoundingClientRect();
        const clickX = e.clientX - rect.left;
        const percentage = Math.max(0, Math.min(1, clickX / rect.width));
        const seekTime = percentage * totalDuration;

        let targetSlideIndex = 1;
        let targetAudioTime = seekTime;

        const allLines = document.querySelectorAll('[data-start-video][data-slide-index]');
        for (let i = allLines.length - 1; i >= 0; i--) {
            const startVideo = parseFloat(allLines[i].getAttribute('data-start-video'));
            const slideIndex = parseInt(allLines[i].getAttribute('data-slide-index'));
            if (seekTime >= startVideo && slideIndex) {
                targetSlideIndex = slideIndex;
                const slideLines = document.querySelectorAll(`[data-slide-index="${slideIndex}"][data-start-video]`);
                if (slideLines.length > 0) {
                    const slideStart = parseFloat(slideLines[0].getAttribute('data-start-video'));
                    targetAudioTime = seekTime - slideStart;
                }
                break;
            }
        }

        return { percentage, seekTime, targetSlideIndex, targetAudioTime };
    }

    function updateDragVisuals(percentage, seekTime) {
        if (fill) fill.style.width = `${percentage * 100}%`;
        if (handle) handle.style.left = `${percentage * 100}%`;
    }

    function seekToPosition(e) {
        const { percentage, seekTime, targetSlideIndex, targetAudioTime } = resolveSeekTarget(e);

        updateDragVisuals(percentage, seekTime);
        
        // Scroll to slide and seek media
        const slideEl = document.getElementById(`slide-${targetSlideIndex}`);
        if (slideEl) {
            slideEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
        if (playerVideoMode) {
            seekConsolidatedVideo(targetSlideIndex, targetAudioTime);
        } else {
            seekConsolidatedAudio(targetSlideIndex, targetAudioTime);
        }
        
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('ExpandSlideNavCard', targetSlideIndex);
            dotNetRef.invokeMethodAsync('UpdateTimestamp', seekTime);
        }
    }
    
    container.addEventListener('mousedown', (e) => {
        isDragging = true;

        // Pause media during drag
        const activeMedia = playerVideoMode ? getConsolidatedVideo() : getConsolidatedAudio();
        wasPlayingBeforeDrag = activeMedia ? !activeMedia.paused : false;
        if (activeMedia && !activeMedia.paused) {
            activeMedia.pause();
        }

        // Only update visuals on initial mousedown, defer seek to mouseup
        const { percentage, seekTime, targetSlideIndex, targetAudioTime } = resolveSeekTarget(e);
        updateDragVisuals(percentage, seekTime);
        lastDragSeekTime = seekTime;
        lastDragSlideIndex = targetSlideIndex;
        lastDragAudioTime = targetAudioTime;

        e.preventDefault();
    });
    
    document.addEventListener('mousemove', (e) => {
        if (isDragging) {
            // Only update visuals during drag, defer actual seek to mouseup
            const { percentage, seekTime, targetSlideIndex, targetAudioTime } = resolveSeekTarget(e);
            updateDragVisuals(percentage, seekTime);
            lastDragSeekTime = seekTime;
            lastDragSlideIndex = targetSlideIndex;
            lastDragAudioTime = targetAudioTime;
        }
        
        // Show thumbnail preview on hover
        const containerEl = document.getElementById('progressBarContainer');
        if (containerEl && (e.target === containerEl || e.target.closest('#progressBarContainer'))) {
            const rect = containerEl.getBoundingClientRect();
            const hoverX = e.clientX - rect.left;
            const percentage = Math.max(0, Math.min(1, hoverX / rect.width));
            const hoverTime = percentage * totalDuration;
            
            const slideIdx = getSlideForTime(hoverTime);

            // The slide's own picture from the media map - a published page names its pictures for the
            // page (images/Matthew-9-Slide1.png), so images/Slide{n}.png was a 404 on every hover.
            const thumbnailSrc = (slideMediaMap[slideIdx] && slideMediaMap[slideIdx].png) || '';

            if (thumbnailImage && thumbnailSrc && thumbnailImage.src !== thumbnailSrc) {
                thumbnailImage.src = thumbnailSrc;
            }
            if (hoverTooltip) {
                hoverTooltip.textContent = formatTimestampLocal(hoverTime);
            }
            
            if (thumbnailPreview) {
                const thumbWidth = 240;
                let xPos = hoverX;
                if (xPos < thumbWidth / 2) xPos = thumbWidth / 2;
                if (xPos > rect.width - thumbWidth / 2) xPos = rect.width - thumbWidth / 2;
                
                thumbnailPreview.style.left = `${xPos}px`;
                thumbnailPreview.style.opacity = thumbnailSrc ? '1' : '0';
            }
            if (hoverTooltip) {
                hoverTooltip.style.opacity = '1';
            }
        }
    });
    
    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;

            // Perform the actual seek at the final drag position
            const slideEl = document.getElementById(`slide-${lastDragSlideIndex}`);
            if (slideEl) {
                slideEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }

            if (playerVideoMode) {
                seekConsolidatedVideo(lastDragSlideIndex, lastDragAudioTime);
            } else {
                seekConsolidatedAudio(lastDragSlideIndex, lastDragAudioTime);
            }

            if (dotNetRef) {
                dotNetRef.invokeMethodAsync('ExpandSlideNavCard', lastDragSlideIndex);
                dotNetRef.invokeMethodAsync('UpdateTimestamp', lastDragSeekTime);
            }

            // Resume playback if it was playing before drag
            if (wasPlayingBeforeDrag) {
                const activeMedia = playerVideoMode ? getConsolidatedVideo() : getConsolidatedAudio();
                if (activeMedia) {
                    const tryPlay = () => {
                        if (activeMedia.readyState >= 2) {
                            activeMedia.play().catch(e => logWarning(`Progress bar resume play failed: ${e.message}`));
                        } else {
                            setTimeout(tryPlay, 50);
                        }
                    };
                    tryPlay();
                }
            }
        }
    });
    
    container.addEventListener('mouseenter', () => {
        container.style.height = '6px';
        if (handle) handle.style.opacity = '1';
    });
    
    container.addEventListener('mouseleave', () => {
        if (!isDragging) {
            container.style.height = '4px';
            if (handle) handle.style.opacity = '0';
        }
        if (thumbnailPreview) thumbnailPreview.style.opacity = '0';
        if (hoverTooltip) hoverTooltip.style.opacity = '0';
    });
    
    logInfo('Progress bar interactions setup complete');
}

// ===== Sidebar Thumbnail Hover Preview =====
// Set up hover on slideNavThumbnail images to show enlarged preview in the main area
export function setupSidebarThumbnailHover(siteBaseUrl) {
    logDebug(`setupSidebarThumbnailHover: siteBaseUrl=${siteBaseUrl}`);
    
    const popup = document.getElementById('slideNavThumbnailPopup');
    const popupImage = document.getElementById('slideNavThumbnailPopupImage');
    
    if (!popup || !popupImage) {
        logWarning('Sidebar thumbnail popup elements not found');
        return;
    }
    
    // Find all sidebar thumbnail images
    const thumbnails = document.querySelectorAll('[id^="slideNavThumbnail-"]');
    logDebug(`Found ${thumbnails.length} sidebar thumbnails`);
    
    thumbnails.forEach(thumbnail => {
        // Get slide index from the id (e.g., slideNavThumbnail-1)  
        const idMatch = thumbnail.id?.match(/slideNavThumbnail-(\d+)/);
        if (!idMatch) return;
        const slideIndex = parseInt(idMatch[1]);
        
        // Find the actual img inside the RadzenImage component
        const imgEl = thumbnail.tagName === 'IMG' ? thumbnail : thumbnail.querySelector('img');
        if (!imgEl) return;
        
        imgEl.addEventListener('mouseenter', (e) => {
            const src = imgEl.src;
            popupImage.src = src;
            
            // Calculate available space
            // Align top of popup to the bottom of the top toolbar
            const topToolbar = document.getElementById('topToolbar');
            const topToolbarHeight = topToolbar ? topToolbar.getBoundingClientRect().bottom : 50;
            const rightMargin = 10;
            const viewportWidth = window.innerWidth;
            const viewportHeight = window.innerHeight;
            
            // Estimate sidebar width (the page nav sidebar)
            const sidebarPane = thumbnail.closest('.page-nav-sidebar');
            const sidebarWidth = sidebarPane ? sidebarPane.offsetWidth : viewportWidth * 0.20;
            
            // Available space for the popup (to the left of the sidebar)
            const borderPadding = 18; // 5px border * 2 + 4px padding * 2
            const availableWidth = viewportWidth - sidebarWidth - rightMargin - borderPadding;
            const availableHeight = viewportHeight - topToolbarHeight - borderPadding;
            
            // Desired 16:9 aspect ratio
            const aspectRatio = 1280 / 720;
            
            let finalWidth = Math.min(1280, availableWidth);
            let finalHeight = finalWidth / aspectRatio;
            
            if (finalHeight > availableHeight) {
                finalHeight = availableHeight;
                finalWidth = finalHeight * aspectRatio;
            }
            
            popupImage.style.width = `${finalWidth}px`;
            popupImage.style.height = `${finalHeight}px`;
            
            const popupWidth = finalWidth + borderPadding;
            const leftPosition = viewportWidth - sidebarWidth - popupWidth - rightMargin;
            
            popup.style.left = `${leftPosition}px`;
            popup.style.top = `${topToolbarHeight}px`;
            popup.classList.add('visible');
        });
        
        imgEl.addEventListener('mouseleave', () => {
            popup.classList.remove('visible');
        });
    });
    
    logInfo('Sidebar thumbnail hover setup complete');
}

/**
 * Performs a full view switch with media state preservation.
 * Captures current playback state, stops media, and prepares for sync.
 * Returns the captured state so C# can pass it to syncToView after initialization.
 */
export function switchViewWithSync(newViewId) {
    // 1. Capture current state BEFORE stopping
    const state = captureCurrentMediaState();
    
    // 2. Soft-stop: pause without resetting positions
    softStopAllMedia();
    
    // 3. Now do full stop (resets positions, destroys article instances)
    // But we already captured the state above
    stopAllMedia();
    
    logInfo(`switchViewWithSync: captured state for ${newViewId}`, state);
    return state;
}

/**
 * Updates the view toggle slider position and size to match the currently active button.
 * Mirrors the static site's updatePillToggle approach: measures the active button's
 * offsetWidth and offsetLeft, then applies them to the slider div with CSS transition.
 * @param {boolean} [animate=true] - Whether to animate the transition. False on initial render.
 */
export function updateViewToggleSlider(animate = true) {
    const selectBar = document.getElementById('viewToggleSelectBar');
    const slider = document.getElementById('viewToggleSlider');
    if (!selectBar || !slider) {
        logInfo('updateViewToggleSlider: selectBar or slider not found');
        return;
    }
    
    const activeButton = selectBar.querySelector('.rz-state-active');
    if (!activeButton) {
        // Radzen may not have applied the active class yet — retry after a frame
        logInfo('updateViewToggleSlider: no active button found, retrying...');
        requestAnimationFrame(() => updateViewToggleSlider(animate));
        return;
    }
    
    // Measure the active button relative to the selectbar container
    const width = activeButton.offsetWidth;
    const left = activeButton.offsetLeft;
    
    if (width === 0) {
        // Button hasn't been laid out yet — retry after a frame
        logInfo('updateViewToggleSlider: active button has zero width, retrying...');
        requestAnimationFrame(() => updateViewToggleSlider(animate));
        return;
    }
    
    // On initial positioning, disable transition so slider appears instantly
    if (!animate) {
        slider.style.transition = 'none';
    }
    
    slider.style.width = width + 'px';
    slider.style.left = left + 'px';
    slider.style.opacity = '1';

    // Re-enable transition after the browser has painted
    if (!animate) {
        requestAnimationFrame(() => {
            slider.style.transition = '';
        });
    }

    watchViewToggleSelectBar(selectBar);

    logInfo(`updateViewToggleSlider: left=${left}px, width=${width}px, animate=${animate}`);
}

// The select bar the slider is being kept in step with, and the observer doing it.
let viewToggleSelectBar = null;
let viewToggleResizeObserver = null;

/**
 * Re-measures the slider whenever the view toggle changes size.
 *
 * The slider is drawn from one measurement of the active button, and the button is not a fixed
 * size: below the toolbar's label breakpoint (MainViewContentViewToolbar__Component's @media rule)
 * the labels are display:none and every button shrinks to its icon. Measured once, the blue stayed
 * label-wide behind an icon-wide button — and a page opened narrow and then widened had the
 * opposite, an icon-wide blue behind a label. A resize of the window does not move the buttons
 * unless it crosses that rule, so it is the select bar's own size that is watched, not the window.
 *
 * The guard is the element, not a flag: a Main View drawn again draws a new select bar, and an
 * observer left on the old one would follow a toolbar nobody can see.
 * @param {HTMLElement} selectBar - The #viewToggleSelectBar just measured.
 */
function watchViewToggleSelectBar(selectBar) {
    if (typeof ResizeObserver === 'undefined' || viewToggleSelectBar === selectBar) {
        return;
    }

    if (viewToggleResizeObserver) {
        viewToggleResizeObserver.disconnect();
    }

    viewToggleSelectBar = selectBar;

    // An observer reports the size it starts with; that one is the measurement just taken, and
    // re-taking it would cut short the animation of the switch that got us here.
    let first = true;

    viewToggleResizeObserver = new ResizeObserver(() => {
        if (first) {
            first = false;
            return;
        }

        if (!selectBar.isConnected) {
            return;
        }

        // Unanimated: the buttons have already changed size, and a slider sliding to catch up with
        // them reads as lag rather than as a change of view.
        updateViewToggleSlider(false);
    });

    viewToggleResizeObserver.observe(selectBar);
}

// ===== Consolidated Audio Player Functions =====

/**
 * Gets the consolidated audio player element.
 */
function getConsolidatedAudio() {
    return document.getElementById('consolidated-audio-player');
}

/**
 * Gets the consolidated video player element.
 */
function getConsolidatedVideo() {
    return document.getElementById('consolidated-video-player');
}

/**
 * Plays the consolidated audio player for a specific slide.
 * Loads the correct audio source if needed.
 */
export function playConsolidatedAudio(slideIndex) {
    logDebug(`playConsolidatedAudio: slide ${slideIndex}`);

    // Don't interfere while a decode error recovery is in-flight
    if (audioRecovering) {
        logDebug('playConsolidatedAudio: skipped (audio recovery in-flight)');
        return;
    }

    const audio = getConsolidatedAudio();
    if (!audio) {
        logWarning('playConsolidatedAudio: consolidated audio element not found');
        return;
    }

    const newSrc = getSlideMediaUrl(slideIndex, false);

    // Check if we need to change the source
    if (newSrc && !audio.src.endsWith(newSrc)) {
        audio.src = newSrc;
        audio.load();
        logDebug(`playConsolidatedAudio: loaded new source ${newSrc}`);
    }

    audio.play().catch(e => logWarning(`playConsolidatedAudio: play failed: ${e.message}`));
    setupConsolidatedAudioListeners(audio, slideIndex);
}

/**
 * Pauses the consolidated audio player.
 */
export function pauseConsolidatedAudio() {
    logDebug('pauseConsolidatedAudio');
    const audio = getConsolidatedAudio();
    if (audio) {
        audio.pause();
    }
}

/**
 * Seeks the consolidated audio player to a specific time for a given slide.
 */
export function seekConsolidatedAudio(slideIndex, time) {
    logDebug(`seekConsolidatedAudio: slide ${slideIndex}, time ${time.toFixed(2)}s`);
    const audio = getConsolidatedAudio();
    if (!audio) return;

    const newSrc = getSlideMediaUrl(slideIndex, false);

    if (newSrc && !audio.src.endsWith(newSrc)) {
        audio.src = newSrc;
        audio.load();
        audio.addEventListener('canplay', function onCanPlay() {
            audio.removeEventListener('canplay', onCanPlay);
            audio.currentTime = time;
        });
    } else {
        audio.currentTime = time;
    }
    setupConsolidatedAudioListeners(audio, slideIndex);
}

// The element the listeners were put on, not a flag: the players are Blazor's elements, and a Main View
// that is drawn again — another page, the Slide Site Preview after an edit, the designer's center —
// replaces them. A flag set once for the life of this module left every later element with no
// listeners at all, so nothing it played moved the notes, the animations or the next slide.
// Each holds { element, handlers } from bindMediaListeners, so the listeners can be taken off again.
let consolidatedAudioListenersAttached = null;
let consolidatedAudioSlideIndex = 0;

// Consolidated Video Player state
let consolidatedVideoListenersAttached = null;
let consolidatedVideoSlideIndex = 0;

// Tracks which player mode is active (true = video, false = audio)
let playerVideoMode = false;

// Decode error recovery state
const MAX_DECODE_RETRIES = 3;
const DECODE_RETRY_COOLDOWN_MS = 2000; // min time between retries to avoid tight loops
const DECODE_SKIP_FORWARD_S = 0.15; // seconds to skip forward past the bad frame on recovery
const RETRY_RESET_AFTER_MS = 5000; // reset retry counter after this much uninterrupted playback
let videoDecodeRetryCount = 0;
let audioDecodeRetryCount = 0;
let lastVideoRetryTime = 0;
let lastAudioRetryTime = 0;
let videoRecovering = false; // true while a video decode recovery is in-flight
let audioRecovering = false; // true while an audio decode recovery is in-flight
let videoRetryResetTimer = null; // timer to reset video retry count after sustained playback
let audioRetryResetTimer = null; // timer to reset audio retry count after sustained playback

// Slide media URL map: { [slideIndex]: { audio: "url", video: "url", png: "url" } }
// Populated from C# after slide data is resolved to absolute/blob URLs.
let slideMediaMap = {};
let isAutoAdvancing = false; // Suppresses pause handler during auto-advance source swap

/**
 * Sets the slide media URL map. Called from C# after slide data initialization
 * (and again after Blob URL creation for authenticated repos).
 * @param {Object} map - { [slideIndex]: { audio: "url", video: "url", png: "url" } }
 */
export function setSlideMediaMap(map) {
    slideMediaMap = map || {};
    logDebug(`setSlideMediaMap: ${Object.keys(slideMediaMap).length} slides`);
}

/**
 * Looks up the resolved media URL for a given slide index.
 * @param {number} slideIndex - 1-based slide index
 * @param {boolean} isVideo - true for video (.mp4), false for audio (.mp3)
 * @returns {string} The resolved URL, or empty string if not found
 */
function getSlideMediaUrl(slideIndex, isVideo) {
    const entry = slideMediaMap[slideIndex];
    if (!entry) return '';
    return isVideo ? (entry.video || '') : (entry.audio || '');
}

/**
 * Syncs the consolidated audio source to a given slide without playing.
 * Call this when navigating slides to keep the audio player in sync.
 */
export function syncConsolidatedAudioSource(slideIndex) {
    logDebug(`syncConsolidatedAudioSource: slide ${slideIndex}`);
    const audio = getConsolidatedAudio();
    if (!audio) return;

    const newSrc = getSlideMediaUrl(slideIndex, false);

    if (newSrc && !audio.src.endsWith(newSrc)) {
        audio.pause();
        audio.currentTime = 0;
        audio.src = newSrc;
        audio.load();
        consolidatedAudioSlideIndex = slideIndex;
        logDebug(`syncConsolidatedAudioSource: loaded ${newSrc}`);
    } else {
        consolidatedAudioSlideIndex = slideIndex;
    }

    // Listened to from the moment it has something to play, not from the first time this module plays
    // it: a recording started from the element's own controls — the Slide Site Designer's Player panel,
    // the Main View's Player panel — went through none of this module's play functions, so it played
    // with no line lit, no animation step and no move to the next slide.
    setupConsolidatedAudioListeners(audio, slideIndex);
}

// ===== Consolidated Video Player Functions =====

/**
 * Plays the consolidated video player for a specific slide.
 * Loads the correct video source if needed.
 */
export function playConsolidatedVideo(slideIndex) {
    logDebug(`playConsolidatedVideo: slide ${slideIndex}`);

    // Don't interfere while a decode error recovery is in-flight
    if (videoRecovering) {
        logDebug('playConsolidatedVideo: skipped (video recovery in-flight)');
        return;
    }

    const video = getConsolidatedVideo();
    if (!video) {
        logWarning('playConsolidatedVideo: consolidated video element not found');
        return;
    }

    const newSrc = getSlideMediaUrl(slideIndex, true);

    // Check if we need to change the source
    if (newSrc && !video.src.endsWith(newSrc)) {
        video.src = newSrc;
        video.load();
        logDebug(`playConsolidatedVideo: loaded new source ${newSrc}`);
    }

    video.play().catch(e => logWarning(`playConsolidatedVideo: play failed: ${e.message}`));
    setupConsolidatedVideoListeners(video, slideIndex);
}

/**
 * Pauses the consolidated video player.
 */
export function pauseConsolidatedVideo() {
    logDebug('pauseConsolidatedVideo');
    const video = getConsolidatedVideo();
    if (video) {
        video.pause();
    }
}

/**
 * Seeks the consolidated video player to a specific time for a given slide.
 */
export function seekConsolidatedVideo(slideIndex, time) {
    logDebug(`seekConsolidatedVideo: slide ${slideIndex}, time ${time.toFixed(2)}s`);
    const video = getConsolidatedVideo();
    if (!video) return;

    const newSrc = getSlideMediaUrl(slideIndex, true);

    if (newSrc && !video.src.endsWith(newSrc)) {
        video.src = newSrc;
        video.load();
        video.addEventListener('canplay', function onCanPlay() {
            video.removeEventListener('canplay', onCanPlay);
            video.currentTime = time;
        });
    } else {
        video.currentTime = time;
    }
    setupConsolidatedVideoListeners(video, slideIndex);
}

/**
 * Syncs the consolidated video source to a given slide without playing.
 * Call this when navigating slides to keep the video player in sync.
 */
export function syncConsolidatedVideoSource(slideIndex) {
    logDebug(`syncConsolidatedVideoSource: slide ${slideIndex}`);
    const video = getConsolidatedVideo();
    if (!video) return;

    const newSrc = getSlideMediaUrl(slideIndex, true);

    if (newSrc && !video.src.endsWith(newSrc)) {
        video.pause();
        video.currentTime = 0;
        video.src = newSrc;
        video.load();
        consolidatedVideoSlideIndex = slideIndex;
        logDebug(`syncConsolidatedVideoSource: loaded ${newSrc}`);
    } else {
        consolidatedVideoSlideIndex = slideIndex;
    }

    // Listened to from the moment it has something to play — see syncConsolidatedAudioSource.
    setupConsolidatedVideoListeners(video, slideIndex);
}

/**
 * Sets the player video mode flag without reloading sources or seeking.
 * Used by SwitchView to sync the JS-side flag before syncToView handles source loading.
 * @param {boolean} toVideo - True for video mode, false for audio mode.
 */
export function setPlayerVideoMode(toVideo) {
    playerVideoMode = toVideo;
    logDebug(`setPlayerVideoMode: ${toVideo ? 'video' : 'audio'}`);
}

/**
 * Switches between audio and video player modes, preserving the current timestamp.
 * Called from C# when the user toggles the video/audio switch.
 * @param {boolean} toVideo - True to switch to video mode, false for audio mode.
 */
export function switchPlayerMode(toVideo) {
    // Suppress note updates during the mode toggle so stale timeupdate events don't cause flashing
    noteUpdateSuppressedUntil = Date.now() + NOTE_SUPPRESS_TRANSITION_MS;

    const fromElement = toVideo ? getConsolidatedAudio() : getConsolidatedVideo();
    const toElement = toVideo ? getConsolidatedVideo() : getConsolidatedAudio();
    const fromSlideIndex = toVideo ? consolidatedAudioSlideIndex : consolidatedVideoSlideIndex;

    let currentTime = 0;
    let wasPlaying = false;

    // Capture state from the currently active player
    if (fromElement) {
        currentTime = fromElement.currentTime || 0;
        wasPlaying = !fromElement.paused;
        fromElement.pause();
        logDebug(`switchPlayerMode: captured time=${currentTime.toFixed(2)}, playing=${wasPlaying} from ${toVideo ? 'audio' : 'video'}`);
    }

    playerVideoMode = toVideo;

    // Transfer state to the new player
    if (toElement) {
        // Ensure the correct source is loaded
        const slideIndex = fromSlideIndex || 1;
        if (toVideo) {
            const newSrc = getSlideMediaUrl(slideIndex, true);
            if (newSrc && !toElement.src.endsWith(newSrc)) {
                toElement.src = newSrc;
                toElement.load();
            }
            consolidatedVideoSlideIndex = slideIndex;
            setupConsolidatedVideoListeners(toElement, slideIndex);
        } else {
            const newSrc = getSlideMediaUrl(slideIndex, false);
            if (newSrc && !toElement.src.endsWith(newSrc)) {
                toElement.src = newSrc;
                toElement.load();
            }
            consolidatedAudioSlideIndex = slideIndex;
            setupConsolidatedAudioListeners(toElement, slideIndex);
        }

        // Wait for the media to be ready, then set time and optionally play.
        // Clear note suppression only after the seek completes (seeked event)
        // to prevent stale timestamps from leaking through before the new
        // player reaches the target position.
        const applyState = () => {
            if (currentTime > 0) {
                toElement.currentTime = currentTime;
                toElement.addEventListener('seeked', function onSeeked() {
                    toElement.removeEventListener('seeked', onSeeked);
                    noteUpdateSuppressedUntil = 0;
                    if (wasPlaying) {
                        toElement.play().catch(e => logWarning(`switchPlayerMode play failed: ${e.message}`));
                    }
                });
            } else {
                noteUpdateSuppressedUntil = 0;
                if (wasPlaying) {
                    toElement.play().catch(e => logWarning(`switchPlayerMode play failed: ${e.message}`));
                }
            }
        };

        if (toElement.readyState >= 2) {
            applyState();
        } else {
            toElement.addEventListener('canplay', function onCanPlay() {
                toElement.removeEventListener('canplay', onCanPlay);
                applyState();
            });
        }
    }

    logInfo(`switchPlayerMode: switched to ${toVideo ? 'video' : 'audio'} mode`);
}

/**
 * Shared timeupdate handler for both consolidated audio and video players.
 * Updates progress bars, timestamps, note highlighting, and fragment navigation.
 */
function handleConsolidatedMediaTimeUpdate(mediaElement, slideIndex) {
    // Only process timeupdate from actively playing media to prevent flashing highlights
    if (mediaElement.paused) return;

    const currentTime = mediaElement.currentTime;
    const idx = slideIndex;

    // Calculate absolute video timestamp for global progress
    const notesEl = document.getElementById(`slideNotes-${idx}`);
    let absoluteTime = currentTime;
    if (notesEl) {
        const firstLine = notesEl.querySelector('[data-start-video]');
        if (firstLine) {
            const slideStartVideo = parseFloat(firstLine.getAttribute('data-start-video')) || 0;
            absoluteTime = slideStartVideo + currentTime;
        }
    }

    // Always update progress bars directly (no throttle for visual smoothness)
    updateProgressBarDirect(absoluteTime);

    // Update slide-level progress bar directly in DOM
    const slideDuration = mediaElement.duration || 0;
    if (slideDuration > 0) {
        const slidePercent = (currentTime / slideDuration) * 100;
        const slideFill = document.getElementById(`slideProgressFill-${idx}`);
        if (slideFill) slideFill.style.width = `${slidePercent}%`;
    }

    // Update slide timestamp directly in DOM
    const slideTimestampEl = document.getElementById(`slideTimestamp-${idx}`);
    if (slideTimestampEl && slideDuration > 0) {
        const formatTs = (s) => {
            if (isNaN(s) || s < 0) return '0:00';
            const m = Math.floor(s / 60);
            const sec = Math.floor(s % 60);
            return `${m}:${String(sec).padStart(2, '0')}`;
        };
        slideTimestampEl.textContent = `${formatTs(currentTime)} / ${formatTs(slideDuration)}`;
    }

    // Update bottom toolbar timestamp directly in DOM
    const timestampDisplayEl = document.getElementById('timestampDisplay');
    if (timestampDisplayEl && cachedTotalDuration > 0) {
        const formatTs = (s) => {
            if (isNaN(s) || s < 0) return '0:00';
            const h = Math.floor(s / 3600);
            const m = Math.floor((s % 3600) / 60);
            const sec = Math.floor(s % 60);
            return h > 0
                ? `${h}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`
                : `${m}:${String(sec).padStart(2, '0')}`;
        };
        // Find the inner span/text node - RadzenText renders text inside
        const innerSpan = timestampDisplayEl.querySelector('.rz-text') || timestampDisplayEl;
        innerSpan.textContent = `${formatTs(absoluteTime)} / ${formatTs(cachedTotalDuration)}`;
    }

    // Find the matching note line for timestamp highlight (throttled for Blazor calls)
    const slideCard = document.getElementById(`slide-${idx}`);
    if (slideCard) {
        const lines = slideCard.querySelectorAll('[data-start][data-end]');
        let matchedLineIndex = -1;
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const start = parseFloat(line.getAttribute('data-start'));
            const end = parseFloat(line.getAttribute('data-end'));
            const noteId = line.getAttribute('data-note-id') || line.id;
            const noteSlideIndex = parseInt(line.getAttribute('data-slide-index') || idx);
            if (currentTime >= start && currentTime < end && noteId) {
                matchedLineIndex = i;
                const videoStart = parseFloat(line.getAttribute('data-start-video') || '0');
                const timestamp = videoStart + (currentTime - start);

                invokeCurrentNoteUpdate(noteId, noteSlideIndex, timestamp);

                const now = Date.now();
                if (now - lastTimestampUpdate >= TIMESTAMP_UPDATE_INTERVAL) {
                    lastTimestampUpdate = now;
                    if (dotNetRef) {
                        dotNetRef.invokeMethodAsync('UpdateTimestamp', timestamp);
                    }
                }
                break;
            }
        }

        // Navigate to matching fragment in article view
        const slideData = articleRevealInstances.find(s => s.index === idx);
        if (slideData && slideData.instance && matchedLineIndex >= 0 && hoverFragmentSlide !== idx) {
            applyArticleFragment(slideData, idx, spokenLineIndex(lines, matchedLineIndex));
        }

        // Navigate to matching fragment in presentation view
        if (presentationReveal && matchedLineIndex >= 0) {
            applyPresentationFragment(spokenLineIndex(lines, matchedLineIndex));
        }
    }
}

/**
 * Shared play handler for both consolidated audio and video players.
 */
function handleConsolidatedMediaPlay(slideIndex, mediaLabel) {
    logDebug(`Consolidated ${mediaLabel} play: slide ${slideIndex}`);

    // Enable fragments on the current slide (article view)
    const revealEl = document.getElementById(`reveal-${slideIndex}`);
    if (revealEl) {
        const fragmentsOff = revealEl.querySelectorAll('.fragment-off');
        logDebug(`Enabling ${fragmentsOff.length} fragments for slide ${slideIndex}`);
        fragmentsOff.forEach(f => {
            f.classList.remove('fragment-off');
            f.classList.remove('visible');
            f.classList.add('fragment');
        });
    }

    // Enable fragments in presentation view
    if (presentationReveal) {
        enablePresentationFragmentsForSlide();
    }

    if (dotNetRef) {
        dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
        dotNetRef.invokeMethodAsync('ExpandSlideNavCard', slideIndex);
        dotNetRef.invokeMethodAsync('ShowArticleSlideFragmentControls', slideIndex);
        dotNetRef.invokeMethodAsync('ShowPresentationFragmentControls');
    }
}

/**
 * Shared pause handler for both consolidated audio and video players.
 */
function handleConsolidatedMediaPause(mediaLabel) {
    if (isAutoAdvancing) return;
    logDebug(`Consolidated ${mediaLabel} pause`);
    if (dotNetRef) {
        dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
    }
}

/**
 * Shared ended handler for both consolidated audio and video players.
 * Auto-advances to the next slide and loads the appropriate media file.
 * @param {HTMLMediaElement} mediaElement - The audio or video element
 * @param {function} getSlideIndex - Function returning the current slide index
 * @param {function} setSlideIndex - Function to set the new slide index
 * @param {boolean} isVideo - True for video, false for audio
 */
function handleConsolidatedMediaEnded(mediaElement, getSlideIndex, setSlideIndex, isVideo) {
    const mediaLabel = isVideo ? 'video' : 'audio';
    const currentIdx = getSlideIndex();
    logDebug(`Consolidated ${mediaLabel} ended: slide ${currentIdx}`);
    const nextIdx = currentIdx + 1;

    // Check if there's a next slide
    const nextSlideEl = document.getElementById(`slide-${nextIdx}`);
    const nextNotesEl = document.getElementById(`slideNotes-${nextIdx}`);

    if (nextSlideEl || nextNotesEl) {
        logDebug(`Auto-advancing to slide ${nextIdx}`);
        setSlideIndex(nextIdx);

        // Scroll to next slide in article view
        if (nextSlideEl) {
            nextSlideEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        // Navigate presentation view to next slide
        if (presentationReveal) {
            presentationReveal.slide(nextIdx - 1);
        }

        // Load and play the next slide's media
        // Load and play the next slide's media (fall back to other type if needed)
        let newSrc = getSlideMediaUrl(nextIdx, isVideo);
        let useVideo = isVideo;
        let targetElement = mediaElement;

        if (!newSrc) {
            // Try the other media type
            const altSrc = getSlideMediaUrl(nextIdx, !isVideo);
            if (altSrc) {
                useVideo = !isVideo;
                newSrc = altSrc;
                targetElement = useVideo ? getConsolidatedVideo() : getConsolidatedAudio();
                playerVideoMode = useVideo;
                logDebug(`Auto-advance: switching to ${useVideo ? 'video' : 'audio'} for slide ${nextIdx}`);
                if (dotNetRef) {
                    dotNetRef.invokeMethodAsync('UpdatePlayerVideoMode', useVideo);
                }
                if (useVideo) {
                    consolidatedVideoSlideIndex = nextIdx;
                    if (targetElement) setupConsolidatedVideoListeners(targetElement, nextIdx);
                } else {
                    consolidatedAudioSlideIndex = nextIdx;
                    if (targetElement) setupConsolidatedAudioListeners(targetElement, nextIdx);
                }
            }
        }

        if (newSrc && targetElement) {
            isAutoAdvancing = true;
            const playTarget = targetElement;
            const tryPlay = () => {
                if (playTarget.readyState >= 2) {
                    isAutoAdvancing = false;
                    playTarget.play().catch(e => logWarning(`Auto-advance ${useVideo ? 'video' : 'audio'} play failed: ${e.message}`));
                } else {
                    playTarget.addEventListener('canplay', function onCanPlay() {
                        playTarget.removeEventListener('canplay', onCanPlay);
                        isAutoAdvancing = false;
                        playTarget.play().catch(e => logWarning(`Auto-advance ${useVideo ? 'video' : 'audio'} play failed: ${e.message}`));
                    });
                }
            };
            playTarget.src = newSrc;
            playTarget.load();
            tryPlay();
        } else {
            logDebug(`Auto-advance: no media URL for slide ${nextIdx}, skipping media load`);
        }

        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('ExpandSlideNavCard', nextIdx);
            dotNetRef.invokeMethodAsync('UpdatePlayingState', true);
        }
    } else {
        logDebug('No more slides to advance to');
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
        }
    }
}

/**
 * Sets up timeupdate listeners on the consolidated audio element.
 * Handles: timestamp/progress updates, fragment navigation, auto-advance, and play/pause state.
 */
function setupConsolidatedAudioListeners(audio, slideIndex) {
    consolidatedAudioSlideIndex = slideIndex;
    if (consolidatedAudioListenersAttached && consolidatedAudioListenersAttached.element === audio) return;
    unbindMediaListeners(consolidatedAudioListenersAttached);

    consolidatedAudioListenersAttached = bindMediaListeners(audio, {
        ...mediaSettingHandlers,
        timeupdate: () => {
            handleConsolidatedMediaTimeUpdate(audio, consolidatedAudioSlideIndex);
            // Schedule a delayed reset of the retry counter after sustained playback.
            if (audioDecodeRetryCount > 0 && !audioRecovering) {
                if (audioRetryResetTimer) clearTimeout(audioRetryResetTimer);
                audioRetryResetTimer = setTimeout(() => {
                    audioDecodeRetryCount = 0;
                    audioRetryResetTimer = null;
                    logDebug('Audio decode retry counter reset after sustained playback');
                }, RETRY_RESET_AFTER_MS);
            }
        },
        play: () => handleConsolidatedMediaPlay(consolidatedAudioSlideIndex, 'audio'),
        pause: () => handleConsolidatedMediaPause('audio'),
        ended: () => handleConsolidatedMediaEnded(
            audio,
            () => consolidatedAudioSlideIndex,
            (idx) => { consolidatedAudioSlideIndex = idx; },
            false
        ),
        error: () => handleMediaDecodeError(audio, consolidatedAudioSlideIndex, false)
    });

    applyMediaSettings(audio);
}

/**
 * Puts listeners on a media element and returns what is needed to take them off again.
 * @param {HTMLMediaElement} element
 * @param {Object<string, function>} handlers - { [eventName]: handler }
 * @returns {{ element: HTMLMediaElement, handlers: Object<string, function> }}
 */
function bindMediaListeners(element, handlers) {
    Object.entries(handlers).forEach(([name, handler]) => element.addEventListener(name, handler));
    return { element, handlers };
}

/**
 * Takes off the listeners bindMediaListeners put on, or does nothing for null.
 *
 * Needed because an element can outlive the Main View that listened to it: the Slide Site Designer's
 * Player panel holds the two consolidated players outside its embedded Main View, so when that view is
 * drawn again the same element was listened to twice — and on `ended` the first handler moved to the
 * next slide and the second, reading the index the first had just moved, moved on past it.
 */
function unbindMediaListeners(binding) {
    if (!binding) return;
    Object.entries(binding.handlers).forEach(([name, handler]) => binding.element.removeEventListener(name, handler));
}

/**
 * Sets up timeupdate listeners on the consolidated video element.
 * Mirrors setupConsolidatedAudioListeners but for the video player.
 */
function setupConsolidatedVideoListeners(video, slideIndex) {
    consolidatedVideoSlideIndex = slideIndex;
    if (consolidatedVideoListenersAttached && consolidatedVideoListenersAttached.element === video) return;
    unbindMediaListeners(consolidatedVideoListenersAttached);

    consolidatedVideoListenersAttached = bindMediaListeners(video, {
        ...mediaSettingHandlers,
        timeupdate: () => {
            handleConsolidatedMediaTimeUpdate(video, consolidatedVideoSlideIndex);
            // Schedule a delayed reset of the retry counter after sustained playback.
            // This avoids immediately resetting between rapid error/recovery cycles.
            if (videoDecodeRetryCount > 0 && !videoRecovering) {
                if (videoRetryResetTimer) clearTimeout(videoRetryResetTimer);
                videoRetryResetTimer = setTimeout(() => {
                    videoDecodeRetryCount = 0;
                    videoRetryResetTimer = null;
                    logDebug('Video decode retry counter reset after sustained playback');
                }, RETRY_RESET_AFTER_MS);
            }
        },
        play: () => handleConsolidatedMediaPlay(consolidatedVideoSlideIndex, 'video'),
        pause: () => handleConsolidatedMediaPause('video'),
        ended: () => handleConsolidatedMediaEnded(
            video,
            () => consolidatedVideoSlideIndex,
            (idx) => { consolidatedVideoSlideIndex = idx; },
            true
        ),
        error: () => handleMediaDecodeError(video, consolidatedVideoSlideIndex, true)
    });

    applyMediaSettings(video);
}


/**
 * Handles media decode errors (MEDIA_ERR_DECODE, code 3) by re-loading the
 * source and resuming playback from the position where the error occurred.
 * Uses a retry counter and cooldown to avoid infinite retry loops.
 * @param {HTMLMediaElement} mediaElement - The audio or video element that errored
 * @param {number} slideIndex - The current slide index
 * @param {boolean} isVideo - True for video, false for audio
 */
function handleMediaDecodeError(mediaElement, slideIndex, isVideo) {
    const mediaLabel = isVideo ? 'video' : 'audio';
    const mediaError = mediaElement.error;

    if (!mediaError) return;

    // Only attempt recovery for decode errors (code 3)
    if (mediaError.code !== 3) {
        logError(`Consolidated ${mediaLabel} error (non-decode, code ${mediaError.code}): ${mediaError.message}`);
        return;
    }

    // Ignore error events that fire while we are already recovering (re-load can trigger
    // a duplicate error event before the new source finishes loading)
    const isRecovering = isVideo ? videoRecovering : audioRecovering;
    if (isRecovering) {
        logDebug(`Consolidated ${mediaLabel} decode error: ignored (recovery already in-flight)`);
        return;
    }

    const retryCount = isVideo ? videoDecodeRetryCount : audioDecodeRetryCount;
    const lastRetryTime = isVideo ? lastVideoRetryTime : lastAudioRetryTime;
    const now = Date.now();

    // Enforce cooldown between retries
    if (now - lastRetryTime < DECODE_RETRY_COOLDOWN_MS) {
        logWarning(`Consolidated ${mediaLabel} decode error: skipping retry (cooldown active, ${now - lastRetryTime}ms since last retry)`);
        return;
    }

    if (retryCount >= MAX_DECODE_RETRIES) {
        logError(`Consolidated ${mediaLabel} decode error: giving up after ${MAX_DECODE_RETRIES} retries. ${mediaError.message}`);
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
        }
        return;
    }

    // Mark recovery in-flight to suppress cascading error events and external play() calls
    if (isVideo) {
        videoRecovering = true;
        videoDecodeRetryCount = retryCount + 1;
        lastVideoRetryTime = now;
        if (videoRetryResetTimer) { clearTimeout(videoRetryResetTimer); videoRetryResetTimer = null; }
    } else {
        audioRecovering = true;
        audioDecodeRetryCount = retryCount + 1;
        lastAudioRetryTime = now;
        if (audioRetryResetTimer) { clearTimeout(audioRetryResetTimer); audioRetryResetTimer = null; }
    }

    // Capture position before error (currentTime may become 0 after error on some browsers)
    const savedTime = mediaElement.currentTime || 0;
    const wasPaused = mediaElement.paused;
    // Skip forward slightly to avoid re-hitting the exact same bad frame/packet
    const resumeTime = savedTime + DECODE_SKIP_FORWARD_S;
    logWarning(`Consolidated ${mediaLabel} decode error on slide ${slideIndex} at ${savedTime.toFixed(2)}s ` +
        `(retry ${retryCount + 1}/${MAX_DECODE_RETRIES}): ${mediaError.message}`);

    // Re-load the same source to reset the decoder pipeline
    const src = getSlideMediaUrl(slideIndex, isVideo);
    if (!src) return;

    if (isVideo) {
        mediaElement.src = src;
    } else {
        const source = mediaElement.querySelector('source');
        if (source) {
            source.src = src;
        }
        mediaElement.load();
    }

    // Once re-loaded, seek past the bad frame and resume if it was playing
    const onReady = () => {
        mediaElement.removeEventListener('canplay', onReady);
        logInfo(`Consolidated ${mediaLabel} decode error recovery: seeking to ${resumeTime.toFixed(2)}s (skipped +${DECODE_SKIP_FORWARD_S}s)`);
        mediaElement.currentTime = resumeTime;
        if (!wasPaused) {
            mediaElement.play().then(() => {
                logDebug(`Decode error recovery play succeeded`);
            }).catch(e => {
                logWarning(`Decode error recovery play failed: ${e.message}`);
                if (dotNetRef) {
                    dotNetRef.invokeMethodAsync('UpdatePlayingState', false);
                }
            }).finally(() => {
                if (isVideo) videoRecovering = false;
                else audioRecovering = false;
            });
        } else {
            if (isVideo) videoRecovering = false;
            else audioRecovering = false;
        }
    };

    // If canplay doesn't fire within 5s (e.g. source also errors on re-load), clear the flag
    const recoveryTimeout = setTimeout(() => {
        mediaElement.removeEventListener('canplay', onReady);
        logWarning(`Consolidated ${mediaLabel} decode error recovery timed out`);
        if (isVideo) videoRecovering = false;
        else audioRecovering = false;
    }, 5000);

    mediaElement.addEventListener('canplay', function onCanPlayWrapper() {
        mediaElement.removeEventListener('canplay', onCanPlayWrapper);
        clearTimeout(recoveryTimeout);
        onReady();
    });
}

// ===== Volume and playback speed =====
//
// One volume and one speed for the page, held here and put on every player the page plays through — the two
// consolidated players, the film and YouTube — whenever one is set and whenever a player is first listened to.
// The toolbar used to set them on the two consolidated players only, once (the user's report, 2026-10-02:
// "not truly reflecting the state"):
//  - every new source (`load()`, which the move to the next slide does) puts `playbackRate` back to
//    `defaultPlaybackRate`, so 1.5x went back to 1x at the next slide while the toolbar still said 1.5x;
//    both are set now;
//  - the film never had either, so the Video view ignored the toolbar;
//  - a player's own controls (the Designer's Player panel, the Main View's Player panel) changed the element
//    and the toolbar never heard; `volumechange` and `ratechange` now report it back (mediaSettingsChanged).

let narrationVolume = 1;
let narrationRate = 1;

/** The volume a player is actually at — a muted player is at nought whatever its volume says. */
function effectiveVolume(element) {
    return element.muted ? 0 : element.volume;
}

/** Every element the page's volume and speed belong on. */
function mediaElementsForSettings() {
    return [getConsolidatedAudio(), getConsolidatedVideo(), getFilmPlayer()].filter(Boolean);
}

/** Puts the page's volume and speed on one player. Setting what it already has fires no event. */
function applyMediaSettings(element) {
    if (!element) return;

    if (element.muted && narrationVolume > 0) element.muted = false;
    if (element.volume !== narrationVolume) element.volume = narrationVolume;
    if (element.defaultPlaybackRate !== narrationRate) element.defaultPlaybackRate = narrationRate;
    if (element.playbackRate !== narrationRate) element.playbackRate = narrationRate;
}

function applyYouTubeSettings() {
    if (!youtubePlayer || !youtubePlayerReady) return;

    try {
        youtubePlayer.setVolume(narrationVolume * 100);
        youtubePlayer.setPlaybackRate(narrationRate);
    } catch (e) {
        logWarning(`YouTube volume/speed error: ${e.message}`);
    }
}

function applyMediaSettingsEverywhere() {
    mediaElementsForSettings().forEach(applyMediaSettings);
    applyYouTubeSettings();
}

/**
 * A player's volume or speed changed. When it is not what the page holds — the user moved the player's own
 * control — the page takes it, the other players follow, and .NET is told so the toolbar says it. A change
 * this module made itself matches what the page holds and is ignored.
 */
function onMediaSettingChanged(event) {
    const element = event.target;
    const volume = effectiveVolume(element);
    const rate = element.playbackRate;

    if (volume === narrationVolume && rate === narrationRate) return;

    narrationVolume = volume;
    narrationRate = rate;

    mediaElementsForSettings().filter(other => other !== element).forEach(applyMediaSettings);
    applyYouTubeSettings();

    if (dotNetRef) dotNetRef.invokeMethodAsync('UpdateMediaSettings', volume, rate);
}

/** The two listeners that report a player's own volume and speed changes — see onMediaSettingChanged. */
const mediaSettingHandlers = {
    volumechange: onMediaSettingChanged,
    ratechange: onMediaSettingChanged
};

/**
 * Sets the playback rate on every player the page plays through.
 */
export function setPlaybackRate(rate) {
    logDebug(`setPlaybackRate: ${rate}`);
    narrationRate = rate;
    applyMediaSettingsEverywhere();
}

/**
 * Sets the volume on every player the page plays through.
 */
export function setVolume(volume) {
    logDebug(`setVolume: ${volume}`);
    narrationVolume = volume;
    applyMediaSettingsEverywhere();
}

/**
 * Forces the consolidated video player to load its media source.
 * Browsers do not automatically load <video> elements that are dynamically
 * inserted into the DOM by frameworks like Blazor. Calling .load() triggers
 * the browser's media resource selection algorithm.
 */
export function loadVideoPlayer() {
    if (_videoPlayerRetryTimer) {
        clearInterval(_videoPlayerRetryTimer);
        _videoPlayerRetryTimer = null;
    }

    const video = document.getElementById('consolidated-video-player');
    if (!video) {
        // Element may not be in the DOM yet. Schedule a non-blocking
        // background retry instead of blocking the C# caller.
        let attempt = 0;
        _videoPlayerRetryTimer = setInterval(() => {
            attempt++;
            const el = document.getElementById('consolidated-video-player');
            if (el) {
                clearInterval(_videoPlayerRetryTimer);
                _videoPlayerRetryTimer = null;
                logInfo(`Video player element appeared on retry #${attempt}`);
                loadVideoPlayer();
            } else if (attempt >= 40) {
                clearInterval(_videoPlayerRetryTimer);
                _videoPlayerRetryTimer = null;
                logWarning('Video player element not found after retries');
            }
        }, 250);
        return;
    }

    const src = video.getAttribute('src');
    console.info(`[ArticleViewer] loadVideoPlayer: element found, src attribute="${src}", video.src="${video.src}"`);
    console.info(`[ArticleViewer] loadVideoPlayer: networkState=${video.networkState}, readyState=${video.readyState}`);

    // Listen for media events to diagnose loading issues
    video.addEventListener('loadstart', () => console.info('[ArticleViewer] video event: loadstart'), { once: true });
    video.addEventListener('loadedmetadata', () => console.info('[ArticleViewer] video event: loadedmetadata'), { once: true });
    video.addEventListener('loadeddata', () => console.info('[ArticleViewer] video event: loadeddata'), { once: true });
    video.addEventListener('canplay', () => console.info('[ArticleViewer] video event: canplay'), { once: true });
    video.addEventListener('error', (e) => {
        const mediaError = video.error;
        // Log for diagnostics; actual recovery is handled by setupConsolidatedVideoListeners
        logWarning(`loadVideoPlayer: video error event`, {
            code: mediaError?.code,
            message: mediaError?.message
        });
    }, { once: true });
    video.addEventListener('stalled', () => console.warn('[ArticleViewer] video event: stalled'), { once: true });
    video.addEventListener('suspend', () => console.info('[ArticleViewer] video event: suspend'), { once: true });

    // Force reload: re-assign src and call load()
    if (src) {
        video.src = src;
    }
    video.load();
    console.info('[ArticleViewer] loadVideoPlayer: load() called');

    // Set up consolidated video listeners so playback controls work
    const slideIndex = consolidatedVideoSlideIndex || 1;
    setupConsolidatedVideoListeners(video, slideIndex);
}

/**
 * Creates a Blob URL from a byte array for efficient media playback.
 * Used for video files in authenticated repos to avoid base64 data URI overhead.
 * The caller must revoke the URL via revokeBlobUrl when no longer needed.
 */
export function createBlobUrl(byteArray, mimeType) {
    const blob = new Blob([new Uint8Array(byteArray)], { type: mimeType });
    return URL.createObjectURL(blob);
}

/**
 * Revokes a Blob URL previously created by createBlobUrl.
 */
export function revokeBlobUrl(url) {
    if (url && url.startsWith('blob:')) {
        URL.revokeObjectURL(url);
    }
}