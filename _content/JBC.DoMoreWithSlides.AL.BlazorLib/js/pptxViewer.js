// The Article (PPTX) and Slides (PPTX) views: Blazor Office's own slide show player, embedded.
//
// Each slide is its Generated (PPTX) drawing — Blazor Office's SVG of the slide, every shape tagged for the
// player — and the model beside it: the slide's transition, its animation effects and the steps they play in
// (see SlidePptx_Helper). The player is slide-show.js from Blazor Office (BlazorOffice.DL.JSInterop), at the
// address Blazor Office serves it from — a head that registers Blazor Office has it there, and the GitSlides
// viewer and the _radzen heads link the one file to the same address in their .csproj — loaded the first time a
// PPTX view is shown. It plays the deck exactly as Blazor Office's own slide show does, through its `embed`
// entry point, with the host deciding when each click comes.
//
// The two engines answer the calls articleViewer.js makes of a Reveal.js instance — slide(h, v, f), getState(),
// getCurrentSlide(), nextFragment(), prevFragment(), on(), layout(), destroy() — so the narration that steps a
// slide's fragments, the step counter and the slide buttons drive either engine the same way. A "fragment" is a
// click of the slide's main sequence; fragment -1 is the slide as it stands once it has arrived (its automatic
// step played). The section a slide is drawn in carries data-pptx-clicks, which is what the step counter counts.
//
// Every slide starts fully built, as the Reveal views' slides do while their fragments are "off" (the user's
// rule, 2026-10-16: every element shows while nothing plays and no line is hovered). articleViewer.js switches
// an engine's steps on at the moments it takes a Reveal slide's `fragment-off` away — the narration playing
// it, a line hovered or clicked, the step controls opened — and off again with the controls.

const PLAYER_PATH = '_content/JBC.DoMoreWithOffice.BlazorOffice.DL.JSInterop/js/slide-show.js';

let playerPromise = null;

/** Blazor Office's player module, loaded once; null when it cannot be had. */
function loadPlayer() {
    if (!playerPromise) {
        // Against the page's base, as Blazor Office's own interop loads it — not against this file.
        playerPromise = import(new URL(PLAYER_PATH, document.baseURI).href).catch(error => {
            console.warn('[pptxViewer] Blazor Office\'s slide show player could not be loaded', error);
            playerPromise = null;
            return null;
        });
    }
    return playerPromise;
}

/** True when the player can be loaded — what the views ask before drawing anything. */
export async function isAvailable() {
    return !!(await loadPlayer());
}

// ---- The slide's row --------------------------------------------------------------------------------------

/** The player's row for a slide — the model with the drawing on it — or null when the model will not parse. */
function rowOf(slide) {
    let model = {};
    try {
        model = typeof slide.model === 'string' && slide.model.length > 0 ? JSON.parse(slide.model) : (slide.model || {});
    } catch (error) {
        console.warn(`[pptxViewer] slide ${slide.index}'s model could not be read`, error);
        model = {};
    }
    return {
        ...model,
        svg: slide.svg || '',
        effects: model.effects || [],
        steps: model.steps || [],
        media: (model.media || []).filter(item => item && (item.url || item.isNarration)),
        morphPairs: model.morphPairs || [],
    };
}

/** How many clicks a row's main sequence has — not the step that plays by itself on arrival. */
function clicksOf(row) {
    return (row?.steps || []).filter(step => step.triggerShapeId == null && !step.isAutomatic).length;
}

/** The size a row's slide is drawn at, from its model, else the host's. */
function sizeOf(row, slide) {
    return {
        width: row.width || slide.width || 1280,
        height: row.height || slide.height || 720,
    };
}

// ---- What both engines share --------------------------------------------------------------------------------

class PptxEngine {
    constructor() {
        this.isPptx = true;
        this.listeners = new Map();
        this.screen = null;
        // Off until something asks for steps: the slide is drawn fully built.
        this.fragmentsEnabled = false;
        this.disposed = false;
    }

    on(name, callback) {
        if (!this.listeners.has(name)) this.listeners.set(name, new Set());
        this.listeners.get(name).add(callback);
    }

    off(name, callback) {
        this.listeners.get(name)?.delete(callback);
    }

    emit(name, detail) {
        for (const callback of this.listeners.get(name) || []) {
            try {
                callback(detail);
            } catch (error) {
                console.warn(`[pptxViewer] a '${name}' listener failed`, error);
            }
        }
    }

    get clickCount() { return this.screen?.clickCount ?? 0; }

    get clicksShown() { return this.screen?.clicksShown ?? 0; }

    /**
     * Reveal's fragment index of what is on screen: -1 before the first click — and with steps off, as a Reveal
     * slide whose fragments are off has none to be on.
     */
    get fragmentIndex() { return this.fragmentsEnabled ? this.clicksShown - 1 : -1; }

    /**
     * With steps off, finishes at once whatever the arrival set moving: a slide with no clicks is handed 0 and
     * plays its automatic step, which would animate the slide in rather than show it built.
     */
    settle() {
        if (!this.fragmentsEnabled && this.screen && !this.disposed) this.screen.finish();
    }

    /** Switches the steps on when they are off — what playing or hovering a line does — and says whether it did. */
    wake() {
        if (this.fragmentsEnabled || !this.screen || this.disposed) return false;
        this.setFragmentsEnabled(true);
        return true;
    }

    /**
     * Shows the slide at fragment `f` — `f + 1` clicks played — animating when it is the next click along. With
     * fragments switched off the slide is always shown fully built.
     */
    showFragment(f, animate = true) {
        if (!this.screen || this.disposed) return;
        const before = this.clicksShown;
        const wanted = this.fragmentsEnabled ? (typeof f === 'number' ? f : -1) + 1 : this.clickCount;
        this.screen.showClicks(Math.max(0, Math.min(wanted, this.clickCount)), animate);
        const after = this.clicksShown;
        if (after > before) this.emit('fragmentshown', { fragmentIndex: after - 1 });
        else if (after < before) this.emit('fragmenthidden', { fragmentIndex: after - 1 });
    }

    nextFragment() {
        if (!this.screen || !this.fragmentsEnabled || this.clicksShown >= this.clickCount) return false;
        this.showFragment(this.clicksShown, true);
        return true;
    }

    prevFragment() {
        if (!this.screen || !this.fragmentsEnabled || this.clicksShown <= 0) return false;
        this.showFragment(this.clicksShown - 2, false);
        return true;
    }

    /** Fragments off draws every slide fully built; on starts the slide again from its first click. */
    setFragmentsEnabled(enabled) {
        this.fragmentsEnabled = !!enabled;
        this.showFragment(enabled ? -1 : this.clickCount - 1, false);
    }

    // Nothing to lay out: the player's layers fill the stage, which the page sizes.
    layout() { }

    destroy() {
        if (this.disposed) return;
        this.disposed = true;
        try {
            this.screen?.dispose();
        } catch (error) {
            console.warn('[pptxViewer] the player could not be disposed', error);
        }
        this.screen = null;
        this.listeners.clear();
    }
}

// ---- One slide, in an article card --------------------------------------------------------------------------

// Each card's engine, by its card, from the moment it is made — before its first drawing is on the page, which
// is when articleViewer.js may already be putting that drawing's pictures back.
const articleEngines = new WeakMap();

/** The engine drawing an Article (PPTX) card, or null. */
export function articleEngineOf(element) {
    return articleEngines.get(element) || null;
}

class PptxArticleEngine extends PptxEngine {
    constructor(element, slide) {
        super();
        this.element = element;
        // Not `slide`: that is the method Reveal's callers step it by.
        this.data = slide;
        this.row = null;
        this.section = element.querySelector('section') || element;
        this.frame = element.querySelector('.pptx-stage') || this.section;
        articleEngines.set(element, this);
    }

    async initialize() {
        const player = await loadPlayer();
        if (!player || this.disposed) return false;

        // The drawing declares its own faces in a <style> of its own, resolved against the page when it was loaded.
        const row = rowOf(this.data);
        this.row = row;
        this.screen = player.embed(this.frame, { ...sizeOf(row, this.data), showNarration: false });

        // No transition in an article: each card is its own slide, arriving with the page — fully built.
        await this.screen.present(row, null, this.fragmentsEnabled ? null : clicksOf(row));
        this.settle();
        this.section.dataset.pptxClicks = String(this.screen.clickCount);
        return true;
    }

    getState() { return { indexh: 0, indexv: 0, indexf: this.fragmentIndex }; }

    getCurrentSlide() { return this.section; }

    slide(h, v, f) { this.showFragment(typeof f === 'number' ? f : -1, true); }

    /**
     * The card's drawing comes with its pictures held back (SlidePictureDeferral_Helper) and articleViewer.js puts
     * them back as the card nears the screen. This puts them back in the row too: the player draws the slide again
     * from it at every step, and the same row object is what its stage holds.
     */
    undeferPictures() {
        if (!this.row?.svg) return;
        this.row.svg = this.row.svg.replace(/ data-dmws-deferred="\d+"/, '').replace(/data-dmws-href=/g, 'href=');
    }

    destroy() {
        articleEngines.delete(this.element);
        super.destroy();
    }
}

/** An article slide's engine; call `initialize()` before anything else. */
export function createArticleEngine(element, slide) {
    return new PptxArticleEngine(element, slide);
}

// ---- The whole deck, as a slide show ------------------------------------------------------------------------

class PptxDeckEngine extends PptxEngine {
    constructor(container, slides) {
        super();
        this.container = container;
        this.frame = container.querySelector('.pptx-stage') || container;
        this.slides = slides || [];
        this.rows = new Map();
        this.indexh = 0;
        this.onKey = (event) => this.handleKey(event);
    }

    rowAt(index) {
        if (!this.rows.has(index)) this.rows.set(index, rowOf(this.slides[index] || {}));
        return this.rows.get(index);
    }

    async initialize() {
        const player = await loadPlayer();
        if (!player || this.disposed || this.slides.length === 0) return false;

        const first = this.rowAt(0);

        this.screen = player.embed(this.frame, { ...sizeOf(first, this.slides[0]), showNarration: false });
        await this.present(0, false, null);

        this.container.addEventListener('keydown', this.onKey);
        return true;
    }

    /**
     * Shows slide `index`, through its own transition when `withTransition`; `clicks` builds it that far at once.
     * With steps off it always arrives fully built.
     */
    present(index, withTransition, clicks) {
        const row = this.rowAt(index);
        this.indexh = index;
        this.container.dataset.pptxClicks = String(clicksOf(row));
        const built = this.fragmentsEnabled ? clicks : clicksOf(row);
        return this.screen.present(row, withTransition ? row.transition || null : null, built).then(() => this.settle());
    }

    getState() { return { indexh: this.indexh, indexv: 0, indexf: this.fragmentIndex }; }

    getCurrentSlide() { return this.container; }

    getTotalSlides() { return this.slides.length; }

    /**
     * Reveal's slide(h, v, f). Moving on to the next slide plays its transition, as the show does; any other
     * move is a cut — a transition is a way in from the slide before, and a Morph pairs shapes with exactly
     * that slide. A fragment given builds the slide that far at once; none plays its automatic step.
     */
    slide(h, v, f) {
        if (!this.screen || this.disposed) return;
        const target = Math.max(0, Math.min(typeof h === 'number' ? h : this.indexh, this.slides.length - 1));

        if (target === this.indexh) {
            if (typeof f === 'number') this.showFragment(f, true);
            return;
        }

        const withTransition = target === this.indexh + 1;
        const clicks = typeof f === 'number' && f >= 0 ? f + 1 : null;
        const arrived = this.present(target, withTransition, clicks);

        this.emit('slidechanged', { indexh: target, indexv: 0, currentSlide: this.container });
        arrived.then(() => {
            if (clicks != null) this.emit('fragmentshown', { fragmentIndex: this.fragmentIndex });
        });
    }

    next() {
        if (this.nextFragment()) return;
        if (this.indexh < this.slides.length - 1) this.slide(this.indexh + 1);
    }

    previous() {
        if (this.prevFragment()) return;
        if (this.indexh > 0) {
            // Back to the slide before, shown fully built, as the show goes back.
            const index = this.indexh - 1;
            this.slide(index, 0, clicksOf(this.rowAt(index)) - 1);
        }
    }

    handleKey(event) {
        switch (event.key) {
            case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'n': case 'N':
                this.next(); break;
            case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P':
                this.previous(); break;
            case 'Home':
                this.slide(0); break;
            case 'End':
                this.slide(this.slides.length - 1); break;
            default:
                return;
        }
        event.preventDefault();
    }

    destroy() {
        this.container.removeEventListener('keydown', this.onKey);
        super.destroy();
    }
}

/** The slide show's engine over `container`; call `initialize()` before anything else. */
export function createDeckEngine(container, slides) {
    return new PptxDeckEngine(container, slides);
}
