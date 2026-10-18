// The BlazorOffice slide show player (ESM module). Loaded lazily by SlideShow__Interop.
//
// The page is handed one slide at a time — its SVG (drawn by SlideSvg_Helper with `animatable`, so
// the background is grouped under [data-bo-bg], each shape is a [data-spid] group carrying its frame,
// and each paragraph a [data-para]), its transition, its effects with their behaviours, the steps
// (what each click plays and when each effect starts, worked out by SlideAnimationTimeline_Helper),
// its media as blob URLs, and the Morph pairs with the slide before it. The model travels in the
// shape of the DL entities (SlideTransition, SlideAnimationEffect, SlideAnimationBehavior,
// SlideAnimationStep_Row; camelCase); change them together.
//
// It plays what the timing tree says rather than a list of named effects: a behaviour moves
// (ppt_x/ppt_y), sizes (ppt_w/ppt_h), turns, fades, reveals through a filter, recolours or plays
// media, with PowerPoint's value formulas (#ppt_x, 1+#ppt_h/2, sin($*pi)…) evaluated here — so a
// preset nobody catalogued still plays. Transitions and Morph run on two stacked layers, the old
// slide under (or over) the new one.
//
// The show draws into an overlay of its own on <body>, not into the element it is given: a host
// page's layout can give an ancestor a transform (Radzen's body does), which turns position:fixed into
// "fixed to that ancestor", and the show must cover the window even where fullscreen is refused.
//
// Everything runs on one clock — the show's `clock`: on screen the real one (performance.now,
// requestAnimationFrame, setTimeout), for Export to Video a virtual one the exporter steps frame by
// frame (see VirtualClock and createVideoRun). Nothing in the player reads the time or schedules a
// frame or a timer except through it. A click while something is moving finishes it, as
// PowerPoint's "seek" does, rather than starting the next thing on top of it.

const NS = 'http://www.w3.org/2000/svg';

// ---- Entry points ------------------------------------------------------------------------------

/**
 * Draw ▸ Ink Replay: an overlay over the editor's slide showing the slide without its ink, and the strokes
 * drawn again point by point at the times the editor planned (SlideInkReplay_Helper: the times the ink was
 * drawn at, or an even pace) — then it removes itself; a click stops it. Runs on a clock like the player.
 */
class InkReplay {
    constructor(host, replay) {
        this.clock = realClock;
        this.disposed = false;
        this.strokes = replay.strokes || [];
        this.overlay = document.createElement('div');
        this.overlay.className = 'bo-show-preview bo-ink-replay';
        const frame = document.createElement('div');
        frame.className = 'bo-show-preview-frame';
        const layer = document.createElement('div');
        layer.className = 'bo-show-layer';
        layer.innerHTML = replay.svg || '';
        frame.appendChild(layer);
        this.ink = document.createElementNS(NS, 'svg');
        this.ink.setAttribute('class', 'bo-show-ink');
        this.ink.setAttribute('viewBox', `0 0 ${replay.width || 1280} ${replay.height || 720}`);
        this.ink.setAttribute('preserveAspectRatio', 'none');
        frame.appendChild(this.ink);
        this.overlay.appendChild(frame);
        host.appendChild(this.overlay);
        this.overlay.addEventListener('click', () => this.dispose());
        this.end = Math.max(0, ...this.strokes.map(s => s.timesMs?.[s.timesMs.length - 1] ?? 0));
        this.started = this.clock.now();
        const tick = () => {
            if (this.disposed) return;
            const t = this.clock.now() - this.started;
            this.draw(t);
            if (t < this.end) this.frameId = this.clock.frame(tick);
            else this.closeTimer = this.clock.after(() => this.dispose(), 900);
        };
        this.frameId = this.clock.frame(tick);
    }

    /** The strokes as they stand `t` milliseconds in: the points drawn by then, the last line to where the pen is. */
    draw(t) {
        let markup = '';
        for (const stroke of this.strokes) {
            const times = stroke.timesMs || [];
            const p = stroke.points || [];
            if (!times.length || times[0] > t) continue;
            let d = `M${p[0]} ${p[1]}`;
            let i = 1;
            for (; i < times.length && times[i] <= t; i++) d += `L${p[2 * i]} ${p[2 * i + 1]}`;
            if (i < times.length) {
                const span = times[i] - times[i - 1];
                const f = span > 0 ? (t - times[i - 1]) / span : 1;
                d += `L${lerp(p[2 * i - 2], p[2 * i], f)} ${lerp(p[2 * i - 1], p[2 * i + 1], f)}`;
            } else if (times.length === 1) d += 'l0.01 0';
            markup += `<path d="${d}" fill="none" stroke="#${stroke.colorHex}" stroke-width="${Math.max(0.5, stroke.widthPx)}" stroke-linecap="${stroke.flatTip ? 'square' : 'round'}" stroke-linejoin="round"${stroke.opacity < 1 ? ` opacity="${stroke.opacity}"` : ''}/>`;
        }
        this.ink.innerHTML = markup;
    }

    dispose() {
        if (this.disposed) return;
        this.disposed = true;
        this.clock.cancelFrame(this.frameId);
        this.clock.cancel(this.closeTimer);
        this.overlay.remove();
    }
}

/** Starts a show; `host` is the component's anchor element (the show draws in an overlay of its own). */
export function start(host, options, dotNetRef) {
    const show = new SlideShow(host, options, dotNetRef);
    show.begin();
    return show;
}

/** Plays a slide's transition (from `previous`, when given) and then all its steps, once, over `host`. */
export function preview(host, previous, slide, options) {
    return new SlidePreview(host, previous, slide, options);
}

/** Draw ▸ Ink Replay: the slide's strokes drawn again over `host`, in order and at their times. */
export function replayInk(host, replay) {
    return new InkReplay(host, replay);
}

/**
 * One slide at a time inside an element the host lays out itself — a card in a page, a slides pane — played the
 * way the show plays it (its transition from the slide before, its automatic step on arrival, each click on
 * request), with the host deciding when each click comes. See EmbeddedScreen.
 */
export function embed(frame, options) {
    return new EmbeddedScreen(frame, options);
}

// ---- Clocks ---------------------------------------------------------------------------------------

/**
 * The real clock a show on screen runs on. Every time the player reads, and every frame and timer
 * it schedules, goes through a clock — never performance.now, requestAnimationFrame or setTimeout
 * directly — so the same player can run on a VirtualClock.
 */
export const realClock = Object.freeze({
    virtual: false,
    now: () => performance.now(),
    frame: (callback) => requestAnimationFrame(callback),
    cancelFrame: (id) => cancelAnimationFrame(id),
    after: (callback, ms) => setTimeout(callback, ms),
    cancel: (id) => clearTimeout(id),
});

/**
 * A clock that only moves when told to: Export to Video steps it to each frame's exact time, so every
 * frame is drawn at that moment whatever the machine's speed, and none is dropped. Timers fire at
 * their own due times, in order, before the frame callbacks run at the time moved to.
 */
export class VirtualClock {
    constructor(start = 0) {
        this.virtual = true;
        this.time = start;
        this.frames = new Map();
        this.timers = [];
        this.nextId = 1;
    }

    now() { return this.time; }

    frame(callback) {
        const id = this.nextId++;
        this.frames.set(id, callback);
        return id;
    }

    cancelFrame(id) { this.frames.delete(id); }

    after(callback, ms) {
        const id = this.nextId++;
        this.timers.push({ id, due: this.time + Math.max(0, Number(ms) || 0), callback });
        return id;
    }

    cancel(id) { this.timers = this.timers.filter(timer => timer.id !== id); }

    /**
     * Moves the clock to `time`: each timer due by then fires at its own time (earliest first), then
     * the frame callbacks waiting run once, at `time`. `settle` (optional) is awaited after each, so
     * work a callback starts asynchronously finishes before the clock moves on.
     */
    async advanceTo(time, settle) {
        for (;;) {
            let next = null;
            for (const timer of this.timers) if (timer.due <= time && (!next || timer.due < next.due || (timer.due === next.due && timer.id < next.id))) next = timer;
            if (!next) break;
            this.timers.splice(this.timers.indexOf(next), 1);
            this.time = Math.max(this.time, next.due);
            next.callback();
            if (settle) await settle();
        }
        this.time = Math.max(this.time, time);
        const callbacks = [...this.frames.values()];
        this.frames.clear();
        for (const callback of callbacks) callback(this.time);
        if (settle) await settle();
    }
}

// ---- Timing helpers -----------------------------------------------------------------------------

/** SMIL's acceleration and deceleration: constant speed in the middle, linear ramps at the ends. */
function ease(u, a, d) {
    if (!a && !d) return u;
    if (a + d > 1) { const s = a + d; a /= s; d /= s; }
    const r = 1 / (1 - a / 2 - d / 2);
    if (u < a) return r * u * u / (2 * a);
    if (u <= 1 - d) return r * (u - a / 2);
    const t = 1 - u;
    return 1 - r * t * t / (2 * d);
}

const smooth = (from, to, p) => {
    if (p <= from) return 0;
    if (p >= to) return 1;
    const t = (p - from) / (to - from);
    return t * t * (3 - 2 * t);
};

const inOut = (p) => p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2;
const lerp = (a, b, p) => a + (b - a) * p;

// ---- Formulas -------------------------------------------------------------------------------------

/**
 * Evaluates a PowerPoint animation value: a number, or an expression over the shape's variables
 * (#ppt_x, #ppt_y, #ppt_w, #ppt_h — its centre and size in slide fractions), $ (the interpolated
 * value in a formula), pi, e, and the usual functions. Returns NaN for anything it cannot read.
 */
function evaluate(text, vars) {
    if (text === null || text === undefined) return NaN;
    const source = String(text).trim();
    if (source === '') return NaN;
    const direct = Number(source);
    if (!Number.isNaN(direct)) return direct;
    let i = 0;
    const peek = () => source[i];
    const skip = () => { while (i < source.length && /\s/.test(source[i])) i++; };
    const functions = {
        sin: Math.sin, cos: Math.cos, tan: Math.tan, asin: Math.asin, acos: Math.acos, atan: Math.atan,
        sqrt: Math.sqrt, abs: Math.abs, ln: Math.log, exp: Math.exp, floor: Math.floor, ceil: Math.ceil,
        rand: (x) => Math.random() * (x ?? 1), min: Math.min, max: Math.max,
    };
    function primary() {
        skip();
        const c = peek();
        if (c === '(') { i++; const v = expression(); skip(); if (peek() === ')') i++; return v; }
        if (c === '-') { i++; return -primary(); }
        if (c === '+') { i++; return primary(); }
        if (c === '$') { i++; return vars.$ ?? 0; }
        if (c === '#') {
            i++;
            const start = i;
            while (i < source.length && /[A-Za-z_0-9]/.test(source[i])) i++;
            return vars[source.slice(start, i)] ?? 0;
        }
        if (/[0-9.]/.test(c)) {
            const start = i;
            while (i < source.length && /[0-9.eE]/.test(source[i])) {
                if ((source[i] === 'e' || source[i] === 'E') && !/[0-9+-]/.test(source[i + 1] ?? '')) break;
                if ((source[i] === 'e' || source[i] === 'E') && /[+-]/.test(source[i + 1] ?? '')) i++;
                i++;
            }
            return Number(source.slice(start, i));
        }
        if (/[A-Za-z_]/.test(c)) {
            const start = i;
            while (i < source.length && /[A-Za-z_0-9]/.test(source[i])) i++;
            const name = source.slice(start, i).toLowerCase();
            if (name === 'pi') return Math.PI;
            if (name === 'e') return Math.E;
            if (vars[name] !== undefined) return vars[name];
            skip();
            if (peek() === '(' && functions[name]) {
                i++;
                const args = [];
                skip();
                if (peek() !== ')') {
                    args.push(expression());
                    skip();
                    while (peek() === ',') { i++; args.push(expression()); skip(); }
                }
                if (peek() === ')') i++;
                return functions[name](...args);
            }
            return NaN;
        }
        throw new Error('bad formula');
    }
    function power() {
        let v = primary();
        skip();
        while (peek() === '^') { i++; v = Math.pow(v, primary()); skip(); }
        return v;
    }
    function term() {
        let v = power();
        skip();
        while (peek() === '*' || peek() === '/') {
            const op = source[i++];
            const r = power();
            v = op === '*' ? v * r : v / r;
            skip();
        }
        return v;
    }
    function expression() {
        let v = term();
        skip();
        while (peek() === '+' || peek() === '-') {
            const op = source[i++];
            const r = term();
            v = op === '+' ? v + r : v - r;
            skip();
        }
        return v;
    }
    try { return expression(); } catch { return NaN; }
}

// ---- Clip shapes -----------------------------------------------------------------------------------

/** A seeded random sequence, so a dissolve's tiles stay put from one frame to the next. */
function seeded(seed) {
    let s = seed >>> 0 || 1;
    return () => { s ^= s << 13; s ^= s >>> 17; s ^= s << 5; return ((s >>> 0) % 100000) / 100000; };
}

const rect = (x, y, w, h) => w <= 0 || h <= 0 ? '' : `M${x} ${y}h${w}v${h}h${-w}Z`;
const polygon = (points) => points.length < 3 ? '' : 'M' + points.map(p => `${p[0]} ${p[1]}`).join('L') + 'Z';

/**
 * The visible part of a box of size w × h, `p` of the way through a reveal by `filter` — the
 * animEffect filters (wipe(down), circle(in), wheel(4)…) and the transitions built on them. Returns
 * { d, evenOdd } as path data in the box's own pixels, or null when the filter is a fade.
 */
function revealShape(filter, p, w, h, seed) {
    const match = /^([a-zA-Z]+)(?:\(([^)]*)\))?/.exec(filter || 'fade');
    const name = (match?.[1] || 'fade').toLowerCase();
    const arg = (match?.[2] || '').trim();
    p = Math.min(1, Math.max(0, p));
    const cx = w / 2, cy = h / 2;
    switch (name) {
        case 'wipe':
            // The direction names the side the reveal starts from: PowerPoint writes "Wipe From
            // Bottom" as wipe(down) and "From Left" as wipe(left).
            switch (arg) {
                case 'up': return { d: rect(0, 0, w, h * p) };
                case 'down': return { d: rect(0, h * (1 - p), w, h * p) };
                case 'right': return { d: rect(w * (1 - p), 0, w * p, h) };
                default: return { d: rect(0, 0, w * p, h) };
            }
        case 'barn': {
            const vertical = /vertical/i.test(arg);
            const inward = /^in/i.test(arg);
            if (vertical) {
                return inward
                    ? { d: rect(0, 0, w * p / 2, h) + rect(w - w * p / 2, 0, w * p / 2, h) }
                    : { d: rect(cx - w * p / 2, 0, w * p, h) };
            }
            return inward
                ? { d: rect(0, 0, w, h * p / 2) + rect(0, h - h * p / 2, w, h * p / 2) }
                : { d: rect(0, cy - h * p / 2, w, h * p) };
        }
        case 'blinds': {
            const n = 6;
            let d = '';
            for (let k = 0; k < n; k++) d += arg === 'vertical' ? rect(k * w / n, 0, w / n * p, h) : rect(0, k * h / n, w, h / n * p);
            return { d };
        }
        case 'box': {
            if (arg === 'out') return { d: rect(cx - w * p / 2, cy - h * p / 2, w * p, h * p) };
            const q = 1 - p;
            return { d: rect(0, 0, w, h) + rect(cx - w * q / 2, cy - h * q / 2, w * q, h * q), evenOdd: true };
        }
        case 'circle':
        case 'diamond': {
            const radius = name === 'circle' ? Math.hypot(cx, cy) : cx + cy;
            const shape = (r) => name === 'circle'
                ? `M${cx - r} ${cy}a${r} ${r} 0 1 0 ${2 * r} 0a${r} ${r} 0 1 0 ${-2 * r} 0Z`
                : polygon([[cx, cy - r], [cx + r, cy], [cx, cy + r], [cx - r, cy]]);
            if (arg === 'in') return { d: rect(0, 0, w, h) + shape(radius * (1 - p)), evenOdd: true };
            return { d: shape(radius * p) };
        }
        case 'plus': {
            const plus = (t) => {
                const q = t * t;
                return rect(cx - w * t / 2, cy - h * q / 2, w * t, h * q) + rect(cx - w * q / 2, cy - h * t / 2, w * q, h * t);
            };
            if (arg === 'in') return { d: rect(0, 0, w, h) + plus(1 - p), evenOdd: false };
            return { d: plus(p) };
        }
        case 'checkerboard': {
            const n = 8;
            let d = '';
            if (arg === 'down') {
                const cell = h / n;
                for (let col = 0; col < n; col++) {
                    const offset = (col % 2) * cell / 2;
                    for (let row = -1; row < n; row++) d += rect(col * w / n, row * cell + offset, w / n, cell * Math.min(1, p * 2));
                }
            } else {
                const cell = w / n;
                for (let row = 0; row < n; row++) {
                    const offset = (row % 2) * cell / 2;
                    for (let col = -1; col < n; col++) d += rect(col * cell + offset, row * h / n, cell * Math.min(1, p * 2), h / n);
                }
            }
            return { d };
        }
        case 'dissolve':
        case 'randombar': {
            const random = seeded(seed || 7);
            let d = '';
            if (name === 'dissolve') {
                const columns = 24, rows = Math.max(4, Math.round(24 * h / Math.max(1, w)));
                for (let row = 0; row < rows; row++)
                    for (let col = 0; col < columns; col++)
                        if (random() < p) d += rect(col * w / columns, row * h / rows, w / columns + 0.5, h / rows + 0.5);
            } else if (arg === 'vertical') {
                const n = 64;
                for (let k = 0; k < n; k++) if (random() < p) d += rect(k * w / n, 0, w / n + 0.5, h);
            } else {
                const n = 48;
                for (let k = 0; k < n; k++) if (random() < p) d += rect(0, k * h / n, w, h / n + 0.5);
            }
            return { d: d || 'M0 0Z' };
        }
        case 'strips': {
            // A diagonal edge sweeping from one corner to the opposite one.
            const t = p * 2;
            const corner = arg.toLowerCase();
            const fromRight = corner.includes('left');
            const fromBottom = corner.includes('up');
            const points = [];
            const X = (u) => fromRight ? w - u : u;
            const Y = (v) => fromBottom ? h - v : v;
            points.push([X(0), Y(0)], [X(Math.min(w, t * w)), Y(0)]);
            if (t > 1) points.push([X(w), Y((t - 1) * h)], [X(w), Y(h)], [X((t - 1) * w), Y(h)]);
            points.push([X(0), Y(Math.min(h, t * h))]);
            return { d: polygon(points) };
        }
        case 'wedge':
        case 'wheel': {
            const spokes = name === 'wheel' ? Math.max(1, parseInt(arg || '1', 10) || 1) : 1;
            const radius = Math.hypot(cx, cy) * 1.05;
            let d = '';
            if (name === 'wedge') {
                const sweep = p * Math.PI;
                d = fan(cx, cy, radius, -Math.PI / 2 - sweep, -Math.PI / 2 + sweep);
            } else {
                const each = 2 * Math.PI / spokes;
                for (let k = 0; k < spokes; k++) d += fan(cx, cy, radius, -Math.PI / 2 + k * each, -Math.PI / 2 + k * each + each * p);
            }
            return { d: d || 'M0 0Z' };
        }
        case 'slide':
            return { d: rect(0, 0, w, h), slide: arg };
        default:
            return null;
    }
}

function fan(cx, cy, r, from, to) {
    if (to - from <= 0.0001) return '';
    const points = [[cx, cy]];
    const steps = Math.max(2, Math.ceil((to - from) / 0.15));
    for (let k = 0; k <= steps; k++) {
        const a = from + (to - from) * k / steps;
        points.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
    }
    return polygon(points);
}

// ---- Colours ------------------------------------------------------------------------------------------

function parseColor(value) {
    if (!value) return null;
    const hex = String(value).replace('#', '').trim();
    if (/^[0-9a-f]{6}$/i.test(hex)) return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
    const rgb = /rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value);
    return rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])] : null;
}

const mixColor = (a, b, p) => `rgb(${Math.round(lerp(a[0], b[0], p))},${Math.round(lerp(a[1], b[1], p))},${Math.round(lerp(a[2], b[2], p))})`;

function rgbToHsl([r, g, b]) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l * 100];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h * 60, s * 100, l * 100];
}

function hslToRgb([h, s, l]) {
    h = ((h % 360) + 360) % 360 / 360; s = Math.min(100, Math.max(0, s)) / 100; l = Math.min(100, Math.max(0, l)) / 100;
    if (s === 0) return [l * 255, l * 255, l * 255];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, pp = 2 * l - q;
    const channel = (t) => {
        t = (t + 1) % 1;
        if (t < 1 / 6) return pp + (q - pp) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return pp + (q - pp) * (2 / 3 - t) * 6;
        return pp;
    };
    return [channel(h + 1 / 3) * 255, channel(h) * 255, channel(h - 1 / 3) * 255];
}

/** A colour moved by an animClr's "by": hsl(h,s,l) in degrees and percent, or rgb(r,g,b) in percent. */
function shiftColor(base, by) {
    const match = /^(hsl|rgb)\(([^)]*)\)$/.exec(by || '');
    if (!match) return null;
    const [a, b, c] = match[2].split(',').map(Number);
    if (match[1] === 'rgb') return [base[0] + a * 2.55, base[1] + b * 2.55, base[2] + c * 2.55].map(v => Math.min(255, Math.max(0, v)));
    const [h, s, l] = rgbToHsl(base);
    return hslToRgb([h + a, s + b, l + c]);
}

/** Between two colours through hue, saturation and lightness, the hue turning clockwise or anticlockwise. */
function mixHsl(from, to, p, direction) {
    const a = rgbToHsl(from), b = rgbToHsl(to);
    let dh = b[0] - a[0];
    if (direction === 'ccw') { if (dh > 0) dh -= 360; } else if (dh < 0) dh += 360;
    const [r, g, bl] = hslToRgb([a[0] + dh * p, lerp(a[1], b[1], p), lerp(a[2], b[2], p)]);
    return `rgb(${Math.round(r)},${Math.round(g)},${Math.round(bl)})`;
}

/** Which colour an attribute names: a shape's fill, its outline, or its text. */
function colorKind(attribute) {
    if (/stroke|line/i.test(attribute)) return 'stroke';
    if (/fill/i.test(attribute)) return 'fill';
    if (/color/i.test(attribute)) return 'text';
    return null;
}

/** tmFilter "t,p; t,p…" as pairs, cached on the behaviour. */
function timeFilter(behavior) {
    if (!behavior.timeFilter) return null;
    if (!behavior._timeFilter) {
        behavior._timeFilter = behavior.timeFilter.split(';').map(pair => pair.split(',').map(Number)).filter(pair => pair.length === 2 && pair.every(Number.isFinite));
    }
    return behavior._timeFilter.length > 1 ? behavior._timeFilter : null;
}

/** Progress through a tmFilter curve. */
function filtered(u, pairs) {
    if (!pairs) return u;
    for (let i = 1; i < pairs.length; i++) {
        const [t0, p0] = pairs[i - 1], [t1, p1] = pairs[i];
        if (u <= t1) return t1 > t0 ? lerp(p0, p1, (u - t0) / (t1 - t0)) : p1;
    }
    return pairs[pairs.length - 1][1];
}

/** Splits an element's text into word or letter spans (once per kind), for effects built a unit at a time. */
function splitUnits(root, type) {
    const cache = root.__boUnits || (root.__boUnits = {});
    if (cache[type]) return cache[type];
    const units = [];
    const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
        acceptNode: (node) => node.parentElement?.closest('foreignObject') || root.closest?.('foreignObject') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT,
    });
    const nodes = [];
    while (walker.nextNode()) if (/\S/.test(walker.currentNode.nodeValue)) nodes.push(walker.currentNode);
    for (const node of nodes) {
        const fragment = root.ownerDocument.createDocumentFragment();
        for (const token of node.nodeValue.split(/(\s+)/)) {
            if (token === '') continue;
            if (/^\s+$/.test(token)) { fragment.appendChild(root.ownerDocument.createTextNode(token)); continue; }
            const unit = (text) => {
                const span = root.ownerDocument.createElement('span');
                span.className = 'bo-show-unit';
                span.style.display = 'inline-block';
                span.style.whiteSpace = 'pre';
                span.textContent = text;
                units.push(span);
                return span;
            };
            if (type === 'wd') fragment.appendChild(unit(token));
            else {
                // Letters of a word stay on one line: a run of inline blocks may break anywhere.
                const word = root.ownerDocument.createElement('span');
                word.style.whiteSpace = 'nowrap';
                for (const letter of token) word.appendChild(unit(letter));
                fragment.appendChild(word);
            }
        }
        node.parentNode.replaceChild(fragment, node);
    }
    cache[type] = units;
    return units;
}

// ---- Targets --------------------------------------------------------------------------------------------

/** What a target looks like at one moment: the unanimated state is all defaults. */
function blankState() {
    return { visible: null, opacity: 1, dx: 0, dy: 0, sx: 1, sy: 1, rot: 0, clip: null, fill: null, stroke: null, text: null, weight: null, italic: null, underline: null, fontScale: 1, family: null, skew: 0 };
}

const copyState = (s) => ({ ...s });

/** The inline text properties effects paint. */
const TEXT_PROPERTIES = ['color', 'font-weight', 'font-style', 'text-decoration', 'font-family', 'font-size'];

/**
 * Paints a state's text overrides — colour, weight, italic, underline, size, font — onto text
 * elements, and takes them off again; the original font size is remembered per element.
 */
function paintText(elements, state, memo) {
    const active = state.text || state.weight || state.italic || state.underline || state.fontScale !== 1 || state.family;
    if (!active && !memo.painted) return;
    for (const element of elements) {
        const style = element.style;
        // The drawing's own inline text style, kept the first time an effect paints it, so what no
        // effect sets goes back to what the slide says rather than to the page's defaults.
        if (element.dataset.boTextStyle === undefined) {
            const own = {};
            for (const name of TEXT_PROPERTIES) own[name] = style.getPropertyValue(name);
            element.dataset.boTextStyle = JSON.stringify(own);
        }
        const own = JSON.parse(element.dataset.boTextStyle);
        const set = (name, value) => {
            if (value) style.setProperty(name, value, 'important');
            else if (own[name]) style.setProperty(name, own[name]);
            else style.removeProperty(name);
        };
        set('color', state.text);
        set('font-weight', state.weight);
        set('font-style', state.italic ? 'italic' : state.italic === false ? 'normal' : null);
        set('text-decoration', state.underline ? 'underline' : null);
        set('font-family', state.family ? `'${state.family}'` : null);
        if (state.fontScale !== 1) {
            if (!element.dataset.boFontSize) element.dataset.boFontSize = String(parseFloat(getComputedStyle(element).fontSize) || 16);
            set('font-size', `${Number(element.dataset.boFontSize) * state.fontScale}px`);
        } else {
            set('font-size', null);
        }
    }
    memo.painted = !!active;
}

/** A shape on the slide: its [data-spid] group, moved, scaled, turned and clipped in slide units. */
class ShapeTarget {
    constructor(element, stage) {
        this.element = element;
        this.stage = stage;
        const box = (element.getAttribute('data-box') || '0 0 0 0').split(' ').map(Number);
        this.box = { x: box[0], y: box[1], w: box[2], h: box[3] };
        this.rotation = Number(element.getAttribute('data-rot') || 0);
        // What the drawing says before any effect: a narration icon starts (and stays) hidden.
        this.initialVisibility = element.getAttribute('visibility');
        this.clipId = null;
        this.originalFill = null;
    }

    /** The shape's own variables for value formulas, in slide fractions. */
    variables() {
        const { width, height } = this.stage;
        return {
            ppt_x: (this.box.x + this.box.w / 2) / width,
            ppt_y: (this.box.y + this.box.h / 2) / height,
            ppt_w: this.box.w / width,
            ppt_h: this.box.h / height,
            ppt_r: this.rotation,
        };
    }

    size() { return { w: this.box.w, h: this.box.h }; }

    render(state) {
        const el = this.element;
        if (state.visible === false) el.setAttribute('visibility', 'hidden');
        else if (state.visible === true || this.initialVisibility === null) el.setAttribute('visibility', 'visible');
        else el.setAttribute('visibility', this.initialVisibility);
        if (state.opacity < 0.999) el.setAttribute('opacity', Math.max(0, state.opacity).toFixed(3));
        else el.removeAttribute('opacity');

        const moved = state.dx || state.dy || state.sx !== 1 || state.sy !== 1 || state.rot || state.skew;
        if (moved) {
            const cx = this.box.x + this.box.w / 2, cy = this.box.y + this.box.h / 2;
            const slideX = state.clip?.slideX || 0, slideY = state.clip?.slideY || 0;
            // xshear is a shear factor: -1 leans the shape 45 degrees.
            const skew = state.skew ? ` skewX(${Math.atan(state.skew) * 180 / Math.PI})` : '';
            el.setAttribute('transform', `translate(${state.dx + slideX} ${state.dy + slideY}) translate(${cx} ${cy}) rotate(${state.rot})${skew} scale(${Math.max(0.0001, state.sx)} ${Math.max(0.0001, state.sy)}) translate(${-cx} ${-cy})`);
        } else if (state.clip?.slideX || state.clip?.slideY) {
            el.setAttribute('transform', `translate(${state.clip.slideX || 0} ${state.clip.slideY || 0})`);
        } else {
            el.removeAttribute('transform');
        }

        if (state.clip?.d) {
            if (!this.clipId) {
                this.clipId = this.stage.newId();
                const clip = document.createElementNS(NS, 'clipPath');
                clip.setAttribute('id', this.clipId);
                clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
                const path = document.createElementNS(NS, 'path');
                clip.appendChild(path);
                this.stage.defs().appendChild(clip);
            }
            const path = this.stage.svg.querySelector(`#${this.clipId} path`);
            path.setAttribute('d', state.clip.d);
            path.setAttribute('transform', `translate(${this.box.x - (state.clip.slideX || 0)} ${this.box.y - (state.clip.slideY || 0)})`);
            path.setAttribute('clip-rule', state.clip.evenOdd ? 'evenodd' : 'nonzero');
            el.setAttribute('clip-path', `url(#${this.clipId})`);
        } else {
            el.removeAttribute('clip-path');
        }

        this.paintColors(state);
    }

    paintColors(state) {
        if (state.fill || this.filled) {
            for (const shape of this.element.querySelectorAll('path,rect,ellipse,circle,polygon')) {
                if (shape.closest('foreignObject') || shape.closest('clipPath') || shape.getAttribute('fill') === 'none') continue;
                shape.style.fill = state.fill || '';
            }
            this.filled = !!state.fill;
        }
        if (state.stroke || this.stroked) {
            for (const shape of this.element.querySelectorAll('path,rect,ellipse,circle,polygon')) {
                if (shape.closest('foreignObject') || shape.closest('clipPath') || shape.getAttribute('stroke') === 'none') continue;
                shape.style.stroke = state.stroke || '';
            }
            this.stroked = !!state.stroke;
        }
        this.textMemo = this.textMemo || {};
        paintText(this.element.querySelectorAll('foreignObject span, foreignObject div'), state, this.textMemo);
    }

    /** Where a colour change starts from: the fill, the outline or the text colour as drawn. */
    baseColor(kind) {
        if (kind === 'fill') return this.currentFill();
        if (kind === 'stroke') {
            const outline = [...this.element.querySelectorAll('path,rect')].find(s => !s.closest('foreignObject') && s.getAttribute('stroke') && s.getAttribute('stroke') !== 'none');
            return parseColor(outline?.getAttribute('stroke')) || [0, 0, 0];
        }
        const text = this.element.querySelector('foreignObject span, foreignObject div');
        return parseColor(text ? getComputedStyle(text).color : null) || [0, 0, 0];
    }

    /** The shape's first filled outline colour, where a colour change starts from. */
    currentFill() {
        if (this.originalFill === null) {
            const first = [...this.element.querySelectorAll('path,rect')].find(s => !s.closest('foreignObject') && s.getAttribute('fill') && s.getAttribute('fill') !== 'none');
            this.originalFill = parseColor(first?.getAttribute('fill')) || [128, 128, 128];
        }
        return this.originalFill;
    }

    reset() {
        this.render(blankState());
        this.element.removeAttribute('visibility');
    }
}

/**
 * Part of a shape's text — a paragraph (effects built by paragraph) or a word or letter (built a
 * unit at a time): an HTML element inside the drawing, with its own frame for the value formulas.
 */
class ParagraphTarget {
    constructor(element, shape, stage) {
        this.element = element;
        this.shape = shape;
        this.stage = stage;
        this.textMemo = {};
        element.style.transformOrigin = '50% 50%';
    }

    /** Its own centre and size in slide fractions, measured once before it moves. */
    variables() {
        if (this.box === undefined) this.box = this.stage.boxOf(this.element);
        if (!this.box) return this.shape.variables();
        const { width, height } = this.stage;
        return {
            ppt_x: (this.box.x + this.box.w / 2) / width,
            ppt_y: (this.box.y + this.box.h / 2) / height,
            ppt_w: this.box.w / width,
            ppt_h: this.box.h / height,
            ppt_r: this.shape.rotation,
        };
    }

    size() { return { w: this.element.offsetWidth || this.shape.box.w, h: this.element.offsetHeight || 20 }; }

    render(state) {
        const style = this.element.style;
        style.visibility = state.visible === false ? 'hidden' : '';
        style.opacity = state.opacity < 0.999 ? Math.max(0, state.opacity).toFixed(3) : '';
        const slideX = state.clip?.slideX || 0, slideY = state.clip?.slideY || 0;
        const moved = state.dx || state.dy || slideX || slideY || state.sx !== 1 || state.sy !== 1 || state.rot || state.skew;
        const skew = state.skew ? ` skewX(${Math.atan(state.skew) * 180 / Math.PI}deg)` : '';
        style.transform = moved ? `translate(${state.dx + slideX}px, ${state.dy + slideY}px) rotate(${state.rot}deg)${skew} scale(${state.sx}, ${state.sy})` : '';
        style.clipPath = state.clip?.d ? `path(${state.clip.evenOdd ? 'evenodd, ' : ''}'${state.clip.d}')` : '';
        paintText([this.element, ...this.element.querySelectorAll('span')], state, this.textMemo);
    }

    currentFill() { return this.shape.currentFill(); }

    baseColor(kind) {
        if (kind !== 'text') return this.shape.baseColor(kind);
        return parseColor(getComputedStyle(this.element).color) || [0, 0, 0];
    }

    reset() { this.render(blankState()); }
}

// ---- One slide on screen ---------------------------------------------------------------------------------

/**
 * One slide drawn into a layer, with its effects ready to play. Each target keeps a "held" state —
 * what finished behaviours left it as — and every frame recomputes the live state from it and the
 * behaviours still running.
 */
class SlideStage {
    constructor(layer, slide, show) {
        this.layer = layer;
        this.slide = slide;
        this.show = show;
        this.width = show.width;
        this.height = show.height;
        layer.innerHTML = slide.svg;
        this.svg = layer.querySelector('svg');
        if (this.svg) {
            this.svg.setAttribute('width', '100%');
            this.svg.setAttribute('height', '100%');
        }
        this.idCounter = 0;
        this.idPrefix = 'bos' + Math.random().toString(36).slice(2, 8);
        this.targets = new Map();
        this.held = new Map();
        this.running = [];
        this.stepIndex = 0;
        this.media = new Map();
        this.frame = 0;
        this.onIdle = null;
        this.disposed = false;

        this.prepareMedia();
        this.applyInitialVisibility();
    }

    /** The owner's clock (the show's, a screen's, a preview's): every time read and frame goes through it. */
    get clock() { return this.show.clock || realClock; }

    newId() { return `${this.idPrefix}-${this.idCounter++}`; }

    defs() {
        let defs = this.svg.querySelector(':scope > defs');
        if (!defs) {
            defs = document.createElementNS(NS, 'defs');
            this.svg.insertBefore(defs, this.svg.firstChild);
        }
        return defs;
    }

    shapeElement(id) { return this.svg?.querySelector(`[data-spid="${id}"]`) || null; }

    /** An element's frame in the drawing's units, from where the page lays it out. */
    boxOf(element) {
        const outer = this.svg?.getBoundingClientRect();
        const inner = element.getBoundingClientRect();
        if (!outer || !outer.width || !inner.width) return null;
        const k = this.width / outer.width;
        return { x: (inner.left - outer.left) * k, y: (inner.top - outer.top) * k, w: inner.width * k, h: inner.height * k };
    }

    /** The words or letters an effect built a unit at a time animates, in the order it plays them; null when it is not. */
    unitsOf(effect) {
        if (effect.iterateType !== 'wd' && effect.iterateType !== 'lt') return null;
        const element = this.shapeElement(effect.shapeId);
        if (!element) return null;
        const shapeKey = `s${effect.shapeId}`;
        if (!this.targets.has(shapeKey)) this.targets.set(shapeKey, new ShapeTarget(element, this));
        const shape = this.targets.get(shapeKey);
        const roots = [];
        if (effect.paragraphStart != null) {
            for (let index = effect.paragraphStart; index <= (effect.paragraphEnd ?? effect.paragraphStart); index++) {
                const paragraph = element.querySelector(`[data-para="${index}"]`);
                if (paragraph) roots.push(paragraph);
            }
        } else {
            roots.push(element);
        }
        const units = roots.flatMap(root => splitUnits(root, effect.iterateType));
        if (effect.iterateBackwards) units.reverse();
        return units.map(span => {
            if (!span.dataset.boUnit) span.dataset.boUnit = String(this.unitCounter = (this.unitCounter || 0) + 1);
            const key = `${shapeKey}u${span.dataset.boUnit}`;
            if (!this.targets.has(key)) this.targets.set(key, new ParagraphTarget(span, shape, this));
            return { key, target: this.targets.get(key) };
        });
    }

    /** How long one unit of an effect plays: its longest behaviour, with repeats and reversal. */
    oneUnitMs(effect) {
        let longest = 0;
        for (const behavior of effect.behaviors || []) {
            const length = Math.max(0, behavior.durationMs) * (behavior.autoReverse ? 2 : 1) * (behavior.repeat > 0 ? behavior.repeat : 1);
            longest = Math.max(longest, (behavior.offsetMs || 0) + length);
        }
        return longest;
    }

    /** The targets an effect animates: the shape, or each paragraph of its range. */
    targetsOf(effect, behavior) {
        const id = behavior?.shapeId ?? effect.shapeId;
        const element = this.shapeElement(id);
        if (!element) return [];
        const shapeKey = `s${id}`;
        if (!this.targets.has(shapeKey)) this.targets.set(shapeKey, new ShapeTarget(element, this));
        const shape = this.targets.get(shapeKey);
        if (behavior?.shapeId == null && effect.paragraphStart != null) {
            const result = [];
            const end = effect.paragraphEnd ?? effect.paragraphStart;
            for (let index = effect.paragraphStart; index <= end; index++) {
                const paragraph = element.querySelector(`[data-para="${index}"]`);
                if (!paragraph) continue;
                const key = `${shapeKey}p${index}`;
                if (!this.targets.has(key)) this.targets.set(key, new ParagraphTarget(paragraph, shape, this));
                result.push({ key, target: this.targets.get(key) });
            }
            return result;
        }
        return [{ key: shapeKey, target: shape }];
    }

    heldState(key) {
        if (!this.held.has(key)) this.held.set(key, blankState());
        return this.held.get(key);
    }

    /**
     * A shape whose first effect is an entrance starts hidden, as in PowerPoint; so does narration
     * and any media set to hide while stopped.
     */
    applyInitialVisibility() {
        const seen = new Set();
        const effects = this.slide.effects || [];
        if (this.show.options.showAnimations !== false) {
            for (const step of this.slide.steps || []) {
                for (const timing of step.effects) {
                    const effect = effects[timing.effectIndex];
                    if (!effect || effect.presetClass === 'mediacall') continue;
                    // Built a unit at a time, the words or letters start hidden as well as the shape.
                    const targets = [...this.targetsOf(effect), ...(this.unitsOf(effect) || [])];
                    for (const { key, target } of targets) {
                        if (seen.has(key)) continue;
                        seen.add(key);
                        if (effect.presetClass === 'entr') {
                            this.heldState(key).visible = false;
                            target.render(this.heldState(key));
                        }
                    }
                }
            }
        }
        for (const item of this.slide.media || []) {
            if (!(item.isNarration || item.hideWhenStopped)) continue;
            const element = this.shapeElement(item.shapeId);
            if (element) element.setAttribute('visibility', 'hidden');
        }
    }

    // ---- Media --------------------------------------------------------------------------------

    prepareMedia() {
        const virtual = !!this.clock.virtual;
        for (const item of this.slide.media || []) {
            if (!item.url) continue;
            let mediaElement = null;
            if (item.kind === 'video') {
                mediaElement = document.createElement('video');
                mediaElement.className = 'bo-show-video';
                mediaElement.playsInline = true;
                const x = item.x, y = item.y, w = item.width, h = item.height;
                mediaElement.style.left = `${100 * x / this.width}%`;
                mediaElement.style.top = `${100 * y / this.height}%`;
                mediaElement.style.width = `${100 * w / this.width}%`;
                mediaElement.style.height = `${100 * h / this.height}%`;
                mediaElement.style.display = 'none';
                this.layer.appendChild(mediaElement);
            } else if (!virtual) {
                mediaElement = new Audio();
            }
            if (mediaElement) {
                mediaElement.preload = 'auto';
                mediaElement.src = item.url;
                mediaElement.volume = Math.max(0, Math.min(1, item.volume ?? 1));
                mediaElement.muted = virtual || !!this.show.muted;
                // Show Media Controls: a film gets the browser's play bar.
                if (item.kind === 'video') mediaElement.controls = !virtual && this.show.options.showMediaControls !== false;
            }
            // On a virtual clock nothing plays: a stand-in keeps the media's place by the clock and
            // writes down what was heard when (Export to Video's sound track); a film keeps its
            // element, which the exporter seeks to draw each frame.
            const player = virtual ? new VirtualMedia(item, mediaElement, this.clock, this.show.mediaLog, this.show.mediaInfo?.[item.url]) : mediaElement;
            if (virtual) item.startSeconds = player.start;
            else this.watchMediaRange(item, player);
            this.media.set(item.shapeId, { item, player });

            // Clicking a sound or a film plays or pauses it, as PowerPoint's media action does.
            const element = this.shapeElement(item.shapeId);
            if (element && !item.isNarration) {
                element.style.cursor = 'pointer';
                element.addEventListener('click', (event) => {
                    event.stopPropagation();
                    this.command(item.shapeId, 'togglePause');
                });
                if (item.kind === 'video') mediaElement.addEventListener('click', (event) => { event.stopPropagation(); this.command(item.shapeId, 'togglePause'); });
            }
        }
    }

    /**
     * Trim, fades and loop, as PowerPoint plays them: playback runs from the trimmed start to the
     * trimmed end (the length less what is cut from the end), fading the volume in after the start
     * and out before the end; a looping sound goes back to the trimmed start.
     */
    watchMediaRange(item, player) {
        const base = Math.max(0, Math.min(1, item.volume ?? 1));
        const start = Math.max(0, (item.trimStartMs || 0) / 1000);
        const end = () => Number.isFinite(player.duration) ? Math.max(start, player.duration - Math.max(0, (item.trimEndMs || 0) / 1000)) : Infinity;
        const fadeIn = Math.max(0, (item.fadeInMs || 0) / 1000);
        const fadeOut = Math.max(0, (item.fadeOutMs || 0) / 1000);
        const trimmed = start > 0 || (item.trimEndMs || 0) > 0;
        if (item.loop && !trimmed) player.loop = true;
        player.addEventListener('timeupdate', () => {
            const at = player.currentTime;
            const stop = end();
            if (at >= stop - 0.02 && trimmed) {
                if (item.loop) { player.currentTime = start; return; }
                player.pause();
                return;
            }
            let gain = 1;
            if (fadeIn > 0 && at - start < fadeIn) gain = Math.min(gain, Math.max(0, (at - start) / fadeIn));
            if (fadeOut > 0 && Number.isFinite(stop) && stop - at < fadeOut) gain = Math.min(gain, Math.max(0, (stop - at) / fadeOut));
            player.volume = base * gain;
        });
        item.startSeconds = start;
    }

    command(shapeId, text) {
        const media = this.media.get(shapeId);
        if (!media) return;
        const { item, player } = media;
        if (item.isNarration && (this.show.options.showNarration === false || this.show.options.record)) return;
        const play = (from) => {
            // A play from the beginning starts at the trimmed start; a resume carries on where it paused.
            if (from !== undefined && !Number.isNaN(from)) from += item.startSeconds || 0;
            else if (player.currentTime < (item.startSeconds || 0)) from = item.startSeconds;
            if (from !== undefined && !Number.isNaN(from)) { try { player.currentTime = from; } catch { /* not seekable yet */ } }
            if (item.kind === 'video') player.style.display = '';
            if (item.hideWhenStopped && !item.isNarration) this.shapeElement(shapeId)?.setAttribute('visibility', 'visible');
            player.play().catch(() => { /* autoplay refused: a click on the page will start it */ });
        };
        const command = String(text || '').trim();
        const playFrom = /^playFrom\(([\d.]+)\)/i.exec(command);
        if (playFrom) play(Number(playFrom[1]));
        else if (/^play/i.test(command) || command === 'resume') play();
        else if (command === 'togglePause') { if (player.paused) play(); else player.pause(); }
        else if (command === 'pause') player.pause();
        else if (command === 'stop') { player.pause(); try { player.currentTime = 0; } catch { /* ignore */ } }
    }

    /** Narrations off pauses the one playing; on resumes it where it was. */
    setNarration(on) {
        for (const { item, player } of this.media.values()) {
            if (!item.isNarration) continue;
            if (!on) player.pause();
            else if (player.currentTime > 0 && !player.ended) player.play().catch(() => { });
        }
    }

    setMediaControls(on) {
        for (const { item, player } of this.media.values()) if (item.kind === 'video') player.controls = on;
    }

    /**
     * Stops the slide's sounds and films; with <c>carry</c>, a sound or film set to play across slides
     * that is playing is handed back instead, to go on over the next slides (a film in the screen's
     * carry layer, above them — ShowScreen.carry).
     */
    stopMedia(carry = false) {
        const carried = [];
        for (const { item, player } of this.media.values()) {
            if (this.handedOff?.has(player)) continue;
            if (carry && item.playAcrossSlides && !player.paused && !player.ended) {
                carried.push(player);
                (this.handedOff ||= new Set()).add(player);
                continue;
            }
            try { player.pause(); } catch { /* gone */ }
        }
        return carried;
    }

    // ---- Steps ----------------------------------------------------------------------------------

    get stepCount() { return (this.slide.steps || []).filter(s => s.triggerShapeId == null).length; }

    mainSteps() { return (this.slide.steps || []).filter(s => s.triggerShapeId == null); }

    hasPendingClick() { return this.stepIndex < this.mainSteps().length; }

    /** True while something with an end is still moving; an indefinite behaviour (until the slide ends, until the next click) does not count. */
    get isRunning() { return this.running.some(run => run.duration < 1e8 && run.repeat < 1e5); }

    /** Plays the automatic step, when the slide has one; true when it did. */
    playAutomatic() {
        const first = this.mainSteps()[0];
        if (!first?.isAutomatic || this.stepIndex !== 0) return false;
        this.playStep(first, false);
        return true;
    }

    playNextStep() {
        const step = this.mainSteps()[this.stepIndex];
        if (!step) return false;
        this.playStep(step, false);
        return true;
    }

    /** A click on a shape that triggers a sequence plays that sequence's next step. */
    playTrigger(shapeId) {
        const steps = (this.slide.steps || []).filter(s => s.triggerShapeId === shapeId);
        if (steps.length === 0) return false;
        this.triggerProgress = this.triggerProgress || new Map();
        const index = this.triggerProgress.get(shapeId) || 0;
        const step = steps[index % steps.length];
        this.triggerProgress.set(shapeId, index + 1);
        this.playStep(step, true);
        return true;
    }

    playStep(step, isTrigger, instant = false) {
        if (!isTrigger) this.stepIndex++;
        this.idleReported = false;
        const effects = this.slide.effects || [];
        const start = this.clock.now();
        for (const timing of step.effects) {
            const effect = effects[timing.effectIndex];
            if (!effect) continue;
            const effectRepeat = effect.repeat > 0 ? effect.repeat : 1;
            const effectReverse = effect.autoReverse;
            const effectStart = start + timing.startMs;
            const oneUnit = this.oneUnitMs(effect);

            // Built a word or a letter at a time: each unit plays the behaviours in turn, and the
            // shape itself (its fill and outline) comes and goes with the first and last unit.
            const units = this.unitsOf(effect);
            const interval = units ? (effect.iterateMs ?? (effect.iteratePercent ?? 0) / 100 * oneUnit) : 0;
            if (units && effect.presetClass === 'entr') {
                for (const { key, target } of this.targetsOf(effect)) {
                    this.running.push({ key, target, effect, behavior: { kind: 'set', attribute: 'style.visibility', to: 'visible', durationMs: 1 }, start: effectStart, duration: 1, repeat: 1, reverse: false, fired: false, seed: 1 });
                }
            }
            if (units && effect.presetClass === 'exit') {
                for (const { key, target } of this.targetsOf(effect)) {
                    this.running.push({ key, target, effect, behavior: { kind: 'set', attribute: 'style.visibility', to: 'hidden', durationMs: 1 }, start: effectStart + oneUnit + interval * Math.max(0, units.length - 1), duration: 1, repeat: 1, reverse: false, fired: false, seed: 1 });
                }
            }
            const passes = units ? units.map((unit, index) => ({ unit, offset: index * interval })) : [{ unit: null, offset: 0 }];
            for (const pass of passes) {
                for (const behavior of effect.behaviors || []) {
                    const targets = pass.unit && behavior.shapeId == null && behavior.kind !== 'command' ? [pass.unit] : this.targetsOf(effect, behavior);
                    for (const { key, target } of targets) {
                        this.running.push({
                            key, target, behavior, effect,
                            effectStart: effectStart + pass.offset,
                            effectDuration: oneUnit,
                            start: effectStart + pass.offset + (behavior.offsetMs || 0),
                            duration: Math.max(1, behavior.durationMs < 0 ? 1e9 : behavior.durationMs),
                            repeat: (behavior.repeat < 0 ? 1e6 : (behavior.repeat || 1)) * effectRepeat,
                            reverse: behavior.autoReverse || effectReverse,
                            fired: false,
                            seed: (timing.effectIndex + 1) * 7919,
                        });
                    }
                    if (pass.unit && behavior.kind === 'command') break;
                }
            }
        }
        if (instant) this.finish(true);
        else this.loop();
    }

    /** Skips every running behaviour to its end — a click while something is moving. */
    finish(silent = false) {
        for (const run of this.running) run.start = -Infinity;
        this.tick(Infinity, silent);
    }

    /** Plays the first `count` main steps instantly (going back to a slide, or one step back). */
    applySteps(count) {
        const steps = this.mainSteps();
        for (let k = 0; k < count && k < steps.length; k++) this.playStep(steps[k], false, true);
    }

    loop() {
        if (this.frame || this.disposed) return;
        const frame = () => {
            this.frame = 0;
            if (this.disposed) return;
            this.tick(this.clock.now(), false);
            if (!this.isRunning && !this.idleReported) { this.idleReported = true; this.onIdle?.(); }
            if (this.running.length > 0) this.frame = this.clock.frame(frame);
        };
        this.frame = this.clock.frame(frame);
    }

    /**
     * One frame: behaviours that have ended fold into their target's held state (unless they
     * remove themselves), the ones still running are applied on top of it, and each touched
     * target is redrawn.
     */
    tick(time, silent) {
        const touched = new Map();
        const live = new Map();
        const remaining = [];
        this.running.sort((a, b) => a.start - b.start);
        for (const run of this.running) {
            const at = this.effectTime(run, time);
            if (at < run.start) { remaining.push(run); continue; }
            const span = run.duration * (run.reverse ? 2 : 1) * run.repeat;
            const ended = at === Infinity || at - run.start >= span;
            if (run.behavior.kind === 'command') {
                if (!run.fired && !silent) this.command(run.behavior.shapeId ?? run.effect.shapeId, run.behavior.command);
                run.fired = true;
                continue;
            }
            touched.set(run.key, run.target);
            if (ended) {
                if (run.behavior.hold !== false) this.apply(run, this.heldState(run.key), filtered(run.reverse ? 0 : 1, timeFilter(run.behavior)), true);
                continue;
            }
            remaining.push(run);
        }
        for (const run of remaining) {
            const at = this.effectTime(run, time);
            if (at < run.start) continue;
            if (!live.has(run.key)) live.set(run.key, copyState(this.heldState(run.key)));
            const elapsed = at - run.start;
            const cycle = run.duration * (run.reverse ? 2 : 1);
            const within = elapsed % cycle;
            let u = run.reverse && within > run.duration ? 1 - (within - run.duration) / run.duration : within / run.duration;
            u = ease(Math.min(1, Math.max(0, u)), run.behavior.accelerate || 0, run.behavior.decelerate || 0);
            this.apply(run, live.get(run.key), filtered(u, timeFilter(run.behavior)), false);
            touched.set(run.key, run.target);
        }
        this.running = remaining;
        for (const [key, target] of touched) target.render(live.get(key) || this.heldState(key));
    }

    /**
     * The time a behaviour sees: the clock, eased through its effect's smooth start and end when
     * the effect has them (PowerPoint puts a motion path's smoothing on the effect, not the path).
     */
    effectTime(run, time) {
        const a = run.effect?.accelerate || 0, d = run.effect?.decelerate || 0;
        if ((!a && !d) || time === Infinity || run.effectStart === undefined) return time;
        const local = time - run.effectStart, length = run.effectDuration;
        if (!(length > 0) || local <= 0 || local >= length) return time;
        return run.effectStart + length * ease(local / length, a, d);
    }

    /** Applies one behaviour, `p` of the way through, to a state. */
    apply(run, state, p, committing) {
        const { behavior, target } = run;
        const vars = target.variables();
        const attribute = behavior.attribute || '';
        switch (behavior.kind) {
            case 'set':
                this.setValue(state, attribute, behavior.to, vars);
                break;
            case 'anim': {
                const value = this.animValue(behavior, p, vars, state, attribute);
                if (value !== null) this.setValue(state, attribute, value, vars, behavior.additive === 'sum');
                break;
            }
            case 'effect': {
                const out = behavior.transition === 'out';
                const progress = out ? 1 - p : p;
                // An "image" filter is a picture adjustment; its opacity setting is what shows.
                if (behavior.filter === 'image') {
                    const opacity = /opacity:\s*([\d.]+)/.exec(behavior.filterProperties || '');
                    if (opacity) state.opacity *= Number(opacity[1]);
                    break;
                }
                if (/^fade/i.test(behavior.filter || 'fade') || !behavior.filter) {
                    // Finished, a fade leaves the opacity alone: visibility does the hiding, so a
                    // later entrance can fade the shape back in. An "out" that its time filter
                    // brings back (Pulse) does not hide.
                    if (committing) { state.opacity = 1; if (out && p >= 1) state.visible = false; }
                    else state.opacity *= progress;
                    break;
                }
                if (committing) {
                    state.clip = null;
                    if (out && p >= 1) state.visible = false;
                    break;
                }
                const { w, h } = target.size();
                const shape = revealShape(behavior.filter, progress, w, h, run.seed);
                if (!shape) { state.opacity *= progress; break; }
                if (shape.slide) {
                    const offset = 1 - progress;
                    state.clip = {
                        d: rect(0, 0, w, h),
                        slideX: shape.slide === 'fromLeft' ? -w * offset : shape.slide === 'fromRight' ? w * offset : 0,
                        slideY: shape.slide === 'fromTop' ? -h * offset : shape.slide === 'fromBottom' ? h * offset : 0,
                    };
                } else {
                    state.clip = shape;
                }
                break;
            }
            case 'scale': {
                const from = { x: (behavior.scaleFromX ?? state.sx * 100) / 100, y: (behavior.scaleFromY ?? state.sy * 100) / 100 };
                if (behavior.scaleToX != null || behavior.scaleToY != null) {
                    state.sx = lerp(from.x, (behavior.scaleToX ?? 100) / 100, p);
                    state.sy = lerp(from.y, (behavior.scaleToY ?? 100) / 100, p);
                } else {
                    state.sx *= lerp(1, (behavior.scaleByX ?? 100) / 100, p);
                    state.sy *= lerp(1, (behavior.scaleByY ?? 100) / 100, p);
                }
                break;
            }
            case 'rotate':
                if (behavior.rotateFrom != null || behavior.rotateTo != null) state.rot = lerp(behavior.rotateFrom ?? state.rot, behavior.rotateTo ?? 0, p);
                else state.rot += (behavior.rotateBy ?? 0) * p;
                break;
            case 'motion': {
                const point = this.pathPoint(behavior.path, p);
                state.dx += point.x;
                state.dy += point.y;
                break;
            }
            case 'color': {
                const kind = colorKind(attribute) || 'fill';
                const from = parseColor(behavior.colorFrom) || target.baseColor(kind);
                const to = parseColor(behavior.colorTo) || shiftColor(from, behavior.colorBy);
                if (!to) break;
                state[kind] = behavior.colorSpace === 'hsl' ? mixHsl(from, to, p, behavior.colorDirection) : mixColor(from, to, p);
                break;
            }
        }
    }

    /** The value of an anim at progress p: its keyframes (with formulas), or from/to/by. */
    animValue(behavior, p, vars, state, attribute) {
        const frames = behavior.keyframes || [];
        const numeric = behavior.valueType !== 'str' && behavior.valueType !== 'clr';
        if (frames.length > 0) {
            let index = frames.findIndex(f => f.time >= p);
            if (index < 0) index = frames.length - 1;
            const after = frames[index];
            const before = frames[Math.max(0, index - 1)];
            if (!numeric || behavior.calcMode === 'discrete') {
                const pick = behavior.calcMode === 'discrete' ? (p >= after.time ? after : before) : (p < 1 ? before : after);
                return pick.value;
            }
            const a = evaluate(before.value, vars);
            const b = evaluate(after.value, vars);
            const span = after.time - before.time;
            const local = span > 0 ? (p - before.time) / span : 1;
            let value = Number.isNaN(a) ? b : Number.isNaN(b) ? a : lerp(a, b, local);
            const formula = local < 1 ? before.formula : after.formula;
            if (formula) {
                const computed = evaluate(formula, { ...vars, $: value });
                if (!Number.isNaN(computed)) value = computed;
            }
            return Number.isNaN(value) ? null : value;
        }
        const current = this.currentValue(state, attribute, vars);
        const from = behavior.from != null ? evaluate(behavior.from, vars) : current;
        let to = behavior.to != null ? evaluate(behavior.to, vars) : NaN;
        if (Number.isNaN(to) && behavior.by != null) to = from + evaluate(behavior.by, vars);
        if (Number.isNaN(from) || Number.isNaN(to)) return null;
        return lerp(from, to, p);
    }

    currentValue(state, attribute, vars) {
        switch (attribute) {
            case 'ppt_x': return vars.ppt_x + state.dx / this.width;
            case 'ppt_y': return vars.ppt_y + state.dy / this.height;
            case 'ppt_w': return vars.ppt_w * state.sx;
            case 'ppt_h': return vars.ppt_h * state.sy;
            case 'style.rotation': case 'r': return vars.ppt_r + state.rot;
            case 'style.opacity': return state.opacity;
            case 'style.fontSize': return state.fontScale;
            case 'xshear': return state.skew;
            default: return 0;
        }
    }

    setValue(state, attribute, value, vars, additive = false) {
        switch (attribute) {
            case 'style.visibility':
                state.visible = String(value) !== 'hidden';
                return;
            case 'style.fontWeight':
                state.weight = String(value) === 'normal' ? null : String(value);
                return;
            case 'style.fontStyle':
                state.italic = String(value) === 'italic';
                return;
            case 'style.textDecorationUnderline':
                state.underline = String(value) === 'true';
                return;
            case 'style.fontFamily':
                state.family = String(value) || null;
                return;
        }
        // A colour-valued attribute: a colour anim's or set's value is a colour.
        const kind = colorKind(attribute);
        if (kind && typeof value === 'string' && /^#?[0-9a-f]{6}$/i.test(value.trim())) {
            state[kind] = value.trim().startsWith('#') ? value.trim() : `#${value.trim()}`;
            return;
        }
        const number = typeof value === 'number' ? value : evaluate(value, vars);
        if (Number.isNaN(number)) return;
        switch (attribute) {
            case 'ppt_x': state.dx = additive ? state.dx + number * this.width : (number - vars.ppt_x) * this.width; break;
            case 'ppt_y': state.dy = additive ? state.dy + number * this.height : (number - vars.ppt_y) * this.height; break;
            case 'ppt_w': if (vars.ppt_w > 0) state.sx = additive ? state.sx + number / vars.ppt_w : number / vars.ppt_w; break;
            case 'ppt_h': if (vars.ppt_h > 0) state.sy = additive ? state.sy + number / vars.ppt_h : number / vars.ppt_h; break;
            case 'style.rotation': case 'r': state.rot = additive ? state.rot + number : number - vars.ppt_r; break;
            case 'style.opacity': state.opacity = Math.max(0, Math.min(1, number)); break;
            case 'style.fontSize': state.fontScale = Math.max(0.05, additive ? state.fontScale + number : number); break;
            case 'xshear': state.skew = additive ? state.skew + number : number; break;
        }
    }

    /** A point along a motion path (slide fractions, relative to the shape), in slide units. */
    pathPoint(path, p) {
        if (!path) return { x: 0, y: 0 };
        this.paths = this.paths || new Map();
        let element = this.paths.get(path);
        if (!element) {
            element = document.createElementNS(NS, 'path');
            element.setAttribute('d', scalePath(path, this.width, this.height));
            element.setAttribute('fill', 'none');
            element.style.display = 'none';
            this.defs().appendChild(element);
            element._length = (() => { try { return element.getTotalLength(); } catch { return 0; } })();
            this.paths.set(path, element);
        }
        if (!element._length) return { x: 0, y: 0 };
        const point = element.getPointAtLength(element._length * Math.min(1, Math.max(0, p)));
        return { x: point.x, y: point.y };
    }

    dispose() {
        this.disposed = true;
        if (this.frame) this.clock.cancelFrame(this.frame);
        this.frame = 0;
        this.stopMedia();
        // A sound or film handed to the show (playing across slides) is the show's to stop now.
        for (const { player } of this.media.values()) {
            if (this.handedOff?.has(player)) continue;
            try { player.removeAttribute('src'); player.load(); } catch { /* gone */ }
        }
        this.media.clear();
    }
}

// ---- Media on a virtual clock ------------------------------------------------------------------------------

/** The <video> element of a film's player (the element itself on screen, a stand-in's on a virtual clock); null for a sound. */
function filmElement(player) {
    const element = player instanceof VirtualMedia ? player.element : player;
    return element?.tagName === 'VIDEO' ? element : null;
}

/**
 * A sound or film on a virtual clock (Export to Video). It never plays: it knows where it is by the
 * clock — the trimmed start and end, and the loop, as the real player applies them (watchMediaRange)
 * — and writes each stretch that would have been heard into `log`, with its volume, fades and trims,
 * which the exporter mixes into the sound track. A film keeps its <video> element (muted, never
 * playing), which the exporter seeks to `positionAt(time)` to draw each frame. The members are the
 * ones the stage uses of a media element.
 */
class VirtualMedia {
    constructor(item, element, clock, log, info) {
        this.item = item;
        this.element = element;
        this.clock = clock;
        this.log = log || [];
        this.duration = Number(info?.duration);
        this.style = element ? element.style : {};
        this.volume = 1;
        this.muted = true;
        this.controls = false;
        this.preload = 'auto';
        this.loop = false;
        this.src = item.url;
        this.isPaused = true;
        this.position = 0;
        this.anchor = 0;
        this.segment = null;
        if (element) element.boMedia = this;
    }

    /** The trimmed start, in seconds of the media. */
    get start() { return Math.max(0, (this.item.trimStartMs || 0) / 1000); }

    /** The trimmed end (the length less what is cut from the end), or Infinity before the length is known. */
    get stop() {
        return Number.isFinite(this.duration) && this.duration > 0
            ? Math.max(this.start, this.duration - Math.max(0, (this.item.trimEndMs || 0) / 1000))
            : Infinity;
    }

    /** Where the media is at clock time `ms`, in seconds: it runs from where it started, stops at the trimmed end, or loops back to the trimmed start. */
    positionAt(ms) {
        if (this.isPaused) return this.position;
        let at = this.position + Math.max(0, ms - this.anchor) / 1000;
        const stop = this.stop;
        if (at >= stop) {
            const span = stop - this.start;
            at = this.item.loop && span > 0.01 ? this.start + ((at - this.start) % span) : stop;
        }
        return at;
    }

    get currentTime() { return this.positionAt(this.clock.now()); }

    set currentTime(value) {
        const playing = !this.isPaused;
        if (playing) this.close();
        this.position = Math.max(0, Number(value) || 0);
        this.anchor = this.clock.now();
        if (playing) this.open();
    }

    /** True once a sound or film that does not loop has reached its trimmed end. */
    get ended() { return !this.isPaused && !this.item.loop && this.positionAt(this.clock.now()) >= this.stop; }

    get paused() { return this.isPaused || this.ended; }

    play() {
        if (this.isPaused) {
            this.isPaused = false;
            this.anchor = this.clock.now();
            this.open();
        }
        return Promise.resolve();
    }

    pause() {
        if (this.isPaused) return;
        this.position = this.positionAt(this.clock.now());
        this.isPaused = true;
        this.close();
    }

    /** A stretch heard from now: from this position in the media, until paused, moved or ended. */
    open() {
        const item = this.item;
        this.segment = {
            url: item.url,
            kind: item.kind,
            isNarration: !!item.isNarration,
            atMs: this.clock.now(),
            fromSeconds: this.position,
            untilMs: null,
            volume: Math.max(0, Math.min(1, item.volume ?? 1)),
            fadeInSeconds: Math.max(0, (item.fadeInMs || 0) / 1000),
            fadeOutSeconds: Math.max(0, (item.fadeOutMs || 0) / 1000),
            startSeconds: this.start,
            trimEndSeconds: Math.max(0, (item.trimEndMs || 0) / 1000),
            loop: !!item.loop,
        };
        this.log.push(this.segment);
    }

    close() {
        if (!this.segment) return;
        this.segment.untilMs = this.clock.now();
        this.segment = null;
    }

    addEventListener() { /* a stand-in raises no events: the exporter asks it where it is */ }
    removeEventListener() { }

    removeAttribute() {
        this.pause();
        this.element?.removeAttribute('src');
    }

    load() { this.element?.load(); }
}

/** A PowerPoint motion path ("M 0 0 L 0.25 0.1 E", slide fractions) as SVG path data in slide units. */
function scalePath(path, width, height) {
    const tokens = String(path).replace(/([MLCZEmlcze])/g, ' $1 ').trim().split(/[\s,]+/);
    let out = '';
    let axis = 0;
    for (const token of tokens) {
        if (/^[MLCZmlcz]$/.test(token)) { out += token.toUpperCase() === 'Z' ? 'Z ' : token.toUpperCase() + ' '; axis = 0; continue; }
        if (/^[Ee]$/.test(token)) continue;
        const value = Number(token);
        if (Number.isNaN(value)) continue;
        out += (axis % 2 === 0 ? value * width : value * height).toFixed(3) + ' ';
        axis++;
    }
    return out.trim() || 'M 0 0';
}

// ---- Transitions ----------------------------------------------------------------------------------------

/**
 * Plays a transition between two layers, `p` driven by the shared clock. `done` is called once
 * at the end (or when skipped). Returns a function that finishes it at once.
 */
function playTransition(frame, oldLayer, newLayer, transition, oldStage, newStage, morphPairs, done, clock = realClock) {
    const effect = transition?.effect || 'none';
    // p14:dur="0" is an instant transition, not the default length.
    const duration = Math.max(1, transition?.durationMs ?? 700);
    const direction = transition?.direction || '';
    const option = transition?.option || '';
    const width = frame.clientWidth, height = frame.clientHeight;
    let raf = 0;
    let finished = false;
    const startTime = clock.now();
    let render = () => { };
    let cleanup = () => { };

    const vector = (dir) => ({
        l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1],
    })[dir] || [-1, 0];

    const clip = (layer, shape) => {
        layer.style.clipPath = shape ? `path(${shape.evenOdd ? 'evenodd, ' : ''}'${shape.d}')` : '';
    };

    const reset = () => {
        for (const layer of [oldLayer, newLayer]) {
            if (!layer) continue;
            layer.style.transform = '';
            layer.style.opacity = '';
            layer.style.clipPath = '';
            layer.style.filter = '';
            layer.style.zIndex = '';
            layer.style.transformOrigin = '';
            layer.style.backfaceVisibility = '';
        }
    };

    if (!oldLayer || effect === 'none') {
        // No old slide, or no transition: straight on.
        done();
        return () => { };
    }

    const overlay = (color) => {
        const element = document.createElement('div');
        element.className = 'bo-show-flash';
        element.style.background = color;
        frame.appendChild(element);
        return element;
    };

    switch (effect) {
        case 'bo-zoom-in': case 'bo-zoom-out': {
            // transition.zoom: the tile as fractions of the slide. Zooming in, the destination grows from the
            // tile to the whole slide while the zoom's slide swells past it and fades; zooming out, the reverse.
            const tile = transition.zoom || { x: 0.25, y: 0.25, w: 0.5, h: 0.5 };
            const into = (layer, e) => {
                // e = 0: the layer fills the tile; e = 1: the whole slide.
                const sx = tile.w + (1 - tile.w) * e, sy = tile.h + (1 - tile.h) * e;
                layer.style.transformOrigin = '0 0';
                layer.style.transform = `translate(${tile.x * (1 - e) * 100}%, ${tile.y * (1 - e) * 100}%) scale(${sx}, ${sy})`;
            };
            const past = (layer, e) => {
                // The slide behind swells about the tile's centre so the tile comes to fill the view.
                const k = 1 + e * (1 / Math.max(tile.w, tile.h) - 1);
                const cx = (tile.x + tile.w / 2) * 100, cy = (tile.y + tile.h / 2) * 100;
                layer.style.transformOrigin = `${cx}% ${cy}%`;
                layer.style.transform = `scale(${k})`;
            };
            if (effect === 'bo-zoom-in') {
                newLayer.style.zIndex = '2';
                render = (p) => { const e = inOut(p); into(newLayer, e); past(oldLayer, e); oldLayer.style.opacity = String(1 - e); };
            } else {
                oldLayer.style.zIndex = '2';
                render = (p) => { const e = 1 - inOut(p); into(oldLayer, e); past(newLayer, e); newLayer.style.opacity = String(1 - e); };
            }
            break;
        }
        case 'cut': {
            const black = option === '1' || option === 'true' ? overlay('#000') : null;
            render = (p) => {
                if (black) black.style.opacity = p < 0.5 ? '1' : '0';
                newLayer.style.opacity = p < 0.5 && black ? '0' : '1';
            };
            cleanup = () => black?.remove();
            break;
        }
        case 'fade': {
            const throughBlack = option === '1' || option === 'true';
            render = (p) => {
                if (throughBlack) {
                    oldLayer.style.opacity = String(1 - smooth(0, 0.5, p));
                    newLayer.style.opacity = String(smooth(0.5, 1, p));
                } else {
                    newLayer.style.opacity = String(p);
                }
            };
            break;
        }
        case 'push': case 'pan': case 'gallery': case 'conveyor': case 'ferris': {
            const [vx, vy] = vector(direction || 'l');
            render = (p) => {
                const e = inOut(p);
                oldLayer.style.transform = `translate(${vx * e * 100}%, ${vy * e * 100}%)`;
                newLayer.style.transform = `translate(${-vx * (1 - e) * 100}%, ${-vy * (1 - e) * 100}%)`;
            };
            break;
        }
        case 'cover': {
            const [vx, vy] = vector(direction || 'l');
            render = (p) => { const e = inOut(p); newLayer.style.transform = `translate(${-vx * (1 - e) * 100}%, ${-vy * (1 - e) * 100}%)`; };
            break;
        }
        case 'pull': case 'reveal': case 'wind': {
            const [vx, vy] = vector(direction || 'l');
            const throughBlack = effect === 'reveal' && (option === '1' || option === 'true');
            oldLayer.style.zIndex = '2';
            render = (p) => {
                const e = inOut(p);
                oldLayer.style.transform = `translate(${vx * e * 100}%, ${vy * e * 100}%)` + (effect === 'wind' ? ` skewX(${-15 * e}deg)` : '');
                if (effect !== 'pull') oldLayer.style.opacity = String(1 - e);
                if (throughBlack) newLayer.style.opacity = String(smooth(0.5, 1, p));
            };
            break;
        }
        case 'wipe': {
            render = (p) => {
                const q = (1 - inOut(p)) * 100;
                const inset = { l: `0 0 0 ${q}%`, r: `0 ${q}% 0 0`, u: `${q}% 0 0 0`, d: `0 0 ${q}% 0` }[direction || 'l'] || `0 0 0 ${q}%`;
                newLayer.style.clipPath = `inset(${inset})`;
            };
            break;
        }
        case 'split': case 'curtains': {
            const vertical = transition.orientation !== 'horz';
            const inward = direction === 'in' || effect === 'curtains';
            if (inward) {
                oldLayer.style.zIndex = '2';
                render = (p) => {
                    const q = inOut(p) * 50;
                    oldLayer.style.clipPath = vertical ? `inset(0 ${q}% 0 ${q}%)` : `inset(${q}% 0 ${q}% 0)`;
                };
            } else {
                render = (p) => {
                    const q = (1 - inOut(p)) * 50;
                    newLayer.style.clipPath = vertical ? `inset(0 ${q}% 0 ${q}%)` : `inset(${q}% 0 ${q}% 0)`;
                };
            }
            break;
        }
        case 'zoom': case 'warp': case 'flythrough': {
            const inward = direction !== 'out';
            render = (p) => {
                const e = inOut(p);
                if (inward) {
                    newLayer.style.transform = `scale(${lerp(effect === 'flythrough' ? 0.4 : 0.1, 1, e)})`;
                    newLayer.style.opacity = String(e);
                } else {
                    oldLayer.style.zIndex = '2';
                    oldLayer.style.transform = `scale(${lerp(1, effect === 'warp' ? 0.1 : 2.5, e)})`;
                    oldLayer.style.opacity = String(1 - e);
                }
            };
            break;
        }
        case 'newsflash': case 'vortex':
            render = (p) => {
                const e = inOut(p);
                newLayer.style.transform = `rotate(${(1 - e) * (effect === 'vortex' ? -270 : 720)}deg) scale(${Math.max(0.01, e)})`;
                newLayer.style.opacity = String(Math.min(1, e * 1.5));
            };
            break;
        case 'flash': {
            const white = overlay('#fff');
            render = (p) => {
                white.style.opacity = String(1 - Math.abs(p - 0.5) * 2);
                newLayer.style.opacity = p < 0.5 ? '0' : '1';
            };
            cleanup = () => white.remove();
            break;
        }
        case 'switch': case 'flip': case 'prism': case 'origami': case 'airplane': case 'peelOff': case 'pageCurlDouble': case 'pageCurlSingle': {
            const sign = direction === 'r' || direction === 'd' ? 1 : -1;
            const axis = direction === 'u' || direction === 'd' ? 'X' : 'Y';
            frame.style.perspective = `${Math.max(width, height) * 2}px`;
            render = (p) => {
                const e = inOut(p);
                if (effect === 'airplane' || effect === 'peelOff' || effect.startsWith('pageCurl')) {
                    oldLayer.style.zIndex = '2';
                    oldLayer.style.transformOrigin = '0% 100%';
                    oldLayer.style.transform = `translate(${-e * 60}%, ${-e * 40}%) rotate(${-e * 35}deg) scale(${1 - e * 0.6})`;
                    oldLayer.style.opacity = String(1 - e);
                    return;
                }
                oldLayer.style.backfaceVisibility = 'hidden';
                newLayer.style.backfaceVisibility = 'hidden';
                oldLayer.style.transform = `rotate${axis}(${sign * e * 180}deg)` + (effect === 'switch' ? ` translateZ(${-e * 100}px)` : '');
                newLayer.style.transform = `rotate${axis}(${sign * (e * 180 - 180)}deg)`;
            };
            cleanup = () => { frame.style.perspective = ''; };
            break;
        }
        case 'doors': case 'window': {
            // The old slide splits down the middle and both halves open away.
            const vertical = direction !== 'horz';
            const second = oldLayer.cloneNode(true);
            second.classList.add('bo-show-layer-copy');
            frame.appendChild(second);
            oldLayer.style.zIndex = '2';
            second.style.zIndex = '2';
            frame.style.perspective = `${Math.max(width, height) * 1.5}px`;
            render = (p) => {
                const e = inOut(p);
                if (vertical) {
                    oldLayer.style.clipPath = 'inset(0 50% 0 0)';
                    second.style.clipPath = 'inset(0 0 0 50%)';
                    oldLayer.style.transformOrigin = '0% 50%';
                    second.style.transformOrigin = '100% 50%';
                    oldLayer.style.transform = effect === 'doors' ? `rotateY(${e * 90}deg)` : `translateX(${-e * 50}%)`;
                    second.style.transform = effect === 'doors' ? `rotateY(${-e * 90}deg)` : `translateX(${e * 50}%)`;
                } else {
                    oldLayer.style.clipPath = 'inset(0 0 50% 0)';
                    second.style.clipPath = 'inset(50% 0 0 0)';
                    oldLayer.style.transformOrigin = '50% 0%';
                    second.style.transformOrigin = '50% 100%';
                    oldLayer.style.transform = effect === 'doors' ? `rotateX(${-e * 90}deg)` : `translateY(${-e * 50}%)`;
                    second.style.transform = effect === 'doors' ? `rotateX(${e * 90}deg)` : `translateY(${e * 50}%)`;
                }
                newLayer.style.transform = `scale(${lerp(0.85, 1, e)})`;
            };
            cleanup = () => { second.remove(); frame.style.perspective = ''; };
            break;
        }
        case 'fracture': case 'crush': case 'prestige': case 'shred': {
            oldLayer.style.zIndex = '2';
            render = (p) => {
                const e = inOut(p);
                oldLayer.style.transform = effect === 'prestige'
                    ? `perspective(${width * 2}px) rotateX(${e * 80}deg) translateY(${e * 30}%)`
                    : `scale(${1 - e * 0.7}) rotate(${(effect === 'crush' ? 25 : -15) * e}deg)`;
                oldLayer.style.opacity = String(1 - e);
                if (effect === 'shred') clip(newLayer, revealShape('dissolve', p, width, height, 31));
            };
            break;
        }
        case 'morph':
            ({ render, cleanup } = morph(oldStage, newStage, morphPairs));
            break;
        default: {
            // Filter-shaped reveals: random bars, shapes, wheel, wedge, blinds, checkerboard, comb, strips, dissolve…
            const filter = transitionFilter(transition);
            render = (p) => {
                const shape = revealShape(filter, p, width, height, 17);
                if (shape) clip(newLayer, shape);
                else newLayer.style.opacity = String(p);
            };
        }
    }

    const finish = () => {
        if (finished) return;
        finished = true;
        if (raf) clock.cancelFrame(raf);
        render(1);
        cleanup();
        reset();
        done();
    };

    const step = () => {
        raf = 0;
        if (finished) return;
        const p = Math.min(1, (clock.now() - startTime) / duration);
        render(p);
        if (p >= 1) finish();
        else raf = clock.frame(step);
    };
    render(0);
    raf = clock.frame(step);
    return finish;
}

/** The reveal filter a classic transition draws with. */
function transitionFilter(transition) {
    const direction = transition.direction || '';
    switch (transition.effect) {
        case 'randomBar': return direction === 'vert' ? 'randombar(vertical)' : 'randombar(horizontal)';
        case 'circle': case 'ripple': return 'circle(out)';
        case 'diamond': return 'diamond(out)';
        case 'plus': return 'plus(out)';
        case 'wedge': return 'wedge';
        case 'wheel': case 'wheelReverse': return `wheel(${transition.option || '1'})`;
        case 'blinds': case 'comb': return direction === 'vert' ? 'blinds(vertical)' : 'blinds(horizontal)';
        case 'checker': return direction === 'vert' ? 'checkerboard(down)' : 'checkerboard(across)';
        case 'strips': return `strips(${{ lu: 'upLeft', ru: 'upRight', ld: 'downLeft', rd: 'downRight' }[direction] || 'downLeft'})`;
        case 'dissolve': case 'honeycomb': case 'glitter': return 'dissolve';
        case 'random': return ['circle(out)', 'diamond(out)', 'wipe(right)', 'dissolve', 'randombar(horizontal)', 'wedge'][Math.floor(Math.random() * 6)];
        default: return 'fade';
    }
}

/**
 * Morph: each shape paired with one on the old slide glides from that shape's frame to its own
 * — position, size and rotation — while the two cross-fade, so a recolour or a changed word
 * blends too. Unpaired old shapes fade out over the first half, unpaired new ones fade in over
 * the second, and the new background fades in over the old.
 */
function morph(oldStage, newStage, pairs) {
    const partners = new Map((pairs || []).map(p => [p.to, p.from]));
    const newGroups = [...(newStage.svg?.querySelectorAll('[data-spid]') || [])];
    const oldGroups = [...(oldStage.svg?.querySelectorAll('[data-spid]') || [])];
    const matchedOld = new Set();
    const items = [];
    const box = (el) => {
        const b = (el.getAttribute('data-box') || '0 0 0 0').split(' ').map(Number);
        return { x: b[0], y: b[1], w: b[2], h: b[3], r: Number(el.getAttribute('data-rot') || 0) };
    };
    for (const element of newGroups) {
        if (element.getAttribute('visibility') === 'hidden') continue;
        const id = Number(element.getAttribute('data-spid'));
        const fromId = partners.get(id);
        const partner = fromId != null ? oldStage.svg.querySelector(`[data-spid="${fromId}"]`) : null;
        if (partner && partner.getAttribute('visibility') !== 'hidden') {
            matchedOld.add(partner);
            items.push({ element, partner, to: box(element), from: box(partner) });
        } else {
            items.push({ element, partner: null });
        }
    }
    const fadingOld = oldGroups.filter(g => !matchedOld.has(g));
    const background = newStage.svg?.querySelector('[data-bo-bg]');
    // The new slide's layer is see-through while its background fades in: its own white would
    // otherwise hide the old slide, whose shapes are gliding and fading beneath it.
    const newLayer = newStage.layer;
    if (newLayer) newLayer.style.background = 'transparent';

    const transform = (el, frame, target, e) => {
        // Draw `el` (whose own frame is `frame`) at the frame `e` of the way from `frame` to `target`.
        const cx = lerp(frame.x + frame.w / 2, target.x + target.w / 2, e);
        const cy = lerp(frame.y + frame.h / 2, target.y + target.h / 2, e);
        const sx = lerp(1, frame.w > 0 ? target.w / frame.w : 1, e);
        const sy = lerp(1, frame.h > 0 ? target.h / frame.h : 1, e);
        const rot = lerp(0, target.r - frame.r, e);
        const ox = frame.x + frame.w / 2, oy = frame.y + frame.h / 2;
        el.setAttribute('transform', `translate(${cx} ${cy}) rotate(${rot}) scale(${Math.max(0.0001, sx)} ${Math.max(0.0001, sy)}) translate(${-ox} ${-oy})`);
    };

    const render = (p) => {
        const e = inOut(p);
        if (background) background.setAttribute('opacity', String(e));
        for (const item of items) {
            if (item.partner) {
                // The new shape starts where the old one is and ends at its own frame; the old one the reverse.
                transform(item.element, item.to, item.from, 1 - e);
                transform(item.partner, item.from, item.to, e);
                item.element.setAttribute('opacity', String(smooth(0, 0.55, e)));
                item.partner.setAttribute('opacity', String(1 - smooth(0.45, 1, e)));
            } else {
                item.element.setAttribute('opacity', String(smooth(0.5, 1, e)));
            }
        }
        for (const element of fadingOld) element.setAttribute('opacity', String(1 - smooth(0, 0.5, e)));
    };
    const cleanup = () => {
        background?.removeAttribute('opacity');
        if (newLayer) newLayer.style.background = '';
        for (const item of items) {
            item.element.removeAttribute('transform');
            item.element.removeAttribute('opacity');
        }
    };
    return { render, cleanup };
}

// ---- The show -------------------------------------------------------------------------------------------

/** Pixels of the drawing in one millimetre: EMU per millimetre over EMU per pixel. */
const PX_PER_MM = 36000 / 9525;

/**
 * One stroke of show ink as SVG: a pen round-tipped and solid, a highlighter flat-tipped and see-through,
 * its thickness in millimetres of the slide (a highlighter's tip is as tall as its thickness and half as wide).
 */
function inkStrokeSvg(stroke) {
    const points = stroke.points;
    if (!points || points.length < 2) return '';
    let d = `M${points[0]} ${points[1]}`;
    for (let i = 2; i < points.length; i += 2) d += `L${points[i]} ${points[i + 1]}`;
    if (points.length === 2) d += `l0.01 0`;
    const highlighter = stroke.tool === 'highlighter';
    const width = Math.max(0.5, stroke.thicknessMm * PX_PER_MM * (highlighter ? 0.75 : 1));
    return `<path d="${d}" fill="none" stroke="#${stroke.colorHex}" stroke-width="${width}" stroke-linecap="${highlighter ? 'square' : 'round'}" stroke-linejoin="round"${highlighter ? ' opacity="0.5"' : ''}/>`;
}

/**
 * Where a recorded laser pointer is `t` milliseconds after its slide arrived: inside a gesture (time, x, y triples)
 * between two of its points, interpolated; null between gestures.
 */
function laserPointAt(traces, t) {
    for (const trace of traces || []) {
        const n = trace.length / 3;
        if (n === 0 || t < trace[0] || t > trace[(n - 1) * 3]) continue;
        for (let i = 0; i < n; i++) {
            if (trace[i * 3] < t && i + 1 < n && trace[(i + 1) * 3] < t) continue;
            if (i + 1 >= n || trace[i * 3] >= t) return { x: trace[i * 3 + 1], y: trace[i * 3 + 2] };
            const span = trace[(i + 1) * 3] - trace[i * 3];
            const f = span > 0 ? (t - trace[i * 3]) / span : 1;
            return { x: lerp(trace[i * 3 + 1], trace[(i + 1) * 3 + 1], f), y: lerp(trace[i * 3 + 2], trace[(i + 1) * 3 + 2], f) };
        }
    }
    return null;
}

/** The pens the show's ink menu offers: PowerPoint's show pen (red), its highlighter (yellow), with their thickness in millimetres. */
const SHOW_PENS = Object.freeze({
    pen: { tool: 'pen', colorHex: 'E71224', thicknessMm: 1 },
    highlighter: { tool: 'highlighter', colorHex: 'FFFF00', thicknessMm: 6 },
});

/** The ink colours of the show's menu: PowerPoint's palette, less white. */
const SHOW_INK_COLORS = ['000000', 'E71224', 'F6630D', 'FFC114', '66CC00', '008C3A', '004F8B', '33CCFF', '5B2D90', 'E71CE0', 'FFFF00', '00FFFF'];

/**
 * One surface the show draws on — the window's frame, the presenter's current-slide pane or the
 * audience window — with its own layers, stage and transition. The show sends the same commands
 * to every screen at the same moment, so they play in step; only one of them is heard.
 */
class ShowScreen {
    constructor(show, frame, muted) {
        this.show = show;
        this.frame = frame;
        this.muted = muted;
        this.layer = null;
        this.stage = null;
        this.finishTransition = null;
    }

    // What a SlideStage reads from its owner.
    get clock() { return this.show.clock; }
    get mediaLog() { return this.show.mediaLog; }
    get mediaInfo() { return this.show.mediaInfo; }
    get width() { return this.show.width; }
    get height() { return this.show.height; }
    get options() { return this.show.options; }

    /** Shows a slide, through its transition when given one; `arrived(stage)` runs when it is on screen. */
    present(slide, transition, built, arrived) {
        this.finishTransition?.();
        const oldLayer = this.layer;
        const oldStage = this.stage;
        oldStage?.finish(true);
        for (const player of oldStage?.stopMedia(true) || []) this.carry(player);

        const layer = this.frame.ownerDocument.createElement('div');
        layer.className = 'bo-show-layer';
        this.frame.appendChild(layer);
        const stage = new SlideStage(layer, slide, this);
        stage.onIdle = () => this.show.maybeAdvance();
        this.layer = layer;
        this.stage = stage;
        if (built) stage.applySteps(stage.stepCount);

        let arrivedAlready = false;
        const finish = playTransition(this.frame, oldLayer, layer, transition, oldStage, stage, slide.morphPairs, () => {
            arrivedAlready = true;
            this.finishTransition = null;
            if (oldLayer) { oldStage?.dispose(); oldLayer.remove(); }
            arrived?.(stage);
        }, this.clock);
        // A transition with nothing to play arrives at once, before playTransition returns.
        this.finishTransition = arrivedAlready ? null : finish;
    }

    /**
     * A sound or film playing across slides, handed to the show; a film moves into this screen's carry
     * layer — over the slides, at its own place — and goes on there until it ends or the show does.
     */
    carry(player) {
        const film = filmElement(player);
        if (film) {
            if (!this.carryLayer) {
                this.carryLayer = this.frame.ownerDocument.createElement('div');
                this.carryLayer.className = 'bo-show-carry';
                this.frame.appendChild(this.carryLayer);
            }
            this.carryLayer.appendChild(film);
        }
        this.show.carry?.(player, film);
    }

    /** Redraws the current slide with its first `steps` steps played. */
    rebuild(steps) {
        if (!this.stage) return;
        const slide = this.stage.slide;
        this.stage.dispose();
        this.layer.innerHTML = '';
        const stage = new SlideStage(this.layer, slide, this);
        stage.onIdle = () => this.show.maybeAdvance();
        stage.applySteps(steps);
        this.stage = stage;
    }

    /**
     * The ink drawn on this slide during the show, and the laser pointer's dot, in a layer of their own over
     * the slides (and over a film carried across them) — redrawn whole on each change, on every screen.
     */
    paintInk(strokes, laser, laserColor = 'FF2A2A') {
        const doc = this.frame.ownerDocument;
        if (!this.inkLayer) {
            this.inkLayer = doc.createElementNS(NS, 'svg');
            this.inkLayer.setAttribute('class', 'bo-show-ink');
            this.inkLayer.setAttribute('viewBox', `0 0 ${this.width} ${this.height}`);
            this.inkLayer.setAttribute('preserveAspectRatio', 'none');
        }
        // Each slide's layer is appended as it arrives: the ink stays the frame's last child, over them all.
        if (this.frame.lastChild !== this.inkLayer) this.frame.appendChild(this.inkLayer);
        let markup = '';
        for (const stroke of strokes || []) markup += inkStrokeSvg(stroke);
        if (laser) {
            const r = Math.max(4, this.width / 160);
            markup += `<circle cx="${laser.x}" cy="${laser.y}" r="${r * 1.8}" fill="#${laserColor}" opacity="0.25"/><circle cx="${laser.x}" cy="${laser.y}" r="${r}" fill="#${laserColor}" stroke="#fff" stroke-width="${r / 4}"/>`;
        }
        this.inkLayer.innerHTML = markup;
    }

    /** Hears or silences the slide's media (the presenter's screen goes quiet when an audience window opens). */
    setMuted(muted) {
        this.muted = muted;
        for (const { player } of this.stage?.media.values() || []) player.muted = muted;
    }

    dispose() {
        this.finishTransition?.();
        this.stage?.dispose();
    }
}

/**
 * Rehearse Timings: PowerPoint's Recording bar — Next, Pause, the slide's time, Repeat, the total —
 * and the record of how long each click and slide took. Each slide's record is the wait before each
 * click and the wait after the last, in milliseconds: what PowerPoint keeps in the slide's TIMING tag,
 * whose sum becomes the slide's "After" time.
 */
class ShowRehearsal {
    constructor(show, record = false) {
        this.show = show;
        this.chunks = [];
        this.capturing = false;
        this.records = new Map();
        this.totalBefore = 0;
        this.slide = null;
        this.pausedAt = null;
        this.paused = 0;

        const bar = document.createElement('div');
        bar.className = 'bo-show-rehearse';
        bar.addEventListener('click', (event) => event.stopPropagation());
        const button = (icon, title, action) => {
            const element = document.createElement('button');
            element.type = 'button';
            element.title = title;
            element.innerHTML = `<i class="rzi">${icon}</i>`;
            element.addEventListener('click', (event) => { event.stopPropagation(); action(); });
            bar.appendChild(element);
            return element;
        };
        const label = document.createElement('span');
        label.className = 'bo-show-rehearse-label';
        label.textContent = record ? 'Recording narration' : 'Recording';
        bar.appendChild(label);
        button('arrow_forward', 'Next (N, →, space)', () => show.next());
        this.pauseButton = button('pause', 'Pause recording', () => this.togglePause());
        this.slideClock = document.createElement('span');
        this.slideClock.className = 'bo-show-rehearse-slide';
        bar.appendChild(this.slideClock);
        button('replay', 'Repeat: time this slide again', () => this.repeat());
        this.totalClock = document.createElement('span');
        this.totalClock.className = 'bo-show-rehearse-total';
        bar.appendChild(this.totalClock);
        show.root.appendChild(bar);
        this.bar = bar;
        this.ticker = setInterval(() => this.paint(), 250);
        this.recorderReady = record ? this.startRecorder() : Promise.resolve();
    }

    /**
     * Record: the microphone, captured as 16-bit PCM for each slide while it is on screen (pauses
     * left out) and kept as a WAV file — a format PowerPoint plays, unlike the browser's own
     * recorder's WebM. Without a microphone (refused, none) only the timings are recorded.
     */
    async startRecorder() {
        try {
            this.stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true } });
        } catch {
            this.show.flash('The microphone is not available, so only the timings are recorded.');
            return;
        }
        this.audio = new AudioContext();
        const source = this.audio.createMediaStreamSource(this.stream);
        // A script processor needs no module file, which a page of any host can run.
        const processor = this.audio.createScriptProcessor(4096, 1, 1);
        processor.onaudioprocess = (event) => {
            if (this.capturing && this.pausedAt === null) this.chunks.push(new Float32Array(event.inputBuffer.getChannelData(0)));
        };
        source.connect(processor);
        processor.connect(this.audio.destination); // silent: its output buffer stays empty
        this.processor = processor;
        if (this.slide) this.capturing = true;
    }

    /** The captured samples as a mono 16-bit WAV file; null when nothing worth keeping was captured. */
    takeWav() {
        const chunks = this.chunks;
        this.chunks = [];
        if (!this.audio || chunks.length === 0) return null;
        const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
        const rate = this.audio.sampleRate;
        if (length < rate * 0.3) return null;
        const buffer = new ArrayBuffer(44 + length * 2);
        const view = new DataView(buffer);
        const text = (offset, value) => { for (let i = 0; i < value.length; i++) view.setUint8(offset + i, value.charCodeAt(i)); };
        text(0, 'RIFF'); view.setUint32(4, 36 + length * 2, true); text(8, 'WAVE');
        text(12, 'fmt '); view.setUint32(16, 16, true); view.setUint16(20, 1, true); view.setUint16(22, 1, true);
        view.setUint32(24, rate, true); view.setUint32(28, rate * 2, true); view.setUint16(32, 2, true); view.setUint16(34, 16, true);
        text(36, 'data'); view.setUint32(40, length * 2, true);
        let offset = 44;
        for (const chunk of chunks) {
            for (let i = 0; i < chunk.length; i++, offset += 2) {
                const sample = Math.max(-1, Math.min(1, chunk[i]));
                view.setInt16(offset, sample < 0 ? sample * 0x8000 : sample * 0x7fff, true);
            }
        }
        return new Uint8Array(buffer);
    }

    /** The time spent on the current slide, pauses left out. */
    elapsed() {
        if (!this.slide) return 0;
        const end = this.pausedAt ?? this.show.clock.now();
        return Math.max(0, end - this.start - this.paused);
    }

    /** A slide has arrived: the record of the one before it is closed; the same slide again starts its record over. */
    begin(slide) {
        if (this.slide && this.slide.deckIndex !== slide.deckIndex) this.finishSlide();
        this.slide = slide;
        this.start = this.show.clock.now();
        this.paused = 0;
        if (this.pausedAt !== null) this.pausedAt = this.show.clock.now();
        this.mark = 0;
        this.waits = [];
        this.chunks = [];
        this.capturing = !!this.audio;
        this.paint();
    }

    /** A click played: the wait before it is kept. */
    click() {
        if (!this.slide) return;
        const at = this.elapsed();
        this.waits.push(Math.round(at - this.mark));
        this.mark = at;
    }

    /** One click undone (Previous): its wait goes too. */
    restart(slide) {
        if (!this.slide || this.slide.deckIndex !== slide.deckIndex || this.waits.length === 0) return;
        this.mark -= this.waits.pop();
    }

    /** The slide is left: the wait after its last click closes its record. */
    finishSlide() {
        if (!this.slide) return;
        const at = this.elapsed();
        this.waits.push(Math.round(at - this.mark));
        this.capturing = false;
        this.records.set(this.slide.deckIndex, { deckIndex: this.slide.deckIndex, timingsMs: this.waits, wav: this.takeWav() });
        this.totalBefore += at;
        this.slide = null;
    }

    repeat() {
        if (!this.slide) return;
        // Forget this slide's record so far; its arrival starts a new one.
        this.slide = null;
        this.show.goTo(this.show.position, false);
    }

    togglePause() {
        if (this.pausedAt === null) {
            this.pausedAt = this.show.clock.now();
            this.pauseButton.innerHTML = '<i class="rzi">play_arrow</i>';
            this.pauseButton.title = 'Resume recording';
        } else {
            this.paused += this.show.clock.now() - this.pausedAt;
            this.pausedAt = null;
            this.pauseButton.innerHTML = '<i class="rzi">pause</i>';
            this.pauseButton.title = 'Pause recording';
        }
    }

    paint() {
        const format = (ms) => {
            const seconds = Math.floor(ms / 1000);
            return `${Math.floor(seconds / 3600)}:${String(Math.floor(seconds / 60) % 60).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
        };
        const slide = this.elapsed();
        this.slideClock.textContent = format(slide);
        this.totalClock.textContent = format(this.totalBefore + slide);
    }

    /** Every slide's record, in the order the slides were left. */
    results() { return [...this.records.values()]; }

    dispose() {
        clearInterval(this.ticker);
        this.bar.remove();
        this.capturing = false;
        try { this.processor?.disconnect(); } catch { /* gone */ }
        for (const track of this.stream?.getTracks() || []) track.stop();
        this.audio?.close().catch(() => { });
    }
}

class SlideShow {
    constructor(host, options, dotNetRef) {
        this.clock = realClock;
        this.options = options || {};
        this.dotNetRef = dotNetRef;
        this.width = this.options.width || 1280;
        this.height = this.options.height || 720;
        this.count = this.options.slideCount || 0;
        // Slides the show reaches only by a zoom or a link (hidden, outside the range): positions count and up.
        this.extraCount = this.options.extraCount || 0;
        this.position = -1;
        this.cache = new Map();
        this.ended = false;
        this.disposed = false;
        this.typed = '';
        this.timer = 0;
        this.timeElapsed = false;
        this.blanks = [];
        this.audience = null;
        // Ink drawn during the show, by deck index: strokes in the drawing's pixels, as PowerPoint keeps
        // annotations on their slide for the rest of the show. The tool is null (the arrow), 'laser',
        // 'pen', 'highlighter' or 'eraser'.
        this.ink = new Map();
        this.inkTool = null;
        this.inkHidden = false;
        this.inkPens = { pen: { ...SHOW_PENS.pen }, highlighter: { ...SHOW_PENS.highlighter } };
        this.build();
    }

    // The presenter's own screen decides: what is on it is what the show is doing.
    get stage() { return this.screen.stage; }
    get finishTransition() { return this.screen.finishTransition; }
    get screens() { return this.audience ? [this.screen, this.audience] : [this.screen]; }

    build() {
        const host = document.createElement('div');
        host.className = 'bo-show bo-show-player';
        host.tabIndex = 0;
        document.body.appendChild(host);
        this.root = host;

        this.frame = document.createElement('div');
        this.frame.className = 'bo-show-frame';
        const aspect = this.width / this.height;
        if (this.options.presenter) {
            host.classList.add('bo-show-presenter-mode');
            this.buildPresenter(aspect);
        } else {
            this.frame.style.width = `min(100vw, calc(100vh * ${aspect}))`;
            this.frame.style.height = `min(100vh, calc(100vw / ${aspect}))`;
            host.appendChild(this.frame);
        }
        this.screen = new ShowScreen(this, this.frame, false);

        this.endScreen = document.createElement('div');
        this.endScreen.className = 'bo-show-end';
        this.endScreen.textContent = 'End of slide show. Click to exit.';
        this.endScreen.hidden = true;
        (this.options.presenter ? this.frame : host).appendChild(this.endScreen);

        this.controls = document.createElement('div');
        this.controls.className = 'bo-show-controls';
        this.button(this.controls, 'arrow_back', 'Previous (P, ←)', () => this.previous());
        this.button(this.controls, 'arrow_forward', 'Next (N, →, space)', () => this.next());
        this.button(this.controls, 'apps', 'See all slides (G)', () => this.toggleList());
        this.button(this.controls, 'contrast', 'Black screen (B)', () => this.toggleBlank('#000'));
        this.inkButton = this.button(this.controls, 'stylus', 'Pen and laser pointer tools (Ctrl+L laser, Ctrl+P pen, Ctrl+I highlighter, Ctrl+E eraser, E erase all, Ctrl+A arrow)', () => this.toggleInkMenu());
        this.button(this.controls, 'tune', 'Show options: timings, narrations, media controls', () => this.toggleSettings());
        this.button(this.controls, 'close', 'End show (Esc)', () => this.exit());
        this.counter = document.createElement('span');
        this.counter.className = 'bo-show-counter';
        this.controls.appendChild(this.counter);
        if (!this.options.presenter) host.appendChild(this.controls);

        this.list = document.createElement('div');
        this.list.className = 'bo-show-list';
        this.list.hidden = true;
        (this.options.titles || []).forEach((title, index) => {
            const item = document.createElement('button');
            item.type = 'button';
            item.className = 'bo-show-list-item';
            item.textContent = `${index + 1}. ${title || 'Slide ' + (index + 1)}`;
            item.addEventListener('click', (event) => { event.stopPropagation(); this.list.hidden = true; this.goTo(index, false); });
            this.list.appendChild(item);
        });
        host.appendChild(this.list);

        this.onKey = (event) => this.key(event);
        this.onClick = (event) => {
            if (event.button !== 0) return;
            if (!this.list.hidden) { this.list.hidden = true; return; }
            if (this.inkMenu) { this.closeInkMenu(); return; }
            // With a pen, a highlighter or the eraser a click draws; it never moves the show on (the laser does).
            if (this.inkTool && this.inkTool !== 'laser') return;
            // In presenter view only the slide itself advances; the panes around it are controls.
            if (this.options.presenter && !event.target?.closest?.('.bo-show-frame')) return;
            if (this.followZoom(event)) return;
            if (this.followAction(event)) return;
            const trigger = event.target?.closest?.('[data-spid]');
            if (trigger && this.playTrigger(Number(trigger.getAttribute('data-spid')))) return;
            if (this.options.kiosk) return;
            this.next(true);
        };
        this.onContext = (event) => { event.preventDefault(); if (!this.options.kiosk) this.previous(); };
        this.onWheel = (event) => {
            if (this.options.kiosk || (this.options.presenter && !event.target?.closest?.('.bo-show-frame'))) return;
            if (event.deltaY > 0) this.next(); else if (event.deltaY < 0) this.previous();
        };
        this.onMove = () => {
            this.controls.classList.add('bo-show-controls-visible');
            clearTimeout(this.hideControls);
            this.hideControls = setTimeout(() => this.controls.classList.remove('bo-show-controls-visible'), 2500);
        };
        host.addEventListener('keydown', this.onKey);
        host.addEventListener('click', this.onClick);
        // Mouse-over actions (a:hlinkHover): pointing at the shape follows its link once, as PowerPoint does.
        this.onHover = (event) => {
            const group = event.target?.closest?.('[data-spid]');
            const id = group?.getAttribute('data-spid') ?? null;
            if (id === this.hovered) return;
            this.hovered = id;
            const slide = this.stage?.slide;
            const url = id != null ? slide?.hoverActions?.[id] : null;
            if (!url || this.ended) return;
            if (url === '#action:ppaction://media') { for (const screen of this.screens) screen.stage?.command(Number(id), 'togglePause'); return; }
            this.follow(url, slide);
        };
        host.addEventListener('mouseover', this.onHover);
        host.addEventListener('contextmenu', this.onContext);
        host.addEventListener('wheel', this.onWheel, { passive: true });
        host.addEventListener('mousemove', this.onMove);
        this.listenForInk(host, () => this.frame);
    }

    // ---- Ink and the laser pointer ----------------------------------------------------------------------

    /**
     * Pointer handling for the ink tools on one window's root: a pen or highlighter draws a stroke while the
     * button is down, the eraser takes away every stroke it passes over, the laser pointer follows the pointer.
     * `frameOf()` is the slide's frame on that window, which the pointer is measured against.
     */
    listenForInk(root, frameOf) {
        const toSlide = (event) => {
            const box = frameOf().getBoundingClientRect();
            if (box.width <= 0 || box.height <= 0) return null;
            const x = (event.clientX - box.left) / box.width * this.width;
            const y = (event.clientY - box.top) / box.height * this.height;
            return x < 0 || y < 0 || x > this.width || y > this.height ? null : { x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10 };
        };
        root.addEventListener('pointerdown', (event) => {
            if (event.button !== 0 || !this.inkTool || this.inkTool === 'laser' || !this.stage || this.ended) return;
            if (event.target?.closest?.('.bo-show-controls, .bo-show-ink-menu, .bo-show-settings, .bo-show-list, .bo-show-presenter-bar, .bo-show-presenter-side')) return;
            const point = toSlide(event);
            if (!point) return;
            event.preventDefault();
            try { event.target.setPointerCapture?.(event.pointerId); } catch { /* not capturable */ }
            if (this.inkTool === 'eraser') { this.erasing = true; this.eraseAt(point); return; }
            const pen = this.inkPens[this.inkTool];
            this.drawing = { tool: pen.tool, colorHex: pen.colorHex, thicknessMm: pen.thicknessMm, points: [point.x, point.y] };
            this.slideInk(true).push(this.drawing);
            this.paintInk();
        });
        root.addEventListener('pointermove', (event) => {
            const point = toSlide(event);
            if (this.inkTool === 'laser') { this.laser = point; this.paintInk(); return; }
            if (!point) return;
            if (this.erasing) { this.eraseAt(point); return; }
            if (!this.drawing) return;
            const n = this.drawing.points.length;
            if (Math.hypot(point.x - this.drawing.points[n - 2], point.y - this.drawing.points[n - 1]) < 0.75) return;
            this.drawing.points.push(point.x, point.y);
            this.paintInk();
        });
        const end = () => {
            if (this.drawing) this.report('SlideShowInkStroke', `slide ${(this.stage?.slide.deckIndex ?? -1) + 1}`);
            this.drawing = null;
            this.erasing = false;
        };
        root.addEventListener('pointerup', end);
        root.addEventListener('pointercancel', end);
        root.addEventListener('pointerleave', () => { if (this.inkTool === 'laser') { this.laser = null; this.paintInk(); } });
    }

    /** The strokes drawn on the slide showing; `create` makes the list when there is none yet. */
    slideInk(create = false) {
        const key = this.stage?.slide?.deckIndex;
        if (key == null) return [];
        if (!this.ink.has(key) && create) this.ink.set(key, []);
        return this.ink.get(key) || [];
    }

    /** Draws the slide's ink and the laser's dot on every screen (Ctrl+M hides the ink, not the laser). */
    paintInk() {
        const strokes = this.inkHidden ? [] : this.slideInk();
        const laser = this.inkTool === 'laser' && this.laser ? this.laser : this.recordedLaser;
        const color = this.inkTool === 'laser' && this.laser ? 'FF2A2A' : (this.options.laserColorHex || 'FF0000');
        for (const screen of this.screens) screen.paintInk(strokes, laser, color);
    }

    /**
     * The laser-pointer gestures a recording kept with the slide, played as the slide plays by its timings: the
     * dot follows each gesture at its times from the slide's arrival, on the show's clock.
     */
    playRecordedLaser(slide) {
        this.clock.cancelFrame(this.laserFrame);
        this.recordedLaser = null;
        if (!slide?.laserTraces?.length || this.options.useTimings === false || slide.transition?.advanceAfterMs == null || this.options.rehearse) return;
        const arrived = this.clock.now();
        const end = Math.max(...slide.laserTraces.map(t => t[t.length - 3] ?? 0));
        const tick = () => {
            if (this.disposed || this.stage?.slide !== slide) return;
            const t = this.clock.now() - arrived;
            const point = laserPointAt(slide.laserTraces, t);
            if ((point?.x ?? null) !== (this.recordedLaser?.x ?? null) || (point?.y ?? null) !== (this.recordedLaser?.y ?? null)) {
                this.recordedLaser = point;
                this.paintInk();
            }
            if (t <= end) this.laserFrame = this.clock.frame(tick);
            else if (this.recordedLaser) { this.recordedLaser = null; this.paintInk(); }
        };
        this.laserFrame = this.clock.frame(tick);
    }

    /** The eraser: every stroke passing within a few pixels of the point goes. */
    eraseAt(point) {
        const strokes = this.slideInk();
        const reach = Math.max(6, this.width / 200);
        const kept = strokes.filter(stroke => {
            const half = stroke.thicknessMm * PX_PER_MM / 2;
            const p = stroke.points;
            for (let i = 0; i + 1 < p.length; i += 2) {
                const ax = p[i], ay = p[i + 1], bx = p[i + 2] ?? ax, by = p[i + 3] ?? ay;
                const dx = bx - ax, dy = by - ay;
                const t = dx || dy ? Math.max(0, Math.min(1, ((point.x - ax) * dx + (point.y - ay) * dy) / (dx * dx + dy * dy))) : 0;
                if (Math.hypot(point.x - (ax + t * dx), point.y - (ay + t * dy)) <= reach + half) return false;
            }
            return true;
        });
        if (kept.length === strokes.length) return;
        this.ink.set(this.stage.slide.deckIndex, kept);
        this.paintInk();
    }

    /** E: erases all the ink on the slide. */
    eraseSlideInk() {
        if (!this.slideInk().length) return;
        this.ink.set(this.stage.slide.deckIndex, []);
        this.paintInk();
    }

    /** Picks an ink tool (null is the arrow): the pointer changes, and a pen stops clicks moving the show on. */
    setInkTool(tool) {
        this.inkTool = tool;
        this.laser = null;
        for (const root of [this.root, this.audienceRoot].filter(Boolean)) {
            root.classList.toggle('bo-show-inking', !!tool && tool !== 'laser');
            root.classList.toggle('bo-show-lasering', tool === 'laser');
            root.classList.toggle('bo-show-erasing', tool === 'eraser');
        }
        this.inkButton?.classList.toggle('bo-show-control-on', !!tool);
        this.paintInk();
    }

    /** The control bar's pen menu: Laser Pointer, Pen, Highlighter, the ink colour, Eraser, Erase All Ink on Slide, Show/Hide Ink Markup. */
    toggleInkMenu() {
        if (this.inkMenu) { this.closeInkMenu(); return; }
        this.settingsMenu?.remove();
        this.settingsMenu = null;
        const menu = document.createElement('div');
        menu.className = 'bo-show-ink-menu';
        const item = (icon, label, on, action) => {
            const row = document.createElement('button');
            row.type = 'button';
            row.className = 'bo-show-ink-item' + (on ? ' bo-show-ink-item-on' : '');
            row.innerHTML = `<i class="rzi">${icon}</i><span>${label}</span>`;
            row.addEventListener('click', (event) => { event.stopPropagation(); action(); this.closeInkMenu(); });
            menu.appendChild(row);
        };
        item('arrow_selector_tool', 'Arrow (Ctrl+A)', !this.inkTool, () => this.setInkTool(null));
        item('highlight_mouse_cursor', 'Laser Pointer (Ctrl+L)', this.inkTool === 'laser', () => this.setInkTool('laser'));
        item('stylus', 'Pen (Ctrl+P)', this.inkTool === 'pen', () => this.setInkTool('pen'));
        item('ink_highlighter', 'Highlighter (Ctrl+I)', this.inkTool === 'highlighter', () => this.setInkTool('highlighter'));
        const colors = document.createElement('div');
        colors.className = 'bo-show-ink-colors';
        const current = this.inkTool === 'highlighter' ? this.inkPens.highlighter : this.inkPens.pen;
        for (const hex of SHOW_INK_COLORS) {
            const swatch = document.createElement('button');
            swatch.type = 'button';
            swatch.className = 'bo-show-ink-swatch' + (current.colorHex === hex ? ' bo-show-ink-swatch-on' : '');
            swatch.style.background = `#${hex}`;
            swatch.title = `Ink colour #${hex}`;
            swatch.addEventListener('click', (event) => {
                event.stopPropagation();
                const tool = this.inkTool === 'highlighter' ? 'highlighter' : 'pen';
                this.inkPens[tool].colorHex = hex;
                this.setInkTool(tool);
                this.closeInkMenu();
            });
            colors.appendChild(swatch);
        }
        menu.appendChild(colors);
        item('ink_eraser', 'Eraser (Ctrl+E)', this.inkTool === 'eraser', () => this.setInkTool('eraser'));
        item('delete_sweep', 'Erase All Ink on Slide (E)', false, () => this.eraseSlideInk());
        item(this.inkHidden ? 'visibility' : 'visibility_off', this.inkHidden ? 'Show Ink Markup (Ctrl+M)' : 'Hide Ink Markup (Ctrl+M)', false, () => { this.inkHidden = !this.inkHidden; this.paintInk(); });
        menu.addEventListener('click', (event) => event.stopPropagation());
        this.root.appendChild(menu);
        this.inkMenu = menu;
    }

    closeInkMenu() {
        this.inkMenu?.remove();
        this.inkMenu = null;
    }

    /** The ink keys: true when the key was one of them. */
    inkKey(event) {
        const key = event.key.toLowerCase();
        if (event.ctrlKey && !event.altKey) {
            const tool = { l: 'laser', p: 'pen', i: 'highlighter', e: 'eraser', a: null }[key];
            if (tool !== undefined) { event.preventDefault(); this.setInkTool(tool === this.inkTool && tool ? null : tool); return true; }
            if (key === 'm') { event.preventDefault(); this.inkHidden = !this.inkHidden; this.paintInk(); return true; }
            return false;
        }
        if (key === 'e' && !event.altKey && !event.metaKey) { event.preventDefault(); this.eraseSlideInk(); return true; }
        // Esc puts the pen down before it ends the show, as in PowerPoint.
        if (event.key === 'Escape' && this.inkTool) { event.preventDefault(); this.setInkTool(null); return true; }
        return false;
    }

    /** True when ink was drawn on any slide: the show then asks whether to keep it. */
    hasInk() {
        for (const strokes of this.ink.values()) if (strokes.some(s => s.points.length >= 2)) return true;
        return false;
    }

    button(container, icon, title, action, label) {
        const element = container.ownerDocument.createElement('button');
        element.type = 'button';
        element.title = title;
        element.className = label ? 'bo-show-presenter-button' : 'bo-show-control';
        // Radzen's icon font, which every host already loads for the editors' toolbars.
        element.innerHTML = `<i class="rzi">${icon}</i>${label ? `<span>${label}</span>` : ''}`;
        element.addEventListener('click', (event) => { event.stopPropagation(); action(); });
        container.appendChild(element);
        return element;
    }

    /**
     * Presenter view: the current slide large, the next one small, the speaker notes, a timer and
     * the controls — and a button that opens the audience window for the projector.
     */
    buildPresenter(aspect) {
        const view = document.createElement('div');
        view.className = 'bo-show-presenter';

        const current = document.createElement('div');
        current.className = 'bo-show-presenter-current';
        this.frame.style.aspectRatio = `${this.width} / ${this.height}`;
        current.appendChild(this.frame);
        view.appendChild(current);

        const side = document.createElement('div');
        side.className = 'bo-show-presenter-side';
        side.innerHTML = '<div class="bo-show-presenter-label">Next</div>';
        this.nextPreview = document.createElement('div');
        this.nextPreview.className = 'bo-show-presenter-next';
        this.nextPreview.style.aspectRatio = `${this.width} / ${this.height}`;
        side.appendChild(this.nextPreview);
        const notesLabel = document.createElement('div');
        notesLabel.className = 'bo-show-presenter-label';
        notesLabel.textContent = 'Notes';
        side.appendChild(notesLabel);
        this.notes = document.createElement('div');
        this.notes.className = 'bo-show-presenter-notes';
        side.appendChild(this.notes);
        view.appendChild(side);

        const bar = document.createElement('div');
        bar.className = 'bo-show-presenter-bar';
        this.wallClock = document.createElement('span');
        this.wallClock.className = 'bo-show-presenter-clock';
        this.elapsed = document.createElement('span');
        this.elapsed.className = 'bo-show-presenter-elapsed';
        bar.appendChild(this.elapsed);
        this.pauseButton = this.button(bar, 'pause', 'Pause the timer', () => this.toggleTimerPause(), 'Pause');
        this.button(bar, 'restart_alt', 'Restart the timer', () => this.resetTimer(), 'Restart');
        this.inkButton = this.button(bar, 'stylus', 'Pen and laser pointer tools', () => this.toggleInkMenu(), 'Pen');
        this.presenterBar = bar;
        bar.appendChild(this.wallClock);
        this.presenterCounter = document.createElement('span');
        this.presenterCounter.className = 'bo-show-presenter-counter';
        bar.appendChild(this.presenterCounter);
        const spacer = document.createElement('span');
        spacer.style.flex = '1';
        bar.appendChild(spacer);
        this.button(bar, 'arrow_back', 'Previous (P, ←)', () => this.previous(), 'Previous');
        this.button(bar, 'arrow_forward', 'Next (N, →, space)', () => this.next(), 'Next');
        this.button(bar, 'apps', 'See all slides (G)', () => this.toggleList(), 'All slides');
        this.button(bar, 'contrast', 'Black screen (B)', () => this.toggleBlank('#000'), 'Black');
        this.audienceButton = this.button(bar, 'cast', 'Open the audience window on the other screen', () => this.toggleAudience(), 'Audience window');
        this.button(bar, 'close', 'End show (Esc)', () => this.exit(), 'End');
        view.appendChild(bar);
        this.root.appendChild(view);

        this.timerStart = this.clock.now();
        this.timerPausedAt = null;
        const tick = () => {
            if (this.disposed) return;
            const elapsed = Math.floor(((this.timerPausedAt ?? this.clock.now()) - this.timerStart) / 1000);
            const h = Math.floor(elapsed / 3600), m = Math.floor(elapsed / 60) % 60, s = elapsed % 60;
            this.elapsed.textContent = `${h > 0 ? h + ':' : ''}${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            this.wallClock.textContent = new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
        };
        tick();
        this.clockTimer = setInterval(tick, 500);
    }

    toggleTimerPause() {
        if (this.timerPausedAt === null) {
            this.timerPausedAt = this.clock.now();
            this.pauseButton.innerHTML = '<i class="rzi">play_arrow</i><span>Resume</span>';
        } else {
            this.timerStart += this.clock.now() - this.timerPausedAt;
            this.timerPausedAt = null;
            this.pauseButton.innerHTML = '<i class="rzi">pause</i><span>Pause</span>';
        }
    }

    resetTimer() {
        this.timerStart = this.clock.now();
        if (this.timerPausedAt !== null) this.timerPausedAt = this.clock.now();
    }

    /** The presenter's panes for the slide now showing: the notes, the counter, the next slide. */
    async updatePresenter(slide) {
        if (!this.options.presenter) return;
        this.notes.textContent = slide.notes || '';
        this.notes.classList.toggle('bo-show-presenter-empty', !slide.notes);
        if (!slide.notes) this.notes.textContent = 'No notes for this slide.';
        this.presenterCounter.textContent = this.position >= this.count
            ? `Hidden slide ${slide.deckIndex + 1} (${this.count} in the show)`
            : `Slide ${this.position + 1} of ${this.count}`;
        const upcoming = this.upcoming(slide);
        const next = upcoming == null ? null : await this.load(upcoming);
        if (this.disposed) return;
        this.nextPreview.innerHTML = next ? next.svg : '<div class="bo-show-presenter-end">End of slide show</div>';
        const svg = this.nextPreview.querySelector('svg');
        if (svg) { svg.setAttribute('width', '100%'); svg.setAttribute('height', '100%'); }
    }

    // ---- The audience window ------------------------------------------------------------------------

    /**
     * Opens (or closes) the audience window: a pop-up of this page's origin the show draws into
     * directly, so it plays the same steps at the same moment — the presenter drags it to the
     * projector and clicks it to go full screen. Sound moves to it.
     */
    toggleAudience() {
        if (this.audience) { this.audienceWindow?.close(); return; }
        const popup = window.open('', 'bo-show-audience', `popup,width=${Math.round(this.width)},height=${Math.round(this.height)}`);
        if (!popup) { this.flash('Allow pop-ups for this site to open the audience window.'); return; }
        this.audienceWindow = popup;
        const doc = popup.document;
        doc.title = 'Slide show';
        doc.head.innerHTML = '';
        for (const sheet of document.querySelectorAll('link[rel="stylesheet"], style')) {
            const copy = sheet.cloneNode(true);
            if (sheet.href) copy.href = sheet.href;
            doc.head.appendChild(copy);
        }
        doc.body.innerHTML = '';
        doc.body.style.cssText = 'margin:0;background:#000;overflow:hidden;';
        const root = doc.createElement('div');
        root.className = 'bo-show bo-show-player';
        root.tabIndex = 0;
        const frame = doc.createElement('div');
        frame.className = 'bo-show-frame';
        const aspect = this.width / this.height;
        frame.style.width = `min(100vw, calc(100vh * ${aspect}))`;
        frame.style.height = `min(100vh, calc(100vw / ${aspect}))`;
        root.appendChild(frame);
        doc.body.appendChild(root);
        this.audienceRoot = root;

        this.audience = new ShowScreen(this, frame, false);
        this.screen.setMuted(true);
        if (this.stage) {
            const steps = this.stage.stepIndex;
            this.audience.present(this.stage.slide, null, false, (stage) => { stage.applySteps(steps); this.paintInk(); });
        }

        // Keys and clicks on the projector drive the show too; a click there also asks for full screen.
        root.addEventListener('keydown', this.onKey);
        root.addEventListener('click', (event) => {
            if (!doc.fullscreenElement) root.requestFullscreen?.().catch(() => { });
            if (this.inkTool && this.inkTool !== 'laser') return;
            if (this.followZoom(event, frame)) return;
            if (this.followAction(event)) return;
            const trigger = event.target?.closest?.('[data-spid]');
            if (trigger && this.playTrigger(Number(trigger.getAttribute('data-spid')))) return;
            if (!this.options.kiosk) this.next(true);
        });
        root.addEventListener('contextmenu', this.onContext);
        this.listenForInk(root, () => frame);
        this.setInkTool(this.inkTool);
        popup.addEventListener('pagehide', () => this.closeAudience());
        this.audienceButton?.classList.add('bo-show-presenter-button-on');
        root.focus();
    }

    closeAudience() {
        if (!this.audience) return;
        this.audience.dispose();
        this.audience = null;
        this.audienceRoot = null;
        this.audienceWindow = null;
        this.screen.setMuted(false);
        this.audienceButton?.classList.remove('bo-show-presenter-button-on');
    }

    flash(message) {
        const note = document.createElement('div');
        note.className = 'bo-show-message';
        note.textContent = message;
        this.root.appendChild(note);
        setTimeout(() => note.remove(), 4000);
    }

    // ---- Running ---------------------------------------------------------------------------------

    async begin() {
        // Full screen where the page may have it (not in presenter view, which is a working
        // window); the show runs in the window otherwise.
        this.onFullscreen = () => {
            if (document.fullscreenElement === this.root) this.wasFullscreen = true;
            else if (this.wasFullscreen) this.exit();
        };
        document.addEventListener('fullscreenchange', this.onFullscreen);
        // "Browsed by an individual" runs in the window, as PowerPoint's does.
        if (!this.options.presenter && !this.options.browsedInWindow) {
            try { await this.root.requestFullscreen?.(); } catch { /* refused: an iframe, a WebView, no user gesture */ }
        }
        this.root.focus({ preventScroll: true });
        await this.goTo(this.options.startPosition || 0, true);
        this.root.focus({ preventScroll: true });
    }

    async load(position) {
        if (position < 0 || position >= this.count + this.extraCount) return null;
        if (!this.cache.has(position)) {
            const promise = this.dotNetRef.invokeMethodAsync('GetSlide', position).catch(() => null);
            this.cache.set(position, promise);
        }
        return this.cache.get(position);
    }

    /** Shows a slide on every screen: with its transition when moving forward, built to its end when going back. */
    async goTo(position, withTransition, built = false, transitionOverride = null, inZoom = false) {
        if (this.disposed) return;
        position = Math.max(0, Math.min(this.count + this.extraCount - 1, position));
        const slide = await this.load(position);
        if (!slide || this.disposed) return;

        this.clearTimer();
        this.ended = false;
        this.endScreen.hidden = true;
        // "Last slide viewed", for a click action that goes back to it.
        if (this.position != null && this.position !== position) this.lastViewed = this.position;
        this.position = position;
        this.counter.textContent = position >= this.count ? `(${slide.deckIndex + 1}) / ${this.count}` : `${position + 1} / ${this.count}`;

        // A move outside the zoom's slides, other than its own return, forgets the zoom.
        if (this.zoomRun && !transitionOverride && !inZoom && !this.zoomRun.positions.includes(position)) this.zoomRun = null;
        this.updatePresenter(slide);
        const upcoming = this.upcoming(slide);
        if (upcoming != null) this.load(upcoming);
        this.drawing = null;
        const transition = this.options.showAnimations === false ? null : transitionOverride ?? (withTransition ? slide.transition : null);
        const playsAnimations = this.options.showAnimations !== false;
        this.clock.cancelFrame(this.laserFrame);
        this.recordedLaser = null;
        for (const screen of this.screens) screen.paintInk([], null);
        for (const screen of this.screens) {
            const primary = screen === this.screen;
            screen.present(slide, transition, built, (stage) => {
                const automatic = !built && playsAnimations && stage.playAutomatic();
                if (!built && !playsAnimations) stage.applySteps(stage.stepCount);
                if (!primary) return;
                this.playRecordedLaser(slide);
                this.paintInk();
                this.dotNetRef.invokeMethodAsync('ReportSlideShown', slide.deckIndex).catch(() => { });
                this.startTimer(slide);
                if (automatic) this.report('SlideShowNextBuild', `slide ${slide.deckIndex + 1}, automatic`);
                this.maybeAdvance();
            });
        }
    }

    startTimer(slide) {
        this.timeElapsed = false;
        if (this.options.rehearse) { this.startRehearsal(slide); return; }
        const after = slide.transition?.advanceAfterMs;
        if (this.options.useTimings === false || after == null) return;
        if (this.followsRehearsedTimings(slide)) { this.scheduleTimedClick(); return; }
        this.timer = this.clock.after(() => { this.timer = 0; this.timeElapsed = true; this.maybeAdvance(); }, Math.max(0, after));
    }

    /**
     * True when the slide's clicks play by themselves: "use timings" is on, the slide is timed
     * ("After"), and it carries rehearsed timings — PowerPoint's TIMING tag, one wait per click and
     * one after the last. Its click effects stay "On Click"; the timings say when.
     */
    followsRehearsedTimings(slide) {
        return this.options.useTimings !== false && this.options.showAnimations !== false && !this.options.rehearse
            && slide?.transition?.advanceAfterMs != null && Array.isArray(slide.clickTimingsMs) && slide.clickTimingsMs.length > 0;
    }

    /** The clicks of the slide's main sequence played so far, and in all — an automatic first step is not a click. */
    clickProgress() {
        const steps = this.stage.mainSteps();
        const automatic = steps[0]?.isAutomatic ? 1 : 0;
        return { played: Math.max(0, this.stage.stepIndex - automatic), total: steps.length - automatic };
    }

    /**
     * Waits for the next click's rehearsed time (counted from the click before it, or the slide's
     * arrival) and plays it; after the last click, waits its last time and moves on. A click by hand
     * plays that click at once, and the waiting starts again from it.
     */
    scheduleTimedClick() {
        if (this.timer) this.clock.cancel(this.timer);
        this.timer = 0;
        const slide = this.stage?.slide;
        if (!slide || !this.followsRehearsedTimings(slide)) return;
        const timings = slide.clickTimingsMs;
        const { played, total } = this.clickProgress();
        if (played < total) {
            const wait = timings[Math.min(played, timings.length - 1)];
            this.timer = this.clock.after(() => { this.timer = 0; this.playTimedClick(); }, Math.max(0, wait));
            return;
        }
        // The last value is the wait after the last click; without one, what is left of the slide's time.
        const spent = timings.slice(0, total).reduce((sum, ms) => sum + ms, 0);
        const wait = timings.length > total ? timings[total] : Math.max(0, slide.transition.advanceAfterMs - spent);
        this.timer = this.clock.after(() => { this.timer = 0; this.timeElapsed = true; this.maybeAdvance(); }, Math.max(0, wait));
    }

    /** A click played by the slide's timings: what moves finishes, the next click plays, the next wait starts. */
    playTimedClick() {
        if (!this.stage || this.ended || this.disposed) return;
        for (const screen of this.screens) screen.finishTransition?.();
        if (this.stage.isRunning) for (const screen of this.screens) screen.stage?.finish();
        if (this.stage.hasPendingClick()) this.playClick();
        this.scheduleTimedClick();
    }

    /** Plays the next click of the main sequence on every screen and tells .NET (PowerPoint's SlideShowNextClick). */
    playClick() {
        for (const screen of this.screens) screen.stage?.playNextStep();
        const detail = `slide ${this.stage.slide.deckIndex + 1}, click ${this.stage.stepIndex}`;
        this.report('SlideShowNextClick', detail);
        this.report('SlideShowNextBuild', detail);
        this.rehearsal?.click();
    }

    clearTimer() {
        if (this.timer) this.clock.cancel(this.timer);
        this.timer = 0;
        this.timeElapsed = false;
    }

    /** Moves on by itself once the slide's time is up and nothing waits for a click or is still moving. */
    maybeAdvance() {
        if (!this.timeElapsed || !this.stage || this.finishTransition) return;
        if (this.stage.isRunning || this.stage.hasPendingClick()) return;
        this.timeElapsed = false;
        this.advance();
    }

    advance() {
        this.rehearsal?.finishSlide();
        const run = this.zoomRun;
        const at = run ? run.positions.indexOf(this.position) : -1;
        // Inside a zoom: its next slide (a section's, hidden ones too) with that slide's own transition,
        // then back to the zoom with Return to Zoom.
        if (at >= 0 && at < run.positions.length - 1) { this.goTo(run.positions[at + 1], true, false, null, true); return; }
        if (at >= 0 && run.returnToZoom) {
            this.zoomRun = null;
            this.goTo(run.position, true, true, run.transition ? { ...run.transition, effect: 'bo-zoom-out' } : null);
            return;
        }
        // A slide hidden from the show goes on to the next slide the show plays, in the deck's order.
        if (this.position >= this.count) {
            const next = this.stage?.slide?.nextPosition;
            if (next != null) this.goTo(next, true);
            else if (this.options.loop || this.options.kiosk) this.goTo(0, true);
            else this.showEnd();
            return;
        }
        if (this.position < this.count - 1) this.goTo(this.position + 1, true);
        else if (this.options.loop || this.options.kiosk) this.goTo(0, true);
        else this.showEnd();
    }

    /**
     * The position the show goes to after a slide's last step: the zoom's next slide or the zoom itself, a
     * hidden slide's next shown slide, else the next position — null at the end. Presenter view shows it as
     * the next slide, and it is loaded ahead.
     */
    upcoming(slide) {
        const run = this.zoomRun;
        const at = run ? run.positions.indexOf(this.position) : -1;
        if (at >= 0 && at < run.positions.length - 1) return run.positions[at + 1];
        if (at >= 0 && run.returnToZoom) return run.position;
        if (this.position >= this.count) return slide?.nextPosition ?? null;
        return this.position + 1 < this.count ? this.position + 1 : null;
    }

    /** Next: finish what is moving, else play the next click, else go to the next slide. */
    next(fromClick = false) {
        this.report('SlideShowOnNext', `slide ${(this.stage?.slide.deckIndex ?? -1) + 1}`);
        if (this.ended) { this.exit(); return; }
        if (this.finishTransition) { for (const screen of this.screens) screen.finishTransition?.(); return; }
        if (!this.stage) return;
        if (this.stage.isRunning) {
            for (const screen of this.screens) screen.stage?.finish();
            this.maybeAdvance();
            return;
        }
        if (this.options.showAnimations !== false && this.stage.hasPendingClick()) {
            this.playClick();
            if (this.followsRehearsedTimings(this.stage.slide)) this.scheduleTimedClick();
            return;
        }
        if (fromClick && this.stage.slide.transition && this.stage.slide.transition.advanceOnClick === false) return;
        this.advance();
    }

    /**
     * A click on a word carrying a link ([data-link]) or on a shape with a click action (an action
     * button, a linked shape): the show follows it on every screen instead of moving on, as PowerPoint's
     * does; false when the click carries none, or one the show does not play (a media action).
     */
    followAction(event) {
        const slide = this.stage?.slide;
        if (!slide || this.ended) return false;
        const word = event.target?.closest?.('[data-link]');
        const group = event.target?.closest?.('[data-spid]');
        const url = word?.getAttribute('data-link') || (group ? slide.shapeActions?.[group.getAttribute('data-spid')] : null);
        return url ? this.follow(url, slide) : false;
    }

    /**
     * A click on a zoom (a Slide or Section Zoom, or a Summary Zoom's tile): the show zooms into its slide,
     * growing it out of the tile over the zoom's transition (or with the slide's own when it has none), and
     * with Return to Zoom comes back to this slide after the zoom's last slide. False when the click is on none.
     */
    followZoom(event, frame = this.frame) {
        const slide = this.stage?.slide;
        if (!slide?.zooms?.length || this.ended) return false;
        const box = frame.getBoundingClientRect();
        if (box.width <= 0) return false;
        const x = (event.clientX - box.left) / box.width * this.width;
        const y = (event.clientY - box.top) / box.height * this.height;
        const zoom = [...slide.zooms].reverse().find(z => x >= z.x && x <= z.x + z.width && y >= z.y && y <= z.y + z.height);
        if (!zoom) return false;
        const transition = zoom.transitionMs == null ? null : {
            effect: 'bo-zoom-in',
            durationMs: zoom.transitionMs,
            zoom: { x: zoom.x / this.width, y: zoom.y / this.height, w: zoom.width / this.width, h: zoom.height / this.height },
        };
        // The zoom's slides, in order — a hidden one too, loaded now — and the way back to this slide.
        const positions = zoom.positions?.length ? zoom.positions : [zoom.targetPosition];
        this.zoomRun = { position: this.position, positions, returnToZoom: !!zoom.returnToZoom, transition };
        this.goTo(positions[0], true, false, transition, true);
        return true;
    }

    /** Goes where a link says: a show jump, a slide (by its position in the show), the end, or an address in a new window. */
    follow(url, slide) {
        switch (url) {
            case '#next': this.advance(); return true;
            case '#previous': if (this.position > 0) this.goTo(this.position - 1, true); return true;
            case '#first': this.goTo(0, true); return true;
            case '#last': this.goTo(this.count - 1, true); return true;
            case '#lastviewed': if (this.lastViewed != null) this.goTo(this.lastViewed, true); return true;
            case '#endshow': this.exit(); return true;
        }
        if (url.startsWith('#slide:')) {
            const position = slide.linkTargets?.[url];
            // A slide deleted, or hidden from the show: the click does nothing, as it has nowhere to go.
            if (position != null) this.goTo(position, true);
            return true;
        }
        if (url.startsWith('#')) return false;
        window.open(url, '_blank', 'noopener');
        return true;
    }

    /** A click on a shape that starts a trigger sequence, on every screen; false when it starts none. */
    playTrigger(shapeId) {
        if (!this.stage?.playTrigger(shapeId)) return false;
        this.audience?.stage?.playTrigger(shapeId);
        return true;
    }

    /** Previous: undo the last click on this slide, else go to the previous slide fully built. */
    previous() {
        this.report('SlideShowOnPrevious', `slide ${(this.stage?.slide.deckIndex ?? -1) + 1}`);
        if (this.ended) { this.ended = false; this.endScreen.hidden = true; return; }
        for (const screen of this.screens) screen.finishTransition?.();
        if (!this.stage) return;
        const automatic = this.stage.mainSteps()[0]?.isAutomatic ? 1 : 0;
        if (this.stage.stepIndex > automatic) {
            const steps = this.stage.stepIndex - 1;
            for (const screen of this.screens) screen.rebuild(steps);
            if (this.followsRehearsedTimings(this.stage.slide)) this.scheduleTimedClick();
            this.rehearsal?.restart(this.stage.slide);
            return;
        }
        const run = this.zoomRun;
        const at = run ? run.positions.indexOf(this.position) : -1;
        if (at > 0) { this.goTo(run.positions[at - 1], false, true, null, true); return; }
        if (at === 0) { this.zoomRun = null; this.goTo(run.position, false, true); return; }
        if (this.position >= this.count) {
            // A hidden slide a link went to: back to the slide the show came from.
            if (this.lastViewed != null) this.goTo(this.lastViewed, false, true);
            return;
        }
        if (this.position > 0) this.goTo(this.position - 1, false, true);
    }

    /**
     * A sound or film playing across slides, kept going after its slide is left until it ends or the
     * show does; a film leaves the screen when it stops (a trimmed one pauses at its trimmed end).
     */
    carry(player, film) {
        (this.carried ||= new Set()).add(player);
        const done = () => { this.carried?.delete(player); film?.remove(); };
        player.addEventListener('ended', done, { once: true });
        if (film) player.addEventListener('pause', done, { once: true });
    }

    stopCarried() {
        for (const player of this.carried || []) {
            try { player.pause(); } catch { /* gone */ }
            filmElement(player)?.remove();
        }
        this.carried?.clear();
    }

    showEnd() {
        this.stopCarried();
        for (const screen of this.screens) screen.stage?.stopMedia();
        this.ended = true;
        this.endScreen.hidden = false;
        if (this.rehearsal) this.endScreen.textContent = 'End of rehearsal. Click to exit and choose whether to keep the new timings.';
    }

    // ---- Rehearse Timings --------------------------------------------------------------------------

    /** Starts (or carries on) timing the slide that has just arrived, with the Recording bar showing. */
    startRehearsal(slide) {
        if (!this.rehearsal) this.rehearsal = new ShowRehearsal(this, !!this.options.record);
        this.rehearsal.begin(slide);
    }


    /** A black or white screen, on the audience window too; the same key again takes it away. */
    toggleBlank(color) {
        const same = this.blanks.length > 0 && this.blanks[0].dataset.color === color;
        for (const blank of this.blanks) blank.remove();
        this.blanks = [];
        if (same) return;
        for (const root of [this.options.presenter ? this.frame : this.root, this.audienceRoot].filter(Boolean)) {
            const blank = root.ownerDocument.createElement('div');
            blank.className = 'bo-show-blank';
            blank.style.background = color;
            blank.dataset.color = color;
            blank.addEventListener('click', (event) => { event.stopPropagation(); this.toggleBlank(color); });
            root.appendChild(blank);
            this.blanks.push(blank);
        }
    }

    toggleList() { this.list.hidden = !this.list.hidden; }

    /**
     * The control bar's options menu — Use timings, Play narrations, Show media controls, Loop —
     * for this run of the show only; the Slide Show tab sets them in the file.
     */
    toggleSettings() {
        if (this.settingsMenu) { this.settingsMenu.remove(); this.settingsMenu = null; return; }
        const menu = document.createElement('div');
        menu.className = 'bo-show-settings';
        const option = (label, key, apply) => {
            const row = document.createElement('label');
            row.className = 'bo-show-settings-item';
            const box = document.createElement('input');
            box.type = 'checkbox';
            box.checked = this.options[key] !== false;
            box.addEventListener('change', () => { this.options[key] = box.checked; apply?.(box.checked); });
            row.append(box, document.createTextNode(' ' + label));
            row.addEventListener('click', (event) => event.stopPropagation());
            menu.appendChild(row);
        };
        option('Use timings', 'useTimings', (on) => {
            this.clearTimer();
            if (on && this.stage) this.startTimer(this.stage.slide);
        });
        option('Play narrations', 'showNarration', (on) => {
            for (const screen of this.screens) screen.stage?.setNarration(on);
        });
        option('Show media controls', 'showMediaControls', (on) => {
            for (const screen of this.screens) screen.stage?.setMediaControls(on);
        });
        option('Loop until Esc', 'loop');
        menu.addEventListener('click', (event) => event.stopPropagation());
        this.root.appendChild(menu);
        this.settingsMenu = menu;
    }

    key(event) {
        const key = event.key;
        if (this.inkKey(event)) return;
        if (/^[0-9]$/.test(key)) { this.typed += key; event.preventDefault(); return; }
        if (key === 'Enter' && this.typed) {
            const number = parseInt(this.typed, 10);
            this.typed = '';
            event.preventDefault();
            if (number >= 1 && number <= this.count) this.goTo(number - 1, false);
            return;
        }
        this.typed = '';
        if (this.blanks.length > 0 && !['Escape', 'b', 'B', '.', 'w', 'W', ','].includes(key)) { this.toggleBlank(this.blanks[0].dataset.color); return; }
        switch (key) {
            case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': case 'N':
                event.preventDefault();
                this.next();
                break;
            case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P':
                event.preventDefault();
                this.previous();
                break;
            case 'Home': event.preventDefault(); this.goTo(0, false); break;
            case 'End': event.preventDefault(); this.goTo(this.count - 1, false); break;
            case 'b': case 'B': case '.': this.toggleBlank('#000'); break;
            case 'w': case 'W': case ',': this.toggleBlank('#fff'); break;
            case 'g': case 'G': case '-': this.toggleList(); break;
            case 'Escape': event.preventDefault(); this.exit(); break;
        }
    }

    /** Tells .NET about a PowerPoint slide show event (SlideShowNextClick…), which it raises to watchers. */
    report(name, detail) {
        if (this.disposed) return;
        this.dotNetRef.invokeMethodAsync('ReportShowEvent', name, detail || '').catch(() => { });
    }

    exit() {
        if (this.disposed) return;
        if (this.hasInk() && !this.inkReported) {
            this.inkReported = true;
            const slides = [...this.ink.entries()]
                .map(([deckIndex, strokes]) => ({ deckIndex, strokes: strokes.filter(s => s.points.length >= 2) }))
                .filter(s => s.strokes.length > 0);
            const bytes = new TextEncoder().encode(JSON.stringify(slides));
            this.dotNetRef.invokeMethodAsync('OnShowInk', DotNet.createJSStreamReference(bytes)).catch(() => { })
                .finally(() => this.exit());
            return;
        }
        // A rehearsal hands its timings to .NET first, which asks whether to keep them.
        if (this.rehearsal && !this.rehearsalReported) {
            this.rehearsalReported = true;
            this.rehearsal.finishSlide();
            const timings = this.rehearsal.results().map(record => ({
                deckIndex: record.deckIndex,
                timingsMs: record.timingsMs,
                narration: record.wav ? DotNet.createJSStreamReference(record.wav) : null,
            }));
            this.rehearsal.dispose();
            // Reported before the end, and waited for: the end takes the show (and its .NET side) away.
            if (timings.length > 0) {
                this.dotNetRef.invokeMethodAsync('OnRehearsed', timings).catch(() => { })
                    .finally(() => this.dotNetRef.invokeMethodAsync('OnShowEnded').catch(() => { }));
                return;
            }
        }
        this.dotNetRef.invokeMethodAsync('OnShowEnded').catch(() => { });
    }

    /** Called by .NET: the next slide (as the presenter's Next would). */
    goNext() { this.next(); }
    goPrevious() { this.previous(); }
    goToPosition(position) { this.goTo(position, true); }
    currentPosition() { return this.position; }

    dispose() {
        this.disposed = true;
        this.stopCarried();
        this.rehearsal?.dispose();
        this.settingsMenu?.remove();
        this.closeInkMenu();
        this.clearTimer();
        clearTimeout(this.hideControls);
        clearInterval(this.clockTimer);
        const audienceWindow = this.audienceWindow;
        this.closeAudience();
        try { audienceWindow?.close(); } catch { /* already gone */ }
        this.screen.dispose();
        document.removeEventListener('fullscreenchange', this.onFullscreen);
        if (document.fullscreenElement === this.root) document.exitFullscreen().catch(() => { });
        this.root.remove();
    }
}

// ---- Export to Video ------------------------------------------------------------------------------------------

/**
 * Export to Video's player: the show on a VirtualClock, drawn offscreen at the video's size into one
 * ShowScreen — the same SlideStage, transitions, Morph and behaviours as on screen — with no controls,
 * no timers of its own and no sound. The exporter (video-export.js) follows the plan the editor worked
 * out (SlideVideoTimeline_Helper): it presents each slide and plays each click at its time, steps the
 * clock to every frame's exact time and draws what the screen shows. `mediaLog` collects what would
 * have been heard, for the sound track.
 */
export function createVideoRun(options) {
    return new VideoRun(options || {});
}

class VideoRun {
    constructor(options) {
        this.clock = new VirtualClock(0);
        this.options = {
            showAnimations: true,
            showNarration: !!options.playsNarrations,
            showMediaControls: false,
            useTimings: false,
        };
        this.width = options.slideWidth || 1280;
        this.height = options.slideHeight || 720;
        this.mediaLog = [];
        this.mediaInfo = options.mediaInfo || {};
        this.carried = new Set();
        this.laserColor = options.laserColorHex || 'FF0000';
        const pixelWidth = Math.max(1, options.width || this.width), pixelHeight = Math.max(1, options.height || this.height);
        // Laid out (the player measures paragraphs), but off the page: never painted where it shows.
        this.root = document.createElement('div');
        this.root.className = 'bo-show-player bo-video-stage';
        this.root.style.cssText = `position:fixed;left:${-pixelWidth - 10000}px;top:0;width:${pixelWidth}px;height:${pixelHeight}px;overflow:hidden;pointer-events:none;`;
        this.frame = document.createElement('div');
        this.frame.className = 'bo-show-frame';
        this.frame.style.width = `${pixelWidth}px`;
        this.frame.style.height = `${pixelHeight}px`;
        this.root.appendChild(this.frame);
        document.body.appendChild(this.root);
        this.screen = new ShowScreen(this, this.frame, true);
    }

    get stage() { return this.screen.stage; }

    /** Brings a slide on through its transition; its automatic step plays as it arrives, as in the show. */
    present(slide) {
        this.screen.present(slide, slide.transition, false, (stage) => { stage.arrivedAt = this.clock.now(); stage.playAutomatic(); });
        // When the slide came on, which its animated GIFs play from.
        this.screen.layer.boShownAt = this.clock.now();
    }

    /**
     * Where the recorded laser pointer is now, in the drawing's pixels, with its colour — null when the video does
     * not use the recorded timings, the slide kept no gesture, or none is under way.
     */
    laserAt() {
        const stage = this.screen.stage;
        if (!this.options.showNarration || !stage?.slide?.laserTraces?.length || stage.arrivedAt == null) return null;
        const point = laserPointAt(stage.slide.laserTraces, this.clock.now() - stage.arrivedAt);
        return point ? { ...point, color: this.laserColor } : null;
    }

    /** One click of the main sequence, as the show's timed click: what moves finishes, the next step plays. */
    click() {
        this.screen.finishTransition?.();
        const stage = this.screen.stage;
        if (!stage) return;
        if (stage.isRunning) stage.finish();
        if (stage.hasPendingClick()) stage.playNextStep();
    }

    /** The plan moves the video on, not the slide's own timings. */
    maybeAdvance() { }

    /** A sound or film playing across slides goes on until it ends or the video does. */
    carry(player) { this.carried.add(player); }

    /** The video is over: whatever is still being heard stops here. */
    finish(endMs) {
        for (const segment of this.mediaLog) if (segment.untilMs == null || segment.untilMs > endMs) segment.untilMs = endMs;
    }

    dispose() {
        this.screen.dispose();
        this.root.remove();
    }
}

// ---- Embedded in a host page --------------------------------------------------------------------------------

/**
 * The few rules the player's layers need, for a page that does not link blazor-office.css (a viewer that only
 * embeds slides). Added once per document; blazor-office.css says the same where it is linked.
 */
function ensureEmbedStyle(doc) {
    if (!doc || doc.getElementById('bo-embed-style')) return;
    const style = doc.createElement('style');
    style.id = 'bo-embed-style';
    style.textContent =
        '.bo-embed-frame { position: relative; overflow: hidden; line-height: 0; }' +
        '.bo-embed-frame > .bo-show-layer { position: absolute; inset: 0; line-height: 0; background: #fff; will-change: transform, opacity, clip-path; }' +
        '.bo-embed-frame > .bo-show-layer > svg { width: 100%; height: 100%; display: block; overflow: hidden; }' +
        '.bo-embed-frame .bo-show-video { position: absolute; object-fit: contain; background: #000; z-index: 1; }' +
        '.bo-embed-frame > .bo-show-carry { position: absolute; inset: 0; pointer-events: none; }';
    (doc.head || doc.documentElement).appendChild(style);
}

/**
 * A ShowScreen in an element of the host's own rather than the show's overlay on <body>. The host owns the
 * timing: nothing here moves on by itself — no "After" time, no rehearsed clicks — because the host is
 * following something else (a narration, a reader's scrolling) and says which click is due.
 *
 * Clicks are counted the way the Animation Pane numbers them: the step that plays by itself as the slide comes
 * up is not a click, so `showClicks(0)` is the slide as it stands once it has arrived.
 *
 * Narration does not play unless asked for (`showNarration: true`): a host that embeds slides usually has a
 * soundtrack of its own.
 */
class EmbeddedScreen {
    constructor(frame, options) {
        this.clock = realClock;
        this.options = {
            showAnimations: true,
            showNarration: false,
            showMediaControls: true,
            useTimings: false,
            ...(options || {}),
        };
        this.width = this.options.width || 1280;
        this.height = this.options.height || 720;
        this.muted = !!this.options.muted;
        this.mediaLog = [];
        this.mediaInfo = {};
        this.disposed = false;
        this.frame = frame;
        ensureEmbedStyle(frame.ownerDocument);
        frame.classList.add('bo-embed-frame');
        this.screen = new ShowScreen(this, frame, this.muted);
    }

    get stage() { return this.screen.stage; }

    /** The slide on screen, as it was handed in; null before the first. */
    get slide() { return this.stage?.slide || null; }

    /** True when the slide's first main step plays by itself on arrival. */
    get hasAutomatic() { return !!this.stage?.mainSteps()[0]?.isAutomatic; }

    /** How many clicks the slide's main sequence has. */
    get clickCount() { return this.stage ? Math.max(0, this.stage.stepCount - (this.hasAutomatic ? 1 : 0)) : 0; }

    /** How many of them have been played. */
    get clicksShown() { return this.stage ? Math.max(0, this.stage.stepIndex - (this.hasAutomatic ? 1 : 0)) : 0; }

    /** True while an effect or a transition is still moving. */
    get isRunning() { return !!this.screen.finishTransition || !!this.stage?.isRunning; }

    // What a ShowScreen asks of its show: the host decides when to move on, so there is nothing to do.
    maybeAdvance() { }

    /**
     * Shows a slide. `transition` plays it in from the slide on screen (pass the slide's own, or null for a cut);
     * `clicks` builds it that far at once (null plays its automatic step as it arrives). Resolves on arrival.
     */
    present(slide, transition = null, clicks = null) {
        if (this.disposed) return Promise.resolve();
        return new Promise(resolve => {
            this.screen.present(slide, transition, false, (stage) => {
                if (clicks != null && clicks > 0) stage.applySteps(clicks + (this.hasAutomatic ? 1 : 0));
                else if (this.options.showAnimations !== false) stage.playAutomatic();
                resolve();
            });
        });
    }

    /**
     * Brings the slide to `clicks` clicks played. The next click along plays (when `animate`), finishing whatever
     * is still moving first, as a click in the show does; any other number is drawn at once, built that far.
     */
    showClicks(clicks, animate = true) {
        const stage = this.stage;
        if (!stage || this.disposed) return;
        const target = Math.max(0, Math.min(clicks, this.clickCount));
        const shown = this.clicksShown;
        if (target === shown && !stage.isRunning) return;
        this.screen.finishTransition?.();
        if (animate && target === shown + 1) {
            stage.finish(true);
            stage.playNextStep();
            return;
        }
        if (target === shown) { stage.finish(true); return; }
        this.screen.rebuild(target + (this.hasAutomatic ? 1 : 0));
    }

    /** Finishes whatever is moving: the transition, and the effects of the click being played. */
    finish() {
        this.screen.finishTransition?.();
        this.stage?.finish(true);
    }

    dispose() {
        if (this.disposed) return;
        this.disposed = true;
        this.screen.dispose();
        this.screen.layer?.remove();
        this.frame.classList.remove('bo-embed-frame');
    }
}

// ---- Preview in the editor ------------------------------------------------------------------------------------

/**
 * The Transitions and Animations tabs' Preview: an overlay on the editor's slide that plays the
 * transition from the previous slide, then every step one after another, then removes itself.
 */
class SlidePreview {
    constructor(host, previous, slide, options) {
        this.clock = realClock;
        this.options = options || {};
        this.width = this.options.width || 1280;
        this.height = this.options.height || 720;
        this.disposed = false;
        this.overlay = document.createElement('div');
        this.overlay.className = 'bo-show-preview';
        host.appendChild(this.overlay);
        this.frame = document.createElement('div');
        this.frame.className = 'bo-show-preview-frame';
        this.overlay.appendChild(this.frame);

        let oldLayer = null, oldStage = null;
        if (previous && this.options.withTransition !== false) {
            oldLayer = this.layer();
            oldStage = new SlideStage(oldLayer, previous, this);
            oldStage.applySteps(oldStage.stepCount);
        }
        const layer = this.layer();
        this.stage = new SlideStage(layer, slide, this);
        const transition = previous && this.options.withTransition !== false ? slide.transition : null;
        let arrivedAlready = false;
        const finish = playTransition(this.frame, oldLayer, layer, transition, oldStage, this.stage, slide.morphPairs, () => {
            arrivedAlready = true;
            this.finishTransition = null;
            oldStage?.dispose();
            oldLayer?.remove();
            if (this.options.withAnimations === false) { this.close(600); return; }
            this.playAll();
        }, this.clock);
        this.finishTransition = arrivedAlready ? null : finish;
        this.overlay.addEventListener('click', () => this.dispose());
    }

    layer() {
        const layer = document.createElement('div');
        layer.className = 'bo-show-layer';
        this.frame.appendChild(layer);
        return layer;
    }

    newId() { return 'bop' + Math.random().toString(36).slice(2, 8); }

    playAll() {
        if (this.disposed) return;
        const next = () => {
            if (this.disposed) return;
            if (!this.stage.playNextStep()) { this.close(900); return; }
        };
        this.stage.onIdle = () => this.clock.after(next, 250);
        next();
    }

    close(delay) { this.closeTimer = this.clock.after(() => this.dispose(), delay); }

    dispose() {
        if (this.disposed) return;
        this.disposed = true;
        this.clock.cancel(this.closeTimer);
        this.finishTransition?.();
        this.stage?.dispose();
        this.overlay.remove();
    }
}
