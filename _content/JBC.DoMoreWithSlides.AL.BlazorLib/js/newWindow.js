// Browser windows of their own, one per key (ESM module).
//
// Imported on first use by BrowserNewWindow_AppService rather than registered in every head document,
// for the reason layout.js is: a head that serves no static assets loses the feature, not the page.

// Asked to answer before it is trusted. A stubbed JS runtime resolves the import and then answers
// every later call with nothing rather than throwing, which would read as a window that was refused.
export function ready() {
    return true;
}

// The windows this document opened, by key. A shortcut rather than the whole answer: reloading the
// page that opened them empties this map and leaves the windows open, which is why openOrFocus asks
// the browser for the window by name as well.
const opened = new Map();

// Opens url in a window called key, or brings forward the window already called that.
// Answers "opened", "focused" or "blocked".
export function openOrFocus(url, key) {
    const known = opened.get(key);

    if (known && !known.closed) {
        known.focus();
        return 'focused';
    }

    // An empty address finds a window that already has this name WITHOUT navigating it, which is what
    // keeps a second click from reloading a preview somebody is part-way down. A name nobody has used
    // yet opens a blank window instead, and that is the one that is sent to the address.
    const win = window.open('', key, features());

    if (!win) {
        return 'blocked';
    }

    opened.set(key, win);

    let fresh;

    try {
        fresh = !win.location.href || win.location.href === 'about:blank';
    } catch {
        // Only a window on another origin refuses to say, and a window this opened never is one.
        fresh = false;
    }

    if (fresh) {
        win.location.href = new URL(url, document.baseURI).href;
    }

    win.focus();

    return fresh ? 'opened' : 'focused';
}

// 80% of the screen, centered - the shape contextMenuHelper.js opens a site link in, so two windows
// this application opens look like the same kind of thing.
function features() {
    const width = Math.round(window.screen.availWidth * 0.8);
    const height = Math.round(window.screen.availHeight * 0.8);
    const left = Math.round((window.screen.availWidth - width) / 2);
    const top = Math.round((window.screen.availHeight - height) / 2);

    return `width=${width},height=${height},left=${left},top=${top}`;
}
