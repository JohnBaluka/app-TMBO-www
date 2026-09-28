// HTML5 drag and drop for a web view that starts a drag and never delivers the drop.
//
// WinUI's WebView2 — the MAUI Windows head — hosts the browser as a composition visual, and the
// Windows drag loop a native HTML5 drag runs on is never handed back to it: `dragstart` and `dragend`
// fire, and not one `dragenter`, `dragover` or `drop` in between. Measured on 2026-09-23 with a real
// mouse drag in the MAUI head (dragstart 1, dragend 1, nothing else), against the WinForms head, whose
// WebView2 is a child window and delivers all four. So every drag on every page — the Slide Site
// Designer's tree and thumbnails, the tools page's Projects and Sections tabs — did nothing there.
//
// What this does is take the drag off Windows and run it on the mouse. The native drag is cancelled
// the moment it starts (the page's own `dragstart` handlers still run, on the same event), and while
// the button is held the element under the pointer is sent `dragenter`, `dragover` and `dragleave`,
// then `drop` — only where the last `dragover` was accepted, the rule a real drag keeps — and the
// source `dragend`. Blazor's `@ondragover:preventDefault` and `@ondrop` see ordinary drag events, so
// no component changes.
//
// Loaded as a classic script by a host that needs it — the MAUI Windows head's index.html — and by no
// other: a browser, the WinForms head and the web heads deliver drops themselves, and this would only
// replace a working drag with an emulated one.

(function () {
    'use strict';

    if (window.__dmwsDragDropShim) return;
    window.__dmwsDragDropShim = true;

    /** The drag in progress, or null. */
    let session = null;

    /** Marks the page while a drag runs: the grabbing cursor, and no text selection under it. */
    const DRAGGING_CLASS = 'dmws-dragging';

    const style = document.createElement('style');
    style.textContent =
        `html.${DRAGGING_CLASS}, html.${DRAGGING_CLASS} * { cursor: grabbing !important; user-select: none !important; }`;
    document.head.appendChild(style);

    /** Sends one drag event to an element, carrying the pointer's place and the drag's data. */
    function fire(type, element, pointer) {
        if (!element) return null;

        const event = new DragEvent(type, {
            bubbles: true,
            cancelable: type !== 'dragleave' && type !== 'dragend',
            composed: true,
            clientX: pointer ? pointer.clientX : 0,
            clientY: pointer ? pointer.clientY : 0,
            screenX: pointer ? pointer.screenX : 0,
            screenY: pointer ? pointer.screenY : 0,
            ctrlKey: pointer ? pointer.ctrlKey : false,
            shiftKey: pointer ? pointer.shiftKey : false,
            altKey: pointer ? pointer.altKey : false,
            metaKey: pointer ? pointer.metaKey : false,
            dataTransfer: session ? session.dataTransfer : null,
        });

        element.dispatchEvent(event);

        return event;
    }

    function end(pointer, dropped) {
        if (!session) return;

        const ending = session;

        if (dropped && ending.target && ending.accepted) {
            fire('drop', ending.target, pointer);
        } else if (ending.target) {
            fire('dragleave', ending.target, pointer);
        }

        fire('dragend', ending.source, pointer);

        session = null;
        document.documentElement.classList.remove(DRAGGING_CLASS);
    }

    // The native drag is cancelled as it starts; the page's own dragstart handlers still see it, which
    // is where every component here records what is being dragged.
    document.addEventListener('dragstart', event => {
        if (!event.isTrusted) return;

        event.preventDefault();

        session = {
            source: event.target,
            dataTransfer: new DataTransfer(),
            target: null,
            accepted: false,
        };

        const selection = window.getSelection ? window.getSelection() : null;
        if (selection) selection.removeAllRanges();

        document.documentElement.classList.add(DRAGGING_CLASS);
    }, true);

    document.addEventListener('mousemove', event => {
        if (!session) return;

        if ((event.buttons & 1) === 0) {
            // The button came up somewhere this page never heard about — outside the window.
            end(event, false);

            return;
        }

        const under = document.elementFromPoint(event.clientX, event.clientY);

        if (under !== session.target) {
            if (session.target) fire('dragleave', session.target, event);
            if (under) fire('dragenter', under, event);

            session.target = under;
            session.accepted = false;
        }

        if (under) {
            const over = fire('dragover', under, event);

            session.accepted = !!over && over.defaultPrevented;
        }

        fire('drag', session.source, event);
    }, true);

    document.addEventListener('mouseup', event => {
        if (!session) return;

        end(event, true);
    }, true);

    document.addEventListener('keydown', event => {
        if (session && event.key === 'Escape') end(null, false);
    }, true);

    window.addEventListener('blur', () => end(null, false));
})();
