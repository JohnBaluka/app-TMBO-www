// Browser window helpers for the shell: the breakpoint badge and the version's shortcut key (ESM module).
//
// Imported on first use by LayoutModule_Helper rather than registered in every head document: each is
// one small feature, and a head that serves no static assets — an Office web add-in pane — must lose
// the feature rather than fail to render the layout.

// Asked to answer before it is trusted. A stubbed JS runtime resolves the import and then answers
// every later call with nothing rather than throwing, which looks exactly like a badge that never
// updates. Every module in this repository that is imported at run time carries one of these.
export function ready() {
    return true;
}

export function getWindowWidth() {
    return window.innerWidth;
}

// Subscribes to window resize and invokes OnWindowWidthChanged on the .NET ref.
// Returns a handle whose dispose() removes the listener.
export function watchWindowWidth(dotNetRef) {
    function handler() {
        dotNetRef.invokeMethodAsync('OnWindowWidthChanged', window.innerWidth);
    }

    window.addEventListener('resize', handler);

    return {
        dispose: function () {
            window.removeEventListener('resize', handler);
        }
    };
}

// Invokes OnShortcutKey on the .NET ref whenever `key` (a KeyboardEvent.key, such as 'F8') is pressed
// on its own - no Ctrl, Alt, Shift or Meta, and not the auto-repeat of a key held down, which would
// toggle the version on and off for as long as the key was held. The key's own browser action, if
// it has one, is cancelled. Listens on the window, so it works wherever focus is in this document
// (not inside an iframe, which is a document of its own). Returns a handle whose dispose() removes
// the listener.
export function watchShortcutKey(dotNetRef, key) {
    function handler(event) {
        if (event.key !== key || event.repeat
            || event.ctrlKey || event.altKey || event.shiftKey || event.metaKey) {
            return;
        }

        event.preventDefault();

        dotNetRef.invokeMethodAsync('OnShortcutKey');
    }

    window.addEventListener('keydown', handler);

    return {
        dispose: function () {
            window.removeEventListener('keydown', handler);
        }
    };
}
