// Browser window helpers for the shell's breakpoint badge (ESM module).
//
// Imported on first use by WindowBreakpoint_Helper rather than registered in every head document:
// the badge is one label, and a head that serves no static assets — an Office web add-in pane — must
// lose the label rather than fail to render the layout.

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
