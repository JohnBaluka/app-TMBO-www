/**
 * Context menu helper for site link actions.
 * Provides window management and clipboard operations.
 */

/**
 * Opens a URL in a new browser window (not tab) with 80% screen size, centered.
 * @param {string} url - The URL to open.
 */
export function openInNewWindow(url) {
    const width = Math.round(screen.availWidth * 0.8);
    const height = Math.round(screen.availHeight * 0.8);
    const left = Math.round((screen.availWidth - width) / 2);
    const top = Math.round((screen.availHeight - height) / 2);
    window.open(url, '_blank', `width=${width},height=${height},left=${left},top=${top}`);
}

/**
 * Opens a URL in a new browser window positioned on the right half of the screen.
 * @param {string} url - The URL to open.
 */
export function openInSplitScreenWindow(url) {
    const width = Math.round(screen.availWidth / 2);
    const height = screen.availHeight;
    const left = Math.round(screen.availWidth / 2);
    const top = 0;
    window.open(url, '_blank', `width=${width},height=${height},left=${left},top=${top}`);
}

/**
 * Copies text to the clipboard.
 * @param {string} text - The text to copy.
 * @returns {Promise<boolean>} True if copy succeeded.
 */
export async function copyToClipboard(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch {
        return false;
    }
}
