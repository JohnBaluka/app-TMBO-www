export function registerGitHubAuthCallback(dotNetRef) {
    function handler(event) {
        if (event.data && event.data.type === 'github-oauth-callback') {
            window.removeEventListener('message', handler);
            dotNetRef.invokeMethodAsync('OnOAuthCallback',
                event.data.accessToken || '',
                event.data.error || '');
        }
    }
    window.addEventListener('message', handler);
}
