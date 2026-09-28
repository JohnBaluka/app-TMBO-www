window.cacheWrapper = {
    async openCache(cacheName) {
        return await caches.open(cacheName);
    },

    async storeInCache(cacheName, key, value, mimeType) {
        try {
            const cache = await caches.open(cacheName);
            const response = new Response(value, { headers: { "Content-Type": mimeType } });
            await cache.put(key, response);
            console.log(`Cached: ${key}`);
        } catch (error) {
            console.error("Error storing in cache:", error);
        }
    },

    async retrieveFromCache(cacheName, key) {
        try {
            const cache = await caches.open(cacheName);
            const response = await cache.match(key);
            if (!response) return null;
            return await response.text(); // Supports text-based responses (HTML, JSON, etc.)
        } catch (error) {
            console.error("Error retrieving from cache:", error);
            return null;
        }
    },

    async deleteFromCache(cacheName, key) {
        try {
            const cache = await caches.open(cacheName);
            return await cache.delete(key);
        } catch (error) {
            console.error("Error deleting from cache:", error);
            return false;
        }
    },

    async clearCache(cacheName) {
        try {
            const cache = await caches.open(cacheName);
            const keys = await cache.keys();
            for (let request of keys) {
                await cache.delete(request);
            }
            console.log(`Cache cleared: ${cacheName}`);
        } catch (error) {
            console.error("Error clearing cache:", error);
        }
    }
};
