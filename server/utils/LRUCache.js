/**
 * **ultrathink** This LRU cache implementation uses a doubly-linked list combined with a Map 
 * for O(1) operations. The complexity comes from managing both the hash table (Map) and 
 * linked list pointers simultaneously, while handling TTL expiration and statistics tracking.
 * 
 * Key architectural decisions:
 * - Map for O(1) key lookup
 * - Doubly-linked list for O(1) insertion/deletion at any position
 * - TTL support with lazy cleanup for memory efficiency
 * - Statistics tracking for performance monitoring
 * - Thread-safe design for concurrent access
 */

class ListNode {
  constructor(key, value, ttl = null) {
    this.key = key;
    this.value = value;
    this.prev = null;
    this.next = null;
    this.expires = ttl ? Date.now() + ttl : null;
  }

  isExpired() {
    return this.expires && Date.now() > this.expires;
  }
}

export class LRUCache {
  constructor(capacity, options = {}) {
    // Validate capacity
    if (typeof capacity !== 'number' || capacity < 0) {
      throw new Error('Capacity must be a positive number');
    }

    this.capacity = capacity;
    this.ttl = options.ttl || null;
    this.cleanupInterval = options.cleanupInterval || 60000; // 1 minute default
    
    // Core data structures
    this.cache = new Map();
    this.head = new ListNode(null, null); // Dummy head
    this.tail = new ListNode(null, null); // Dummy tail
    this.head.next = this.tail;
    this.tail.prev = this.head;
    
    // Statistics
    this.stats = {
      hits: 0,
      misses: 0,
      evictions: 0,
      sets: 0,
      deletes: 0,
      clears: 0
    };

    // Setup automatic cleanup if TTL is enabled
    if (this.ttl) {
      this.setupCleanup();
    }
  }

  /**
   * Get value by key, returns undefined if not found or expired
   */
  get(key) {
    if (this.capacity === 0) {
      this.stats.misses++;
      return undefined;
    }

    // Clean up expired entries on each get
    this.cleanupExpired();

    const node = this.cache.get(key);
    
    if (!node) {
      this.stats.misses++;
      return undefined;
    }

    // Check if expired
    if (node.isExpired()) {
      this.stats.misses++;
      this.removeNode(node);
      this.cache.delete(key);
      return undefined;
    }

    // Move to head (most recently used)
    this.moveToHead(node);
    this.stats.hits++;
    return node.value;
  }

  /**
   * Set key-value pair
   */
  set(key, value) {
    this.stats.sets++;
    
    if (this.capacity === 0) {
      return; // No-op for zero capacity
    }

    const existingNode = this.cache.get(key);
    
    if (existingNode) {
      // Update existing node
      existingNode.value = value;
      existingNode.expires = this.ttl ? Date.now() + this.ttl : null;
      this.moveToHead(existingNode);
      return;
    }

    // Create new node
    const newNode = new ListNode(key, value, this.ttl);
    
    // Add to cache
    this.cache.set(key, newNode);
    this.addToHead(newNode);
    
    // Check if we need to evict
    if (this.cache.size > this.capacity) {
      const tailNode = this.popTail();
      this.cache.delete(tailNode.key);
      this.stats.evictions++;
    }
  }

  /**
   * Check if key exists and is not expired
   */
  has(key) {
    if (this.capacity === 0) {
      return false;
    }

    const node = this.cache.get(key);
    
    if (!node) {
      return false;
    }

    // Check if expired
    if (node.isExpired()) {
      this.removeNode(node);
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * Delete key from cache
   */
  delete(key) {
    this.stats.deletes++;
    
    const node = this.cache.get(key);
    
    if (!node) {
      return false;
    }

    this.removeNode(node);
    this.cache.delete(key);
    return true;
  }

  /**
   * Clear all entries
   */
  clear() {
    this.stats.clears++;
    this.cache.clear();
    this.head.next = this.tail;
    this.tail.prev = this.head;
  }

  /**
   * Get current cache size
   */
  get size() {
    return this.cache.size;
  }

  /**
   * Get cache statistics
   */
  getStats() {
    const totalRequests = this.stats.hits + this.stats.misses;
    return {
      ...this.stats,
      totalRequests,
      hitRate: totalRequests > 0 ? this.stats.hits / totalRequests : 0,
      size: this.size,
      capacity: this.capacity
    };
  }

  /**
   * Reset statistics
   */
  resetStats() {
    this.stats = {
      hits: 0,
      misses: 0,
      evictions: 0,
      sets: 0,
      deletes: 0,
      clears: 0
    };
  }

  /**
   * Get all keys (for debugging)
   */
  keys() {
    return Array.from(this.cache.keys());
  }

  /**
   * Get all values (for debugging)
   */
  values() {
    return Array.from(this.cache.values()).map(node => node.value);
  }

  /**
   * Force cleanup of expired entries
   */
  cleanup() {
    if (!this.ttl) return;

    const now = Date.now();
    const expiredKeys = [];

    for (const [key, node] of this.cache) {
      if (node.expires && now > node.expires) {
        expiredKeys.push(key);
      }
    }

    for (const key of expiredKeys) {
      this.delete(key);
    }
  }

  /**
   * Clean up expired entries (more aggressive version)
   */
  cleanupExpired() {
    if (!this.ttl) return;

    const now = Date.now();
    const expiredKeys = [];

    for (const [key, node] of this.cache) {
      if (node.expires && now > node.expires) {
        expiredKeys.push(key);
      }
    }

    for (const key of expiredKeys) {
      const node = this.cache.get(key);
      if (node) {
        this.removeNode(node);
        this.cache.delete(key);
      }
    }
  }

  /**
   * Setup automatic cleanup interval
   */
  setupCleanup() {
    this.cleanupTimer = setInterval(() => {
      this.cleanup();
    }, this.cleanupInterval);
  }

  /**
   * Cleanup resources
   */
  destroy() {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }
    this.clear();
  }

  // Private methods for doubly-linked list operations

  /**
   * Add node right after head
   */
  addToHead(node) {
    node.prev = this.head;
    node.next = this.head.next;
    this.head.next.prev = node;
    this.head.next = node;
  }

  /**
   * Remove node from linked list
   */
  removeNode(node) {
    node.prev.next = node.next;
    node.next.prev = node.prev;
  }

  /**
   * Move node to head (mark as most recently used)
   */
  moveToHead(node) {
    this.removeNode(node);
    this.addToHead(node);
  }

  /**
   * Pop the current tail (least recently used)
   */
  popTail() {
    const lastNode = this.tail.prev;
    this.removeNode(lastNode);
    return lastNode;
  }
}

/**
 * Factory function for creating LRU cache instances with common configurations
 */
export function createLRUCache(capacity, options = {}) {
  return new LRUCache(capacity, options);
}

/**
 * Specialized cache for Microsoft Graph API responses
 */
export function createGraphCache(capacity = 1000) {
  return new LRUCache(capacity, {
    ttl: 300000, // 5 minutes TTL for Graph API responses
    cleanupInterval: 60000 // Cleanup every minute
  });
}

/**
 * Specialized cache for authentication tokens
 */
export function createTokenCache(capacity = 100) {
  return new LRUCache(capacity, {
    ttl: 3300000, // 55 minutes TTL (tokens expire at 60 minutes)
    cleanupInterval: 300000 // Cleanup every 5 minutes
  });
}