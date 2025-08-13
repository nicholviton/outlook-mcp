import { describe, it, expect, beforeEach } from 'vitest';
import { LRUCache } from '../../utils/LRUCache.js';

describe('LRUCache', () => {
  let cache;
  
  beforeEach(() => {
    cache = new LRUCache(3); // Small cache for testing
  });

  describe('Basic Operations', () => {
    it('should set and get values', () => {
      cache.set('key1', 'value1');
      expect(cache.get('key1')).toBe('value1');
    });

    it('should return undefined for non-existent keys', () => {
      expect(cache.get('nonexistent')).toBeUndefined();
    });

    it('should check if key exists', () => {
      cache.set('key1', 'value1');
      expect(cache.has('key1')).toBe(true);
      expect(cache.has('nonexistent')).toBe(false);
    });

    it('should delete keys', () => {
      cache.set('key1', 'value1');
      expect(cache.has('key1')).toBe(true);
      cache.delete('key1');
      expect(cache.has('key1')).toBe(false);
    });

    it('should clear all entries', () => {
      cache.set('key1', 'value1');
      cache.set('key2', 'value2');
      expect(cache.size).toBe(2);
      cache.clear();
      expect(cache.size).toBe(0);
    });
  });

  describe('LRU Behavior', () => {
    it('should evict least recently used items when capacity is exceeded', () => {
      cache.set('key1', 'value1');
      cache.set('key2', 'value2');
      cache.set('key3', 'value3');
      
      // Cache is now full
      expect(cache.size).toBe(3);
      
      // Adding a new item should evict key1 (least recently used)
      cache.set('key4', 'value4');
      expect(cache.size).toBe(3);
      expect(cache.has('key1')).toBe(false);
      expect(cache.has('key2')).toBe(true);
      expect(cache.has('key3')).toBe(true);
      expect(cache.has('key4')).toBe(true);
    });

    it('should update LRU order when accessing items', () => {
      cache.set('key1', 'value1');
      cache.set('key2', 'value2');
      cache.set('key3', 'value3');
      
      // Access key1 to make it most recently used
      cache.get('key1');
      
      // Add new item, should evict key2 (now least recently used)
      cache.set('key4', 'value4');
      expect(cache.has('key1')).toBe(true);
      expect(cache.has('key2')).toBe(false);
      expect(cache.has('key3')).toBe(true);
      expect(cache.has('key4')).toBe(true);
    });

    it('should update LRU order when setting existing keys', () => {
      cache.set('key1', 'value1');
      cache.set('key2', 'value2');
      cache.set('key3', 'value3');
      
      // Update key1 to make it most recently used
      cache.set('key1', 'updated_value1');
      
      // Add new item, should evict key2 (now least recently used)
      cache.set('key4', 'value4');
      expect(cache.get('key1')).toBe('updated_value1');
      expect(cache.has('key2')).toBe(false);
      expect(cache.has('key3')).toBe(true);
      expect(cache.has('key4')).toBe(true);
    });
  });

  describe('Edge Cases', () => {
    it('should handle zero capacity', () => {
      const zeroCache = new LRUCache(0);
      zeroCache.set('key1', 'value1');
      expect(zeroCache.size).toBe(0);
      expect(zeroCache.has('key1')).toBe(false);
    });

    it('should handle single item capacity', () => {
      const singleCache = new LRUCache(1);
      singleCache.set('key1', 'value1');
      expect(singleCache.size).toBe(1);
      
      singleCache.set('key2', 'value2');
      expect(singleCache.size).toBe(1);
      expect(singleCache.has('key1')).toBe(false);
      expect(singleCache.has('key2')).toBe(true);
    });

    it('should handle negative capacity', () => {
      expect(() => new LRUCache(-1)).toThrow('Capacity must be a positive number');
    });

    it('should handle non-numeric capacity', () => {
      expect(() => new LRUCache('invalid')).toThrow('Capacity must be a positive number');
    });
  });

  describe('Performance and Memory', () => {
    it('should maintain constant time complexity for operations', () => {
      const largeCache = new LRUCache(1000);
      
      // Fill cache
      for (let i = 0; i < 1000; i++) {
        largeCache.set(`key${i}`, `value${i}`);
      }
      
      // Time operations
      const start = performance.now();
      largeCache.get('key500');
      largeCache.set('newKey', 'newValue');
      largeCache.delete('key100');
      const end = performance.now();
      
      // Should complete very quickly (under 10ms even on slow systems)
      expect(end - start).toBeLessThan(10);
    });

    it('should not grow beyond capacity', () => {
      const capacity = 100;
      const testCache = new LRUCache(capacity);
      
      // Add more items than capacity
      for (let i = 0; i < capacity * 2; i++) {
        testCache.set(`key${i}`, `value${i}`);
      }
      
      expect(testCache.size).toBe(capacity);
    });
  });

  describe('TTL (Time To Live)', () => {
    it('should support TTL for cache entries', async () => {
      const ttlCache = new LRUCache(10, { ttl: 100 }); // 100ms TTL
      
      ttlCache.set('key1', 'value1');
      expect(ttlCache.get('key1')).toBe('value1');
      
      // Wait for TTL to expire
      await new Promise(resolve => setTimeout(resolve, 150));
      
      expect(ttlCache.get('key1')).toBeUndefined();
      expect(ttlCache.has('key1')).toBe(false);
    });

    it('should clean up expired entries automatically', async () => {
      const ttlCache = new LRUCache(10, { ttl: 50 });
      
      ttlCache.set('key1', 'value1');
      ttlCache.set('key2', 'value2');
      expect(ttlCache.size).toBe(2);
      
      // Wait for TTL to expire
      await new Promise(resolve => setTimeout(resolve, 100));
      
      // Access should trigger cleanup
      ttlCache.get('key1');
      expect(ttlCache.size).toBe(0);
    });
  });

  describe('Statistics and Monitoring', () => {
    it('should track hit and miss statistics', () => {
      cache.set('key1', 'value1');
      
      // Hit
      cache.get('key1');
      
      // Miss
      cache.get('nonexistent');
      
      const stats = cache.getStats();
      expect(stats.hits).toBe(1);
      expect(stats.misses).toBe(1);
      expect(stats.hitRate).toBe(0.5);
    });

    it('should track eviction count', () => {
      cache.set('key1', 'value1');
      cache.set('key2', 'value2');
      cache.set('key3', 'value3');
      
      // This should cause eviction
      cache.set('key4', 'value4');
      
      const stats = cache.getStats();
      expect(stats.evictions).toBe(1);
    });

    it('should reset statistics', () => {
      cache.set('key1', 'value1');
      cache.get('key1');
      cache.get('nonexistent');
      
      let stats = cache.getStats();
      expect(stats.hits).toBe(1);
      expect(stats.misses).toBe(1);
      
      cache.resetStats();
      
      stats = cache.getStats();
      expect(stats.hits).toBe(0);
      expect(stats.misses).toBe(0);
    });
  });
});