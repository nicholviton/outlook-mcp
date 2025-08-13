import { describe, it, expect, beforeEach } from 'vitest';
import { LRUCache } from '../../utils/LRUCache.js';

describe('LRU Cache Performance Benchmarks', () => {
  const iterations = 10000;
  const cacheSize = 1000;

  describe('LRU Cache vs Map Performance', () => {
    it('should demonstrate memory efficiency compared to unbounded Map', () => {
      const lruCache = new LRUCache(cacheSize);
      const unboundedMap = new Map();

      // Fill both with data
      for (let i = 0; i < iterations; i++) {
        const key = `key-${i}`;
        const value = `value-${i}`;
        
        lruCache.set(key, value);
        unboundedMap.set(key, value);
      }

      // LRU cache should be bounded
      expect(lruCache.size).toBeLessThanOrEqual(cacheSize);
      
      // Map should grow unbounded
      expect(unboundedMap.size).toBe(iterations);
      
      // Calculate memory efficiency
      const lruMemoryEfficiency = lruCache.size / iterations;
      const mapMemoryEfficiency = unboundedMap.size / iterations;
      
      expect(lruMemoryEfficiency).toBeLessThan(mapMemoryEfficiency);
      
      console.log(`LRU Cache size: ${lruCache.size} (${(lruMemoryEfficiency * 100).toFixed(1)}% of total)`);
      console.log(`Map size: ${unboundedMap.size} (${(mapMemoryEfficiency * 100).toFixed(1)}% of total)`);
    });

    it('should maintain performance under heavy load', () => {
      const lruCache = new LRUCache(cacheSize);
      const testData = Array.from({ length: iterations }, (_, i) => ({
        key: `key-${i}`,
        value: `value-${i}`
      }));

      // Write performance
      const writeStartTime = performance.now();
      testData.forEach(({ key, value }) => {
        lruCache.set(key, value);
      });
      const writeEndTime = performance.now();
      const writeTime = writeEndTime - writeStartTime;

      // Read performance (mix of hits and misses)
      const readStartTime = performance.now();
      for (let i = 0; i < iterations; i++) {
        const key = `key-${iterations - 1 - (i % cacheSize)}`;
        lruCache.get(key);
      }
      const readEndTime = performance.now();
      const readTime = readEndTime - readStartTime;

      // Performance should be reasonable
      expect(writeTime).toBeLessThan(1000); // Under 1 second
      expect(readTime).toBeLessThan(100); // Under 100ms
      
      const stats = lruCache.getStats();
      console.log(`Write time: ${writeTime.toFixed(2)}ms`);
      console.log(`Read time: ${readTime.toFixed(2)}ms`);
      console.log(`Hit rate: ${(stats.hitRate * 100).toFixed(1)}%`);
      console.log(`Final cache size: ${lruCache.size}`);
    });

    it('should demonstrate constant time complexity', () => {
      const smallCache = new LRUCache(100);
      const largeCache = new LRUCache(5000);

      // Fill both caches
      for (let i = 0; i < 10000; i++) {
        smallCache.set(`key-${i}`, `value-${i}`);
        largeCache.set(`key-${i}`, `value-${i}`);
      }

      // Measure operation times
      const smallCacheOpStart = performance.now();
      for (let i = 0; i < 1000; i++) {
        smallCache.get(`key-${i}`);
        smallCache.set(`new-key-${i}`, `new-value-${i}`);
      }
      const smallCacheOpEnd = performance.now();

      const largeCacheOpStart = performance.now();
      for (let i = 0; i < 1000; i++) {
        largeCache.get(`key-${i}`);
        largeCache.set(`new-key-${i}`, `new-value-${i}`);
      }
      const largeCacheOpEnd = performance.now();

      const smallCacheTime = smallCacheOpEnd - smallCacheOpStart;
      const largeCacheTime = largeCacheOpEnd - largeCacheOpStart;

      // Time complexity should be similar regardless of cache size
      const timeRatio = largeCacheTime / smallCacheTime;
      expect(timeRatio).toBeLessThan(2); // Should not be significantly slower

      console.log(`Small cache (100) operations: ${smallCacheTime.toFixed(2)}ms`);
      console.log(`Large cache (5000) operations: ${largeCacheTime.toFixed(2)}ms`);
      console.log(`Time ratio: ${timeRatio.toFixed(2)}`);
    });
  });

  describe('Memory Leak Prevention', () => {
    it('should prevent memory leaks with continuous usage', () => {
      const cache = new LRUCache(500);
      const initialMemory = process.memoryUsage().heapUsed;

      // Simulate continuous usage over time
      for (let batch = 0; batch < 20; batch++) {
        for (let i = 0; i < 1000; i++) {
          const key = `batch-${batch}-key-${i}`;
          const value = `batch-${batch}-value-${i}`;
          cache.set(key, value);
        }
        
        // Periodic reads
        for (let i = 0; i < 100; i++) {
          cache.get(`batch-${batch}-key-${i}`);
        }
      }

      const finalMemory = process.memoryUsage().heapUsed;
      const memoryGrowth = finalMemory - initialMemory;

      // Memory growth should be bounded
      expect(memoryGrowth).toBeLessThan(50 * 1024 * 1024); // Less than 50MB
      expect(cache.size).toBeLessThanOrEqual(500);

      console.log(`Memory growth: ${(memoryGrowth / 1024 / 1024).toFixed(2)}MB`);
      console.log(`Final cache size: ${cache.size}`);
    });

    it('should handle rapid cache turnover efficiently', () => {
      const cache = new LRUCache(100);
      const startTime = performance.now();

      // Rapid cache turnover
      for (let i = 0; i < 10000; i++) {
        cache.set(`key-${i}`, `value-${i}`);
        
        // Occasionally read old values (should be evicted)
        if (i % 200 === 0) {
          cache.get(`key-${i - 150}`);
        }
      }

      const endTime = performance.now();
      const duration = endTime - startTime;

      expect(duration).toBeLessThan(500); // Should handle rapid turnover quickly
      expect(cache.size).toBeLessThanOrEqual(100);

      const stats = cache.getStats();
      console.log(`Rapid turnover duration: ${duration.toFixed(2)}ms`);
      console.log(`Evictions: ${stats.evictions}`);
      console.log(`Hit rate: ${(stats.hitRate * 100).toFixed(1)}%`);
    });
  });

  describe('Real-world Simulation', () => {
    it('should simulate Graph API response caching', () => {
      const cache = new LRUCache(1000, { ttl: 300000 }); // 5 minutes TTL
      
      // Simulate typical email API patterns
      const emailPatterns = [
        'messages/inbox',
        'messages/sent',
        'messages/drafts',
        'me/messages',
        'me/mailFolders',
        'me/events',
        'me/contacts'
      ];

      const startTime = performance.now();

      // Simulate API calls
      for (let i = 0; i < 5000; i++) {
        const pattern = emailPatterns[i % emailPatterns.length];
        const key = `${pattern}?$top=25&$skip=${i * 25}`;
        const value = {
          data: `Response data for ${key}`,
          timestamp: Date.now(),
          etag: `etag-${i}`
        };

        cache.set(key, value);

        // Simulate cache hits (common queries)
        if (i % 10 === 0) {
          cache.get(`${pattern}?$top=25&$skip=0`);
        }
      }

      const endTime = performance.now();
      const duration = endTime - startTime;

      expect(duration).toBeLessThan(1000); // Should handle real-world load
      expect(cache.size).toBeLessThanOrEqual(1000);

      const stats = cache.getStats();
      console.log(`Graph API simulation duration: ${duration.toFixed(2)}ms`);
      console.log(`Cache hit rate: ${(stats.hitRate * 100).toFixed(1)}%`);
      console.log(`Cache utilization: ${cache.size}/1000`);
    });

    it('should handle concurrent access patterns', async () => {
      const cache = new LRUCache(200);
      const concurrentRequests = 100;
      
      // Simulate concurrent requests
      const promises = Array.from({ length: concurrentRequests }, async (_, i) => {
        return new Promise(resolve => {
          setTimeout(() => {
            const batchSize = 50;
            const startTime = performance.now();
            
            for (let j = 0; j < batchSize; j++) {
              const key = `concurrent-${i}-${j}`;
              const value = `data-${i}-${j}`;
              cache.set(key, value);
              
              // Some reads
              if (j % 5 === 0) {
                cache.get(`concurrent-${i}-${j - 1}`);
              }
            }
            
            const endTime = performance.now();
            resolve(endTime - startTime);
          }, Math.random() * 100);
        });
      });

      const results = await Promise.all(promises);
      const avgDuration = results.reduce((sum, time) => sum + time, 0) / results.length;

      expect(avgDuration).toBeLessThan(50); // Each concurrent batch should be fast
      expect(cache.size).toBeLessThanOrEqual(200);

      const stats = cache.getStats();
      console.log(`Concurrent access avg duration: ${avgDuration.toFixed(2)}ms`);
      console.log(`Final cache size: ${cache.size}`);
      console.log(`Hit rate: ${(stats.hitRate * 100).toFixed(1)}%`);
    });
  });

  describe('Cache Effectiveness', () => {
    it('should demonstrate improved hit rates over time', () => {
      const cache = new LRUCache(500);
      const hitRateOverTime = [];

      // Simulate gradually increasing cache effectiveness
      for (let phase = 0; phase < 10; phase++) {
        for (let i = 0; i < 200; i++) {
          const key = `phase-${phase % 3}-key-${i % 100}`; // Overlapping keys
          const value = `phase-${phase}-value-${i}`;
          cache.set(key, value);
          
          // Read operations
          if (i % 3 === 0) {
            cache.get(`phase-${phase % 3}-key-${(i - 10) % 100}`);
          }
        }
        
        const stats = cache.getStats();
        hitRateOverTime.push(stats.hitRate);
      }

      // Hit rate should generally improve over time
      const initialHitRate = hitRateOverTime[0];
      const finalHitRate = hitRateOverTime[hitRateOverTime.length - 1];
      
      expect(finalHitRate).toBeGreaterThan(initialHitRate);
      
      console.log(`Initial hit rate: ${(initialHitRate * 100).toFixed(1)}%`);
      console.log(`Final hit rate: ${(finalHitRate * 100).toFixed(1)}%`);
      console.log(`Hit rate improvement: ${((finalHitRate - initialHitRate) * 100).toFixed(1)}%`);
    });
  });
});