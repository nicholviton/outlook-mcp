/**
 * **ultrathink** HTTP connection pooling configuration for Microsoft Graph API calls.
 * 
 * This module implements connection pooling to improve performance and reduce latency
 * when making multiple requests to the Graph API. The complexity involves:
 * 1. Managing HTTP/HTTPS agents with optimized connection pooling
 * 2. Balancing connection reuse vs memory usage
 * 3. Handling connection timeouts and keep-alive settings
 * 4. Supporting both HTTP and HTTPS endpoints
 * 5. Monitoring connection pool health and metrics
 * 
 * Connection pooling reduces the overhead of establishing new TCP connections
 * for each request, which is particularly beneficial for the Graph API's high-frequency
 * request patterns in email, calendar, and contact operations.
 */

import http from 'http';
import https from 'https';
import { performance } from 'perf_hooks';
import { convertErrorToToolError, createServiceUnavailableError, createRateLimitError, createValidationError } from '../utils/mcpErrorResponse.js';

export class HttpConnectionPool {
  constructor(options = {}) {
    this.options = {
      // Connection pool settings
      maxSockets: options.maxSockets || 10,
      maxFreeSockets: options.maxFreeSockets || 5,
      timeout: options.timeout || 30000,
      keepAlive: options.keepAlive !== false,
      keepAliveMsecs: options.keepAliveMsecs || 1000,
      maxCachedSessions: options.maxCachedSessions || 100,
      
      // Security settings
      rejectUnauthorized: options.rejectUnauthorized !== false,
      secureProtocol: options.secureProtocol || 'TLSv1_2_method',
      
      // Performance settings
      family: options.family || 0, // IPv4 and IPv6
      hints: options.hints || 0,
      
      // Monitoring
      enableMetrics: options.enableMetrics !== false,
      ...options
    };
    
    // Create HTTP and HTTPS agents
    this.httpAgent = this.createHttpAgent();
    this.httpsAgent = this.createHttpsAgent();
    
    // Metrics tracking
    this.metrics = this.initializeMetrics();
    
    // Connection monitoring
    this.connectionMonitor = this.setupConnectionMonitor();
  }

  createHttpAgent() {
    const agent = new http.Agent({
      keepAlive: this.options.keepAlive,
      keepAliveMsecs: this.options.keepAliveMsecs,
      maxSockets: this.options.maxSockets,
      maxFreeSockets: this.options.maxFreeSockets,
      timeout: this.options.timeout,
      family: this.options.family,
      hints: this.options.hints
    });

    // Monitor agent events
    this.setupAgentMonitoring(agent, 'http');
    
    return agent;
  }

  createHttpsAgent() {
    const agent = new https.Agent({
      keepAlive: this.options.keepAlive,
      keepAliveMsecs: this.options.keepAliveMsecs,
      maxSockets: this.options.maxSockets,
      maxFreeSockets: this.options.maxFreeSockets,
      timeout: this.options.timeout,
      family: this.options.family,
      hints: this.options.hints,
      
      // TLS/SSL settings
      rejectUnauthorized: this.options.rejectUnauthorized,
      secureProtocol: this.options.secureProtocol,
      maxCachedSessions: this.options.maxCachedSessions,
      
      // Security enhancements
      ciphers: 'ECDHE-RSA-AES128-GCM-SHA256:ECDHE-RSA-AES256-GCM-SHA384',
      honorCipherOrder: true,
      secureOptions: require('constants').SSL_OP_NO_SSLv3 | require('constants').SSL_OP_NO_TLSv1
    });

    // Monitor agent events
    this.setupAgentMonitoring(agent, 'https');
    
    return agent;
  }

  setupAgentMonitoring(agent, protocol) {
    if (!this.options.enableMetrics) return;

    // Track socket creation
    agent.on('socket', (socket) => {
      this.metrics.socketsCreated++;
      this.metrics.activeSockets++;
      
      socket.on('close', () => {
        this.metrics.activeSockets--;
        this.metrics.socketsClosed++;
      });
      
      socket.on('error', (error) => {
        this.metrics.socketErrors++;
        console.warn(`Socket error (${protocol}):`, error.message);
        // Could return MCP error here if needed for socket-level errors
      });
      
      socket.on('timeout', () => {
        this.metrics.socketTimeouts++;
        console.warn(`Socket timeout (${protocol})`);
        // Could return MCP error here if needed for socket timeouts
      });
    });

    // Track connection reuse
    agent.on('socket', (socket) => {
      if (socket.readyState === 'open') {
        this.metrics.connectionsReused++;
      }
    });
  }

  initializeMetrics() {
    return {
      socketsCreated: 0,
      socketsClosed: 0,
      activeSockets: 0,
      socketErrors: 0,
      socketTimeouts: 0,
      connectionsReused: 0,
      requestsMade: 0,
      averageConnectionTime: 0,
      connectionTimes: [],
      poolUtilization: 0,
      lastUpdated: Date.now()
    };
  }

  setupConnectionMonitor() {
    return setInterval(() => {
      this.updateMetrics();
      this.cleanupStaleConnections();
    }, 30000); // Every 30 seconds
  }

  updateMetrics() {
    if (!this.options.enableMetrics) return;

    // Calculate pool utilization
    const totalSockets = this.options.maxSockets * 2; // HTTP + HTTPS
    this.metrics.poolUtilization = (this.metrics.activeSockets / totalSockets) * 100;
    
    // Calculate average connection time
    if (this.metrics.connectionTimes.length > 0) {
      const sum = this.metrics.connectionTimes.reduce((a, b) => a + b, 0);
      this.metrics.averageConnectionTime = sum / this.metrics.connectionTimes.length;
    }
    
    // Keep only last 100 connection times
    if (this.metrics.connectionTimes.length > 100) {
      this.metrics.connectionTimes = this.metrics.connectionTimes.slice(-100);
    }
    
    this.metrics.lastUpdated = Date.now();
  }

  cleanupStaleConnections() {
    // Force cleanup of stale connections
    const agents = [this.httpAgent, this.httpsAgent];
    
    agents.forEach(agent => {
      if (agent.sockets) {
        Object.values(agent.sockets).forEach(sockets => {
          sockets.forEach(socket => {
            if (socket.readyState === 'closed' || socket.destroyed) {
              socket.destroy();
            }
          });
        });
      }
    });
  }

  /**
   * Get the appropriate agent for a URL
   */
  getAgent(url) {
    const protocol = new URL(url).protocol;
    return protocol === 'https:' ? this.httpsAgent : this.httpAgent;
  }

  /**
   * Get axios configuration with connection pooling
   */
  getAxiosConfig(baseURL = 'https://graph.microsoft.com') {
    const url = new URL(baseURL);
    const agent = url.protocol === 'https:' ? this.httpsAgent : this.httpAgent;
    
    return {
      timeout: this.options.timeout,
      [url.protocol === 'https:' ? 'httpsAgent' : 'httpAgent']: agent,
      maxRedirects: 5,
      maxContentLength: 50 * 1024 * 1024, // 50MB
      maxBodyLength: 50 * 1024 * 1024, // 50MB
      
      // Connection and performance settings
      validateStatus: (status) => status < 500, // Don't throw on 4xx
      transformResponse: [(data) => {
        try {
          return JSON.parse(data);
        } catch (e) {
          return data;
        }
      }],
      
      // Headers for optimization
      headers: {
        'Connection': 'keep-alive',
        'Accept-Encoding': 'gzip, deflate, br',
        'User-Agent': 'outlook-mcp-server/1.0.0'
      }
    };
  }

  /**
   * Get fetch configuration with connection pooling
   */
  getFetchConfig(url) {
    const agent = this.getAgent(url);
    
    return {
      agent,
      timeout: this.options.timeout,
      headers: {
        'Connection': 'keep-alive',
        'Accept-Encoding': 'gzip, deflate, br',
        'User-Agent': 'outlook-mcp-server/1.0.0'
      }
    };
  }

  /**
   * Track request metrics
   */
  trackRequest(startTime) {
    if (!this.options.enableMetrics) return;

    const endTime = performance.now();
    const duration = endTime - startTime;
    
    this.metrics.requestsMade++;
    this.metrics.connectionTimes.push(duration);
  }

  /**
   * Get connection pool metrics
   */
  getMetrics() {
    return {
      ...this.metrics,
      httpAgent: this.getAgentStats(this.httpAgent),
      httpsAgent: this.getAgentStats(this.httpsAgent),
      config: {
        maxSockets: this.options.maxSockets,
        maxFreeSockets: this.options.maxFreeSockets,
        timeout: this.options.timeout,
        keepAlive: this.options.keepAlive
      }
    };
  }

  getAgentStats(agent) {
    const stats = {
      sockets: {},
      requests: {},
      totalSockets: 0,
      totalRequests: 0
    };

    // Count sockets by host
    if (agent.sockets) {
      Object.entries(agent.sockets).forEach(([host, sockets]) => {
        stats.sockets[host] = sockets.length;
        stats.totalSockets += sockets.length;
      });
    }

    // Count requests by host
    if (agent.requests) {
      Object.entries(agent.requests).forEach(([host, requests]) => {
        stats.requests[host] = requests.length;
        stats.totalRequests += requests.length;
      });
    }

    return stats;
  }

  /**
   * Health check for connection pool
   */
  async healthCheck() {
    const health = {
      healthy: true,
      issues: [],
      metrics: this.getMetrics()
    };

    // Check socket utilization
    if (this.metrics.poolUtilization > 90) {
      health.issues.push({
        level: 'warning',
        message: `High pool utilization: ${this.metrics.poolUtilization.toFixed(1)}%`,
        suggestion: 'Consider increasing maxSockets or reducing concurrent requests',
        mcpError: createRateLimitError(30) // Suggest waiting 30 seconds
      });
    }

    // Check error rates
    const errorRate = this.metrics.socketErrors / Math.max(this.metrics.socketsCreated, 1);
    if (errorRate > 0.1) {
      health.issues.push({
        level: 'error',
        message: `High socket error rate: ${(errorRate * 100).toFixed(1)}%`,
        suggestion: 'Check network connectivity and server health',
        mcpError: createServiceUnavailableError('Network connection pool')
      });
    }

    // Check timeout rates
    const timeoutRate = this.metrics.socketTimeouts / Math.max(this.metrics.socketsCreated, 1);
    if (timeoutRate > 0.05) {
      health.issues.push({
        level: 'warning',
        message: `High socket timeout rate: ${(timeoutRate * 100).toFixed(1)}%`,
        suggestion: 'Consider increasing timeout or checking network latency',
        mcpError: createServiceUnavailableError('Network connection (timeout issues)')
      });
    }

    // Check connection reuse
    const reuseRate = this.metrics.connectionsReused / Math.max(this.metrics.requestsMade, 1);
    if (reuseRate < 0.5) {
      health.issues.push({
        level: 'info',
        message: `Low connection reuse rate: ${(reuseRate * 100).toFixed(1)}%`,
        suggestion: 'Connection pooling may not be working optimally'
      });
    }

    health.healthy = health.issues.filter(issue => issue.level === 'error').length === 0;
    
    return health;
  }

  /**
   * Reset metrics
   */
  resetMetrics() {
    this.metrics = this.initializeMetrics();
  }

  /**
   * Cleanup resources
   */
  destroy() {
    // Clear monitoring interval
    if (this.connectionMonitor) {
      clearInterval(this.connectionMonitor);
    }

    // Destroy all sockets
    [this.httpAgent, this.httpsAgent].forEach(agent => {
      if (agent.sockets) {
        Object.values(agent.sockets).forEach(sockets => {
          sockets.forEach(socket => socket.destroy());
        });
      }
      agent.destroy();
    });
  }
}

/**
 * Factory function to create a connection pool optimized for Graph API
 */
export function createGraphConnectionPool(options = {}) {
  return new HttpConnectionPool({
    maxSockets: 10, // Conservative limit for Graph API
    maxFreeSockets: 5,
    timeout: 30000, // 30 seconds
    keepAlive: true,
    keepAliveMsecs: 1000,
    maxCachedSessions: 100,
    enableMetrics: true,
    ...options
  });
}

/**
 * Pre-configured connection pool for production use
 */
export const defaultGraphConnectionPool = createGraphConnectionPool();