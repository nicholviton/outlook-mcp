import { createValidationError } from '../utils/mcpErrorResponse.js';

/**
 * Helper class for Microsoft Graph API $batch operations
 * Batch API allows up to 20 operations per request
 * https://learn.microsoft.com/en-us/graph/json-batching
 */
export class BatchRequestHelper {
  constructor(graphApiClient) {
    this.graphApiClient = graphApiClient;
    this.maxBatchSize = 20; // Microsoft Graph API limit
  }

  /**
   * Build a batch request for moving multiple emails
   * @param {Array<string>} messageIds - Array of message IDs to move
   * @param {string} destinationFolderId - Target folder ID
   * @returns {Object} Batch request object
   */
  buildMoveEmailsBatch(messageIds, destinationFolderId) {
    if (!Array.isArray(messageIds) || messageIds.length === 0) {
      throw new Error('messageIds must be a non-empty array');
    }

    if (messageIds.length > this.maxBatchSize) {
      throw new Error(`Batch size ${messageIds.length} exceeds maximum of ${this.maxBatchSize}`);
    }

    return {
      requests: messageIds.map((messageId, index) => ({
        id: String(index + 1),
        method: 'POST',
        url: `/me/messages/${messageId}/move`,
        body: {
          destinationId: destinationFolderId
        },
        headers: {
          'Content-Type': 'application/json'
        }
      }))
    };
  }

  /**
   * Execute a batch request and process responses
   * @param {Object} batchRequest - Batch request object
   * @returns {Promise<Object>} Object with successful and failed operations
   */
  async executeBatch(batchRequest) {
    try {
      const response = await this.graphApiClient.makeRequest('/$batch', {}, 'POST', batchRequest);

      // Check if this is an MCP error response
      if (response.content && response.isError !== undefined) {
        return response;
      }

      const results = {
        successful: [],
        failed: [],
        totalRequested: batchRequest.requests.length
      };

      // Process individual responses
      if (response.responses && Array.isArray(response.responses)) {
        response.responses.forEach((resp) => {
          if (resp.status >= 200 && resp.status < 300) {
            results.successful.push({
              id: resp.id,
              status: resp.status,
              body: resp.body
            });
          } else {
            results.failed.push({
              id: resp.id,
              status: resp.status,
              error: resp.body?.error || { message: 'Unknown error' }
            });
          }
        });
      }

      return results;
    } catch (error) {
      // If the entire batch request fails, return error info
      return {
        successful: [],
        failed: batchRequest.requests.map((req, idx) => ({
          id: String(idx + 1),
          status: 500,
          error: { message: error.message }
        })),
        totalRequested: batchRequest.requests.length,
        batchError: error.message
      };
    }
  }

  /**
   * Split a large array into chunks suitable for batch processing
   * @param {Array} items - Array to chunk
   * @param {number} chunkSize - Size of each chunk (default: maxBatchSize)
   * @returns {Array<Array>} Array of chunked arrays
   */
  chunkArray(items, chunkSize = this.maxBatchSize) {
    const chunks = [];
    for (let i = 0; i < items.length; i += chunkSize) {
      chunks.push(items.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Execute multiple batch requests in sequence (for >20 operations)
   * @param {Array<string>} messageIds - All message IDs to process
   * @param {string} destinationFolderId - Target folder ID
   * @returns {Promise<Object>} Aggregated results
   */
  async executeMoveEmailsBatches(messageIds, destinationFolderId) {
    const chunks = this.chunkArray(messageIds);
    const aggregatedResults = {
      successful: [],
      failed: [],
      totalRequested: messageIds.length,
      batchesProcessed: 0
    };

    for (const chunk of chunks) {
      const batchRequest = this.buildMoveEmailsBatch(chunk, destinationFolderId);
      const result = await this.executeBatch(batchRequest);

      // Check if result is an MCP error
      if (result.content && result.isError !== undefined) {
        // Return the error immediately
        return result;
      }

      aggregatedResults.successful.push(...result.successful);
      aggregatedResults.failed.push(...result.failed);
      aggregatedResults.batchesProcessed++;
    }

    return aggregatedResults;
  }
}
