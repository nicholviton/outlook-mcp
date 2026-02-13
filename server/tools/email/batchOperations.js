import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { BatchRequestHelper } from '../../graph/batchRequestHelper.js';

/**
 * Batch move emails by date range
 * Uses Microsoft Graph $batch API for efficient bulk operations
 */
export async function batchMoveEmailsTool(authManager, args) {
  const {
    sourceFolderId = 'inbox',
    destinationFolderId,
    startDate,
    endDate,
    limit = 100,
    dryRun = true
  } = args;

  // Validation
  if (!destinationFolderId) {
    return createValidationError('destinationFolderId', 'Parameter is required');
  }

  if (!startDate) {
    return createValidationError('startDate', 'Parameter is required (format: YYYY-MM-DDTHH:MM:SSZ)');
  }

  if (!endDate) {
    return createValidationError('endDate', 'Parameter is required (format: YYYY-MM-DDTHH:MM:SSZ)');
  }

  // Enforce limits for safety
  const effectiveLimit = Math.min(limit, 500); // Max 500 emails per operation
  if (limit > 500) {
    console.warn(`Limit ${limit} exceeds maximum of 500, capped at 500`);
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Step 1: Resolve source folder name to ID if needed
    let resolvedSourceFolderId = sourceFolderId;
    if (!sourceFolderId.match(/^[A-Za-z0-9\-_]+$/)) {
      // Looks like a folder name, not an ID - resolve it
      const folderResolver = graphApiClient.getFolderResolver();
      const resolved = await folderResolver.resolveFoldersToIds([sourceFolderId]);
      if (resolved.length === 0) {
        return createValidationError('sourceFolderId', `Folder '${sourceFolderId}' not found`);
      }
      resolvedSourceFolderId = resolved[0];
    }

    // Step 2: Resolve destination folder name to ID if needed
    let resolvedDestinationFolderId = destinationFolderId;
    if (!destinationFolderId.match(/^[A-Za-z0-9\-_]+$/)) {
      // Looks like a folder name, not an ID - resolve it
      const folderResolver = graphApiClient.getFolderResolver();
      const resolved = await folderResolver.resolveFoldersToIds([destinationFolderId]);
      if (resolved.length === 0) {
        return createValidationError('destinationFolderId', `Folder '${destinationFolderId}' not found`);
      }
      resolvedDestinationFolderId = resolved[0];
    }

    // Step 3: Search for emails matching the date range
    const searchEndpoint = `/me/mailFolders/${resolvedSourceFolderId}/messages`;
    const filterQuery = `receivedDateTime ge ${startDate} and receivedDateTime le ${endDate}`;

    const searchResult = await graphApiClient.makeRequest(searchEndpoint, {
      filter: filterQuery,
      select: 'id,subject,from,receivedDateTime',
      top: effectiveLimit,
      orderby: 'receivedDateTime desc'
    });

    // Check if search returned an MCP error
    if (searchResult.content && searchResult.isError !== undefined) {
      return searchResult;
    }

    const emails = searchResult.value || [];
    const foundCount = emails.length;

    if (foundCount === 0) {
      return {
        content: [{
          type: 'text',
          text: `No emails found in '${sourceFolderId}' between ${startDate} and ${endDate}.`
        }]
      };
    }

    // Step 4: If dry run, return summary without executing
    if (dryRun) {
      const emailSummary = emails.slice(0, 10).map(email => ({
        subject: email.subject,
        from: email.from?.emailAddress?.address || 'Unknown',
        receivedDateTime: email.receivedDateTime
      }));

      const summary = {
        dryRun: true,
        totalFound: foundCount,
        sourceFolderId: resolvedSourceFolderId,
        destinationFolderId: resolvedDestinationFolderId,
        dateRange: { startDate, endDate },
        limitApplied: effectiveLimit,
        sampleEmails: emailSummary,
        message: foundCount > 10 ? `Showing first 10 of ${foundCount} emails` : `All ${foundCount} emails shown`
      };

      return createSafeResponse(summary);
    }

    // Step 5: Execute batch move operation
    const messageIds = emails.map(e => e.id);
    const batchHelper = new BatchRequestHelper(graphApiClient);

    console.error(`Executing batch move of ${messageIds.length} emails...`);
    const batchResults = await batchHelper.executeMoveEmailsBatches(
      messageIds,
      resolvedDestinationFolderId
    );

    // Check if batch execution returned an MCP error
    if (batchResults.content && batchResults.isError !== undefined) {
      return batchResults;
    }

    // Step 6: Build response with results
    const successCount = batchResults.successful.length;
    const failureCount = batchResults.failed.length;

    let responseText = `Batch move completed:\n`;
    responseText += `✓ Successfully moved: ${successCount} emails\n`;
    responseText += `✗ Failed: ${failureCount} emails\n`;
    responseText += `Source: ${sourceFolderId}\n`;
    responseText += `Destination: ${destinationFolderId}\n`;
    responseText += `Date range: ${startDate} to ${endDate}`;

    if (failureCount > 0) {
      responseText += `\n\nFailed operations:\n`;
      batchResults.failed.slice(0, 5).forEach(failure => {
        responseText += `- ID ${failure.id}: ${failure.error.message}\n`;
      });
      if (failureCount > 5) {
        responseText += `... and ${failureCount - 5} more failures`;
      }
    }

    return {
      content: [{
        type: 'text',
        text: responseText
      }]
    };

  } catch (error) {
    return convertErrorToToolError(error, 'Failed to batch move emails');
  }
}
