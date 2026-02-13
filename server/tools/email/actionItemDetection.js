import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

/**
 * Identify emails requiring user response based on multiple criteria
 * Prioritizes: flagged > unread from VIPs > keyword matches > recent unread
 */
export async function identifyActionItemsTool(authManager, args) {
  const {
    criteria = {},
    limit = 50
  } = args;

  const {
    unread = true,
    flagged = true,
    fromVIPs = [],
    keywords = ['urgent', 'action required', 'asap', 'follow up', 'deadline', 'please review'],
    daysOld = 30,
    folder = 'inbox'
  } = criteria;

  const effectiveLimit = Math.min(limit, 100); // Cap at 100 for performance

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Step 1: Resolve folder name to ID if needed
    let resolvedFolderId = folder;
    if (!folder.match(/^[A-Za-z0-9\-_]+$/)) {
      const folderResolver = graphApiClient.getFolderResolver();
      const resolved = await folderResolver.resolveFoldersToIds([folder]);
      if (resolved.length === 0) {
        return createValidationError('folder', `Folder '${folder}' not found`);
      }
      resolvedFolderId = resolved[0];
    }

    // Step 2: Build filter query
    const filterConditions = [];

    // Date filter - only emails from the last N days
    if (daysOld) {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - daysOld);
      const isoDate = cutoffDate.toISOString();
      filterConditions.push(`receivedDateTime ge ${isoDate}`);
    }

    // Build the filter query
    let filterQuery = filterConditions.join(' and ');

    // Step 3: Retrieve emails with all necessary fields
    const searchOptions = {
      select: 'id,subject,from,toRecipients,receivedDateTime,bodyPreview,importance,isRead,hasAttachments,flag',
      top: Math.min(effectiveLimit * 2, 200), // Get more than needed for filtering
      orderby: 'receivedDateTime desc'
    };

    if (filterQuery) {
      searchOptions.filter = filterQuery;
    }

    const searchEndpoint = `/me/mailFolders/${resolvedFolderId}/messages`;
    const result = await graphApiClient.makeRequest(searchEndpoint, searchOptions);

    // Check for MCP error
    if (result.content && result.isError !== undefined) {
      return result;
    }

    const emails = result.value || [];
    if (emails.length === 0) {
      return {
        content: [{
          type: 'text',
          text: `No emails found in '${folder}' from the last ${daysOld} days.`
        }]
      };
    }

    // Step 4: Score and filter emails based on criteria
    const scoredEmails = emails.map(email => {
      const emailData = {
        id: email.id,
        subject: email.subject,
        from: {
          address: email.from?.emailAddress?.address || 'Unknown',
          name: email.from?.emailAddress?.name || 'Unknown'
        },
        receivedDateTime: email.receivedDateTime,
        bodyPreview: email.bodyPreview,
        importance: email.importance,
        isRead: email.isRead,
        hasAttachments: email.hasAttachments,
        isFlagged: email.flag?.flagStatus === 'flagged',
        score: 0,
        reasons: []
      };

      // Scoring logic
      // Highest priority: Flagged emails
      if (flagged && emailData.isFlagged) {
        emailData.score += 100;
        emailData.reasons.push('Flagged');
      }

      // High priority: Unread emails from VIPs
      if (unread && !emailData.isRead) {
        emailData.score += 50;
        emailData.reasons.push('Unread');
      }

      // Check if from VIP sender
      const isFromVIP = fromVIPs.length > 0 &&
        fromVIPs.some(vip =>
          emailData.from.address.toLowerCase().includes(vip.toLowerCase())
        );

      if (isFromVIP) {
        emailData.score += 75;
        emailData.reasons.push('From VIP');
      }

      // Check for action keywords in subject
      const subject = (email.subject || '').toLowerCase();
      const matchedKeywords = keywords.filter(keyword =>
        subject.includes(keyword.toLowerCase())
      );

      if (matchedKeywords.length > 0) {
        emailData.score += 30 * matchedKeywords.length;
        emailData.reasons.push(`Keywords: ${matchedKeywords.join(', ')}`);
      }

      // Bonus for high importance
      if (emailData.importance === 'high') {
        emailData.score += 20;
        emailData.reasons.push('High importance');
      }

      // Bonus for recent emails (within last 3 days)
      const daysSinceReceived = (Date.now() - new Date(email.receivedDateTime).getTime()) / (1000 * 60 * 60 * 24);
      if (daysSinceReceived <= 3) {
        emailData.score += 15;
        emailData.reasons.push('Recent');
      }

      return emailData;
    });

    // Step 5: Filter emails with score > 0 and sort by score
    const actionItems = scoredEmails
      .filter(email => email.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, effectiveLimit);

    if (actionItems.length === 0) {
      return {
        content: [{
          type: 'text',
          text: `No action items found matching the specified criteria in '${folder}'.`
        }]
      };
    }

    // Step 6: Generate summary statistics
    const stats = {
      totalActionItems: actionItems.length,
      flaggedCount: actionItems.filter(e => e.isFlagged).length,
      unreadCount: actionItems.filter(e => !e.isRead).length,
      fromVIPsCount: actionItems.filter(e => e.reasons.includes('From VIP')).length,
      withKeywordsCount: actionItems.filter(e => e.reasons.some(r => r.startsWith('Keywords:'))).length,
      highImportanceCount: actionItems.filter(e => e.importance === 'high').length
    };

    const response = {
      folder,
      criteria: {
        unread,
        flagged,
        fromVIPs: fromVIPs.length > 0 ? fromVIPs : 'None specified',
        keywords,
        daysOld
      },
      statistics: stats,
      actionItems: actionItems.map(email => ({
        id: email.id,
        subject: email.subject,
        from: email.from,
        receivedDateTime: email.receivedDateTime,
        bodyPreview: email.bodyPreview?.substring(0, 150) + (email.bodyPreview?.length > 150 ? '...' : ''),
        isRead: email.isRead,
        isFlagged: email.isFlagged,
        importance: email.importance,
        hasAttachments: email.hasAttachments,
        priorityScore: email.score,
        reasons: email.reasons
      }))
    };

    return createSafeResponse(response);

  } catch (error) {
    return convertErrorToToolError(error, 'Failed to identify action items');
  }
}
