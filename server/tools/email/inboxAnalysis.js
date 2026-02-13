import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

/**
 * Analyze inbox and provide categorization insights
 * Retrieves metadata only (no body content) for performance
 */
export async function analyzeInboxTool(authManager, args) {
  const {
    folder = 'inbox',
    analysisType = 'categories',
    startDate,
    endDate,
    limit = 100
  } = args;

  // Validation
  const validAnalysisTypes = ['categories', 'senders', 'time-patterns', 'action-items'];
  if (!validAnalysisTypes.includes(analysisType)) {
    return createValidationError(
      'analysisType',
      `Must be one of: ${validAnalysisTypes.join(', ')}`
    );
  }

  const effectiveLimit = Math.min(limit, 1000); // Cap at 1000 for performance

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

    // Step 2: Build filter query based on date range
    const filterConditions = [];
    if (startDate) {
      filterConditions.push(`receivedDateTime ge ${startDate}`);
    }
    if (endDate) {
      filterConditions.push(`receivedDateTime le ${endDate}`);
    }

    const searchOptions = {
      select: 'id,subject,from,toRecipients,receivedDateTime,importance,isRead,hasAttachments,flag,categories',
      top: effectiveLimit,
      orderby: 'receivedDateTime desc'
    };

    if (filterConditions.length > 0) {
      searchOptions.filter = filterConditions.join(' and ');
    }

    // Step 3: Retrieve emails (metadata only)
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
          text: `No emails found in '${folder}'${startDate ? ` from ${startDate}` : ''}${endDate ? ` to ${endDate}` : ''}.`
        }]
      };
    }

    // Step 4: Perform analysis based on type
    let analysis;
    switch (analysisType) {
      case 'categories':
        analysis = analyzeBySenderDomains(emails);
        break;
      case 'senders':
        analysis = analyzeTopSenders(emails);
        break;
      case 'time-patterns':
        analysis = analyzeTimePatterns(emails);
        break;
      case 'action-items':
        analysis = analyzeActionItems(emails);
        break;
    }

    const response = {
      folder,
      analysisType,
      totalEmails: emails.length,
      dateRange: {
        startDate: startDate || 'N/A',
        endDate: endDate || 'N/A'
      },
      limitApplied: effectiveLimit,
      analysis
    };

    return createSafeResponse(response);

  } catch (error) {
    return convertErrorToToolError(error, 'Failed to analyze inbox');
  }
}

/**
 * Analyze emails by sender domains and categories
 */
function analyzeBySenderDomains(emails) {
  const domainCounts = {};
  const subjectKeywords = {};
  const categoryCounts = {};

  emails.forEach(email => {
    // Count sender domains
    const senderAddress = email.from?.emailAddress?.address || 'unknown';
    const domain = senderAddress.split('@')[1] || 'unknown';
    domainCounts[domain] = (domainCounts[domain] || 0) + 1;

    // Extract subject keywords (words 4+ characters)
    const subject = email.subject || '';
    const words = subject.toLowerCase()
      .split(/\s+/)
      .filter(word => word.length >= 4 && !word.match(/^(this|that|with|from|your|have|been|will)$/));

    words.forEach(word => {
      subjectKeywords[word] = (subjectKeywords[word] || 0) + 1;
    });

    // Count categories
    if (email.categories && email.categories.length > 0) {
      email.categories.forEach(cat => {
        categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
      });
    }
  });

  // Sort and get top results
  const topDomains = Object.entries(domainCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([domain, count]) => ({ domain, count, percentage: ((count / emails.length) * 100).toFixed(1) }));

  const topKeywords = Object.entries(subjectKeywords)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15)
    .map(([keyword, count]) => ({ keyword, count }));

  const topCategories = Object.entries(categoryCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([category, count]) => ({ category, count }));

  return {
    topSenderDomains: topDomains,
    commonSubjectKeywords: topKeywords,
    categories: topCategories.length > 0 ? topCategories : 'No categories assigned',
    insights: {
      totalUniqueDomains: Object.keys(domainCounts).length,
      totalUniqueKeywords: Object.keys(subjectKeywords).length,
      categorizedEmails: topCategories.reduce((sum, cat) => sum + cat.count, 0)
    }
  };
}

/**
 * Analyze top email senders by volume
 */
function analyzeTopSenders(emails) {
  const senderCounts = {};
  const senderDetails = {};

  emails.forEach(email => {
    const senderAddress = email.from?.emailAddress?.address || 'unknown';
    const senderName = email.from?.emailAddress?.name || senderAddress;

    senderCounts[senderAddress] = (senderCounts[senderAddress] || 0) + 1;

    if (!senderDetails[senderAddress]) {
      senderDetails[senderAddress] = {
        name: senderName,
        address: senderAddress,
        unreadCount: 0,
        flaggedCount: 0,
        withAttachments: 0
      };
    }

    if (!email.isRead) senderDetails[senderAddress].unreadCount++;
    if (email.flag?.flagStatus === 'flagged') senderDetails[senderAddress].flaggedCount++;
    if (email.hasAttachments) senderDetails[senderAddress].withAttachments++;
  });

  const topSenders = Object.entries(senderCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 20)
    .map(([address, count]) => ({
      ...senderDetails[address],
      totalEmails: count,
      percentage: ((count / emails.length) * 100).toFixed(1)
    }));

  return {
    topSenders,
    insights: {
      totalUniqueSenders: Object.keys(senderCounts).length,
      averageEmailsPerSender: (emails.length / Object.keys(senderCounts).length).toFixed(1)
    }
  };
}

/**
 * Analyze email time patterns (day of week, hour of day)
 */
function analyzeTimePatterns(emails) {
  const dayOfWeekCounts = { Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0, Sun: 0 };
  const hourCounts = Array(24).fill(0);
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  emails.forEach(email => {
    if (!email.receivedDateTime) return;

    const date = new Date(email.receivedDateTime);
    const dayName = dayNames[date.getDay()];
    const hour = date.getHours();

    dayOfWeekCounts[dayName]++;
    hourCounts[hour]++;
  });

  // Find peak hours
  const peakHours = hourCounts
    .map((count, hour) => ({ hour, count }))
    .filter(item => item.count > 0)
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  // Calculate business vs non-business hours
  const businessHours = hourCounts.slice(9, 17).reduce((sum, count) => sum + count, 0);
  const nonBusinessHours = emails.length - businessHours;

  return {
    emailsByDayOfWeek: Object.entries(dayOfWeekCounts).map(([day, count]) => ({
      day,
      count,
      percentage: ((count / emails.length) * 100).toFixed(1)
    })),
    peakHours: peakHours.map(({ hour, count }) => ({
      hour: `${hour}:00-${hour + 1}:00`,
      count,
      percentage: ((count / emails.length) * 100).toFixed(1)
    })),
    businessVsNonBusiness: {
      businessHours: { count: businessHours, percentage: ((businessHours / emails.length) * 100).toFixed(1) },
      nonBusinessHours: { count: nonBusinessHours, percentage: ((nonBusinessHours / emails.length) * 100).toFixed(1) }
    }
  };
}

/**
 * Analyze emails for action items
 */
function analyzeActionItems(emails) {
  const unreadEmails = emails.filter(e => !e.isRead);
  const flaggedEmails = emails.filter(e => e.flag?.flagStatus === 'flagged');
  const importantEmails = emails.filter(e => e.importance === 'high');

  // Count emails with action-related keywords in subjects
  const actionKeywords = ['action required', 'urgent', 'asap', 'deadline', 'follow up', 'please review', 'approval needed'];
  const actionKeywordEmails = emails.filter(email => {
    const subject = (email.subject || '').toLowerCase();
    return actionKeywords.some(keyword => subject.includes(keyword));
  });

  return {
    unreadEmails: {
      count: unreadEmails.length,
      percentage: ((unreadEmails.length / emails.length) * 100).toFixed(1)
    },
    flaggedEmails: {
      count: flaggedEmails.length,
      percentage: ((flaggedEmails.length / emails.length) * 100).toFixed(1)
    },
    importantEmails: {
      count: importantEmails.length,
      percentage: ((importantEmails.length / emails.length) * 100).toFixed(1)
    },
    actionKeywordEmails: {
      count: actionKeywordEmails.length,
      percentage: ((actionKeywordEmails.length / emails.length) * 100).toFixed(1),
      keywords: actionKeywords
    },
    summary: {
      totalRequiringAttention: new Set([
        ...unreadEmails.map(e => e.id),
        ...flaggedEmails.map(e => e.id),
        ...importantEmails.map(e => e.id),
        ...actionKeywordEmails.map(e => e.id)
      ]).size
    }
  };
}
