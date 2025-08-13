// List emails from a specific folder
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

export async function listEmailsTool(authManager, args) {
  const { folder = 'inbox', limit = 10, filter } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'subject,from,receivedDateTime,bodyPreview,isRead',
      top: limit,
      orderby: 'receivedDateTime desc',
    };

    if (filter) {
      options.filter = filter;
    }

    const result = await graphApiClient.makeRequest(`/me/mailFolders/${folder}/messages`, options);

    const emails = result.value.map(email => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.emailAddress?.address || 'Unknown',
      fromName: email.from?.emailAddress?.name || 'Unknown',
      receivedDateTime: email.receivedDateTime,
      preview: email.bodyPreview,
      isRead: email.isRead,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ emails, count: emails.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list emails');
  }
}

// Get detailed information about a specific email
export async function getEmailTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,importance,isRead,hasAttachments,attachments,conversationId'
    };

    const email = await graphApiClient.makeRequest(`/me/messages/${messageId}`, options);

    const emailData = {
      id: email.id,
      subject: email.subject,
      from: {
        address: email.from?.emailAddress?.address || 'Unknown',
        name: email.from?.emailAddress?.name || 'Unknown'
      },
      toRecipients: email.toRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      ccRecipients: email.ccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      bccRecipients: email.bccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      receivedDateTime: email.receivedDateTime,
      sentDateTime: email.sentDateTime,
      body: {
        contentType: email.body?.contentType || 'Text',
        content: email.body?.content || ''
      },
      bodyPreview: email.bodyPreview,
      importance: email.importance,
      isRead: email.isRead,
      hasAttachments: email.hasAttachments,
      attachments: email.attachments?.map(a => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size
      })) || [],
      conversationId: email.conversationId
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(emailData, null, 2),
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get email');
  }
}