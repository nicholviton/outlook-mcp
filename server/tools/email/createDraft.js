import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Create draft email with user styling
export async function createDraftTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], importance = 'normal', preserveUserStyling = true } = args;

  if (!to || to.length === 0) {
    return createValidationError('to', 'At least one recipient is required');
  }

  if (!subject) {
    return createValidationError('subject', 'Subject is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Apply user styling if enabled
    let finalBody = body || '';
    let finalBodyType = bodyType;
    
    if (preserveUserStyling && finalBody) {
      const styledBody = await applyUserStyling(graphApiClient, finalBody, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const draft = {
      subject,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
      importance,
    };

    if (cc.length > 0) {
      draft.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      draft.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    const result = await graphApiClient.postWithRetry('/me/messages', draft);

    return {
      content: [
        {
          type: 'text',
          text: `Draft created successfully. Draft ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create draft');
  }
}