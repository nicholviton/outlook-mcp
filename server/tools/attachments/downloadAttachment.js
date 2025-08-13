import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Helper function to format file size
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Download attachment
export async function downloadAttachmentTool(authManager, args) {
  const { messageId, attachmentId, includeContent = false } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!attachmentId) {
    return createValidationError('attachmentId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const selectFields = 'id,name,contentType,size,isInline,lastModifiedDateTime';

    const attachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
      select: selectFields
    });

    const attachmentInfo = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime
    };

    if (includeContent) {
      try {
        const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
          select: selectFields + ',contentBytes'
        });

        // Only include content for FileAttachment type
        if (fullAttachment['@odata.type'] === '#microsoft.graph.fileAttachment' && fullAttachment.contentBytes) {
          attachmentInfo.contentBytes = fullAttachment.contentBytes;
          attachmentInfo.contentIncluded = true;
        } else {
          attachmentInfo.contentIncluded = false;
          attachmentInfo.contentError = 'Content not available or not a file attachment';
        }
      } catch (error) {
        attachmentInfo.contentIncluded = false;
        attachmentInfo.contentError = error.message;
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(attachmentInfo, null, 2),
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to download attachment');
  }
}