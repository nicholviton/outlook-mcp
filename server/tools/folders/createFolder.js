import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Create mail folder
export async function createFolderTool(authManager, args) {
  const { displayName, parentFolderId } = args;

  if (!displayName) {
    return createValidationError('displayName', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const folderData = {
      displayName: displayName
    };

    let endpoint = '/me/mailFolders';
    if (parentFolderId) {
      endpoint = `/me/mailFolders/${parentFolderId}/childFolders`;
    }

    const result = await graphApiClient.postWithRetry(endpoint, folderData);

    return {
      content: [
        {
          type: 'text',
          text: `Folder "${displayName}" created successfully. Folder ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create folder');
  }
}