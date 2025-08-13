#!/usr/bin/env node

// Add global error handlers for debugging
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  console.error('Stack trace:', reason.stack || 'No stack trace available');
  // Don't exit immediately - let the MCP server handle errors gracefully
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack trace:', error.stack);
  // Only exit on truly fatal errors
  if (error.code === 'MODULE_NOT_FOUND' || error.name === 'SyntaxError') {
    process.exit(1);
  }
});

// Main initialization function using IIFE to handle async imports
(async function initializeServer() {
  console.error('Debug: Script starting...');
  
  try {
    console.error('Debug: Loading dotenv...');
    await import('dotenv/config');
    
    console.error('Debug: Loading MCP SDK...');
    const { Server } = await import('@modelcontextprotocol/sdk/server/index.js');
    const { StdioServerTransport } = await import('@modelcontextprotocol/sdk/server/stdio.js');
    const { ListToolsRequestSchema, CallToolRequestSchema, InitializeRequestSchema, InitializedNotificationSchema } = await import('@modelcontextprotocol/sdk/types.js');
    
    console.error('Debug: Loading auth manager...');
    const { OutlookAuthManager } = await import('./auth/auth.js');
    
    console.error('Debug: Loading MCP error utilities...');
    const { createToolError, createProtocolError, ErrorCodes, convertErrorToToolError } = await import('./utils/mcpErrorResponse.js');
    
    console.error('Debug: Loading tools...');
    const tools = await import('./tools/index.js');
    console.error('Debug: Tools imported, available:', Object.keys(tools).length);
    
    // Extract the specific tools we need
    const {
      listEmailsTool,
      sendEmailTool,
      listEventsTool,
      createEventTool,
      getEmailTool,
      searchEmailsTool,
      createDraftTool,
      replyToEmailTool,
      replyAllTool,
      forwardEmailTool,
      deleteEmailTool,
      // Email Management Tools
      moveEmailTool,
      markAsReadTool,
      flagEmailTool,
      categorizeEmailTool,
      archiveEmailTool,
      batchProcessEmailsTool,
      // Folder Management Tools
      listFoldersTool,
      createFolderTool,
      renameFolderTool,
      getFolderStatsTool,
      // Attachment Tools
      listAttachmentsTool,
      downloadAttachmentTool,
      addAttachmentTool,
      scanAttachmentsTool
    } = tools;
    
    console.error('Debug: All required tools extracted successfully');
    console.error('Debug: All imports successful');

    const server = new Server(
      {
        name: 'outlook-mcp',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    const authManager = new OutlookAuthManager(
      process.env.AZURE_CLIENT_ID,
      process.env.AZURE_TENANT_ID
    );

    server.setRequestHandler(InitializeRequestSchema, async (request) => {
      console.error('Debug: Handling MCP initialization...');
      console.error('Debug: Initialize request:', JSON.stringify(request, null, 2));
      
      const response = {
        protocolVersion: '2025-06-18',
        capabilities: {
          tools: {},
        },
        serverInfo: {
          name: 'outlook-mcp',
          version: '1.0.0',
        },
      };
      
      console.error('Debug: Initialize response:', JSON.stringify(response, null, 2));
      return response;
    });

    server.setNotificationHandler(InitializedNotificationSchema, async () => {
      console.error('Debug: Client initialized');
    });

    server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'outlook_list_emails',
            description: 'List emails from Outlook inbox or specified folder',
            inputSchema: {
              type: 'object',
              properties: {
                folder: {
                  type: 'string',
                  description: 'Folder to list emails from (default: inbox)',
                  default: 'inbox',
                },
                limit: {
                  type: 'number',
                  description: 'Maximum number of emails to return',
                  default: 10,
                },
                filter: {
                  type: 'string',
                  description: 'OData filter query for emails',
                },
              },
            },
          },
          {
            name: 'outlook_send_email',
            description: 'Send an email through Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Recipient email addresses',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject',
                },
                body: {
                  type: 'string',
                  description: 'Email body content',
                },
                bodyType: {
                  type: 'string',
                  enum: ['text', 'html'],
                  default: 'text',
                  description: 'Body content type',
                },
                cc: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'CC recipients',
                },
                bcc: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'BCC recipients',
                },
                preserveUserStyling: {
                  type: 'boolean',
                  description: 'Apply user\'s default Outlook styling, font preferences, and signature',
                  default: true,
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
          {
            name: 'outlook_list_events',
            description: 'List calendar events from Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                startDateTime: {
                  type: 'string',
                  description: 'Start date/time in ISO 8601 format',
                },
                endDateTime: {
                  type: 'string',
                  description: 'End date/time in ISO 8601 format',
                },
                limit: {
                  type: 'number',
                  description: 'Maximum number of events to return',
                  default: 10,
                },
                calendar: {
                  type: 'string',
                  description: 'Calendar ID (default: primary calendar)',
                },
              },
            },
          },
          {
            name: 'outlook_create_event',
            description: 'Create a new calendar event in Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Event subject/title',
                },
                start: {
                  type: 'object',
                  properties: {
                    dateTime: {
                      type: 'string',
                      description: 'Start date/time in ISO 8601 format',
                    },
                    timeZone: {
                      type: 'string',
                      description: 'Time zone (e.g., "Pacific Standard Time")',
                    },
                  },
                  required: ['dateTime', 'timeZone'],
                },
                end: {
                  type: 'object',
                  properties: {
                    dateTime: {
                      type: 'string',
                      description: 'End date/time in ISO 8601 format',
                    },
                    timeZone: {
                      type: 'string',
                      description: 'Time zone (e.g., "Pacific Standard Time")',
                    },
                  },
                  required: ['dateTime', 'timeZone'],
                },
                body: {
                  type: 'string',
                  description: 'Event description',
                },
                location: {
                  type: 'string',
                  description: 'Event location',
                },
                attendees: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Attendee email addresses',
                },
                isOnlineMeeting: {
                  type: 'boolean',
                  description: 'Whether to create this as a Teams meeting (default: false)',
                },
              },
              required: ['subject', 'start', 'end'],
            },
          },
          {
            name: 'outlook_get_email',
            description: 'Get detailed information about a specific email',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the email message to retrieve',
                },
              },
              required: ['messageId'],
            },
          },
          {
            name: 'outlook_search_emails',
            description: 'Search emails across all folders with advanced filters',
            inputSchema: {
              type: 'object',
              properties: {
                query: {
                  type: 'string',
                  description: 'Free-text search query across email content',
                },
                subject: {
                  type: 'string',
                  description: 'Search emails with specific subject text',
                },
                from: {
                  type: 'string',
                  description: 'Filter emails from specific sender',
                },
                startDate: {
                  type: 'string',
                  description: 'Start date for email search (ISO 8601 format)',
                },
                endDate: {
                  type: 'string',
                  description: 'End date for email search (ISO 8601 format)',
                },
                folders: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Specific folders to search in',
                },
                limit: {
                  type: 'number',
                  description: 'Maximum number of emails to return',
                  default: 100,
                },
              },
            },
          },
          {
            name: 'outlook_list_folders',
            description: 'List all email folders',
            inputSchema: {
              type: 'object',
              properties: {
                includeHidden: {
                  type: 'boolean',
                  description: 'Include hidden folders',
                  default: false,
                },
                includeChildFolders: {
                  type: 'boolean',
                  description: 'Include nested child folders',
                  default: true,
                },
                top: {
                  type: 'number',
                  description: 'Maximum number of folders to return',
                  default: 100,
                },
              },
            },
          },
        ],
      };
    });

    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'outlook_list_emails':
            return await listEmailsTool(authManager, args);
          
          case 'outlook_send_email':
            return await sendEmailTool(authManager, args);
          
          case 'outlook_list_events':
            return await listEventsTool(authManager, args);
          
          case 'outlook_create_event':
            return await createEventTool(authManager, args);
          
          case 'outlook_get_email':
            return await getEmailTool(authManager, args);
          
          case 'outlook_search_emails':
            return await searchEmailsTool(authManager, args);
          
          case 'outlook_create_draft':
            return await createDraftTool(authManager, args);
          
          case 'outlook_reply_to_email':
            return await replyToEmailTool(authManager, args);
          
          case 'outlook_reply_all':
            return await replyAllTool(authManager, args);
          
          case 'outlook_forward_email':
            return await forwardEmailTool(authManager, args);
          
          case 'outlook_delete_email':
            return await deleteEmailTool(authManager, args);
          
          case 'outlook_move_email':
            return await moveEmailTool(authManager, args);
          
          case 'outlook_mark_as_read':
            return await markAsReadTool(authManager, args);
          
          case 'outlook_flag_email':
            return await flagEmailTool(authManager, args);
          
          case 'outlook_categorize_email':
            return await categorizeEmailTool(authManager, args);
          
          case 'outlook_archive_email':
            return await archiveEmailTool(authManager, args);
          
          case 'outlook_batch_process_emails':
            return await batchProcessEmailsTool(authManager, args);
          
          case 'outlook_list_folders':
            return await listFoldersTool(authManager, args);
          
          case 'outlook_create_folder':
            return await createFolderTool(authManager, args);
          
          case 'outlook_rename_folder':
            return await renameFolderTool(authManager, args);
          
          case 'outlook_get_folder_stats':
            return await getFolderStatsTool(authManager, args);
          
          case 'outlook_list_attachments':
            return await listAttachmentsTool(authManager, args);
          
          case 'outlook_download_attachment':
            return await downloadAttachmentTool(authManager, args);
          
          case 'outlook_add_attachment':
            return await addAttachmentTool(authManager, args);
          
          case 'outlook_scan_attachments':
            return await scanAttachmentsTool(authManager, args);
          
          default:
            return createProtocolError(
              ErrorCodes.METHOD_NOT_FOUND,
              `Unknown tool: ${name}`,
              { availableTools: Object.keys(tools).filter(key => key.endsWith('Tool')) }
            );
        }
      } catch (error) {
        console.error('Unexpected error in tool handler:', error);
        
        // If it's already an MCP error response, return it as-is
        if (error.content && error.isError !== undefined) {
          return error;
        }
        
        // Convert other errors to MCP tool error format
        return convertErrorToToolError(error, 'Tool execution failed');
      }
    });

    // Start the server
    async function main() {
      console.error('Debug: Starting main function...');
      console.error(`Debug: AZURE_CLIENT_ID = ${process.env.AZURE_CLIENT_ID ? 'SET' : 'NOT SET'}`);
      console.error(`Debug: AZURE_TENANT_ID = ${process.env.AZURE_TENANT_ID ? 'SET' : 'NOT SET'}`);
      
      if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_TENANT_ID) {
        console.error('Error: AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables are required.');
        console.error('Please set these in your MCP server configuration.');
        console.error('Note: This server uses OAuth 2.0 with PKCE for secure delegated authentication.');
        process.exit(1);
      }

      console.error('Starting Outlook MCP Server...');
      console.error('Authentication will be performed when first tool is called.');
      
      try {
        console.error('Debug: Creating StdioServerTransport...');
        const transport = new StdioServerTransport();
        
        console.error('Debug: Connecting server to transport...');
        await server.connect(transport);
        
        console.error('Outlook MCP server is ready and connected');
        
      } catch (error) {
        console.error('Error during server connection:', error);
        throw error;
      }
    }

    await main();
    
  } catch (error) {
    console.error('Server error:', error);
    process.exit(1);
  }
})();
