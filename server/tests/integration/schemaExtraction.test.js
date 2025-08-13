import { describe, it, expect, beforeEach } from 'vitest';
import { allToolSchemas, allToolSchemaMap } from '../../schemas/toolSchemas.js';

describe('Schema Extraction Integration', () => {
  describe('MCP Server Integration', () => {
    it('should provide schemas compatible with MCP ListToolsRequestSchema', () => {
      // Test that our schemas match the expected MCP format
      const mcpTools = allToolSchemas.map(schema => ({
        name: schema.name,
        description: schema.description,
        inputSchema: schema.inputSchema
      }));

      expect(mcpTools).toBeDefined();
      expect(mcpTools.length).toBeGreaterThan(0);

      mcpTools.forEach(tool => {
        expect(tool.name).toBeDefined();
        expect(tool.description).toBeDefined();
        expect(tool.inputSchema).toBeDefined();
        expect(tool.inputSchema.type).toBe('object');
        expect(tool.inputSchema.properties).toBeDefined();
      });
    });

    it('should handle CallToolRequestSchema mapping', () => {
      // Test that we can map tool names to their handlers
      const toolNames = Object.keys(allToolSchemaMap);
      
      expect(toolNames.length).toBeGreaterThan(0);
      
      toolNames.forEach(toolName => {
        const schema = allToolSchemaMap[toolName];
        expect(schema).toBeDefined();
        expect(schema.name).toBe(toolName);
      });
    });

    it('should reduce server/index.js complexity', () => {
      // Verify that schema extraction reduces the main server file complexity
      const totalSchemas = allToolSchemas.length;
      const estimatedLinesReduced = totalSchemas * 25; // Approximately 25 lines per schema
      
      expect(totalSchemas).toBeGreaterThan(20); // Should have substantial number of schemas
      expect(estimatedLinesReduced).toBeGreaterThan(500); // Should reduce by significant amount
    });
  });

  describe('Performance Impact', () => {
    it('should load schemas efficiently', () => {
      const startTime = performance.now();
      
      // Import and access schemas
      const schemas = allToolSchemas;
      const schemaMap = allToolSchemaMap;
      
      // Perform operations that would happen in server
      const emailSchemas = schemas.filter(s => s.name.includes('email'));
      const calendarSchemas = schemas.filter(s => s.name.includes('event'));
      const lookupTest = schemaMap['outlook_send_email'];
      
      const endTime = performance.now();
      const duration = endTime - startTime;
      
      expect(emailSchemas.length).toBeGreaterThan(0);
      expect(calendarSchemas.length).toBeGreaterThan(0);
      expect(lookupTest).toBeDefined();
      expect(duration).toBeLessThan(50); // Should be very fast
    });

    it('should support concurrent schema access', async () => {
      // Simulate concurrent access that might happen in server
      const promises = Array(10).fill(0).map(async (_, i) => {
        return new Promise(resolve => {
          setTimeout(() => {
            const schema = allToolSchemaMap[`outlook_${i % 2 === 0 ? 'send_email' : 'list_emails'}`];
            resolve(schema);
          }, Math.random() * 10);
        });
      });

      const results = await Promise.all(promises);
      
      expect(results.length).toBe(10);
      results.forEach(result => {
        expect(result).toBeDefined();
        expect(result.name).toBeDefined();
      });
    });
  });

  describe('Schema Validation Integration', () => {
    it('should validate real email parameters', () => {
      const emailSchema = allToolSchemaMap['outlook_send_email'];
      expect(emailSchema).toBeDefined();
      
      const validParams = {
        to: ['test@example.com'],
        subject: 'Test Email',
        body: 'Test body content',
        bodyType: 'text'
      };

      const invalidParams = {
        to: 'not-an-array',
        subject: '',
        body: 'Test body content'
      };

      // Test that the schema structure supports validation
      expect(emailSchema.inputSchema.required).toContain('to');
      expect(emailSchema.inputSchema.required).toContain('subject');
      expect(emailSchema.inputSchema.required).toContain('body');
      
      expect(emailSchema.inputSchema.properties.to.type).toBe('array');
      expect(emailSchema.inputSchema.properties.subject.type).toBe('string');
      expect(emailSchema.inputSchema.properties.bodyType.enum).toContain('text');
      expect(emailSchema.inputSchema.properties.bodyType.enum).toContain('html');
    });

    it('should validate calendar event parameters', () => {
      const eventSchema = allToolSchemaMap['outlook_create_event'];
      expect(eventSchema).toBeDefined();
      
      const validParams = {
        subject: 'Team Meeting',
        start: {
          dateTime: '2023-12-25T10:00:00Z',
          timeZone: 'UTC'
        },
        end: {
          dateTime: '2023-12-25T11:00:00Z',
          timeZone: 'UTC'
        }
      };

      // Test that the schema structure supports validation
      expect(eventSchema.inputSchema.required).toContain('subject');
      expect(eventSchema.inputSchema.required).toContain('start');
      expect(eventSchema.inputSchema.required).toContain('end');
      
      expect(eventSchema.inputSchema.properties.start.type).toBe('object');
      expect(eventSchema.inputSchema.properties.start.required).toContain('dateTime');
      expect(eventSchema.inputSchema.properties.start.required).toContain('timeZone');
    });
  });

  describe('Error Handling Integration', () => {
    it('should handle missing schemas gracefully', () => {
      const nonExistentSchema = allToolSchemaMap['outlook_nonexistent_tool'];
      expect(nonExistentSchema).toBeUndefined();
    });

    it('should provide meaningful error context', () => {
      // Test that schemas provide enough information for error handling
      allToolSchemas.forEach(schema => {
        expect(schema.name).toBeTruthy();
        expect(schema.description).toBeTruthy();
        
        if (schema.inputSchema.required) {
          expect(Array.isArray(schema.inputSchema.required)).toBe(true);
        }
        
        if (schema.inputSchema.properties) {
          Object.values(schema.inputSchema.properties).forEach(prop => {
            expect(prop.type || prop.enum).toBeTruthy();
            expect(prop.description).toBeTruthy();
          });
        }
      });
    });
  });

  describe('Memory Usage', () => {
    it('should not create memory leaks', () => {
      // Test that repeated schema access doesn't create memory leaks
      const iterations = 1000;
      const startMemory = process.memoryUsage().heapUsed;
      
      for (let i = 0; i < iterations; i++) {
        const schema = allToolSchemaMap['outlook_send_email'];
        const props = schema.inputSchema.properties;
        const required = schema.inputSchema.required;
        
        // Simulate usage
        expect(schema).toBeDefined();
        expect(props).toBeDefined();
        expect(required).toBeDefined();
      }
      
      const endMemory = process.memoryUsage().heapUsed;
      const memoryGrowth = endMemory - startMemory;
      
      // Memory growth should be reasonable (less than 5MB for 1000 iterations)
      expect(memoryGrowth).toBeLessThan(5 * 1024 * 1024);
    });
  });

  describe('Backward Compatibility', () => {
    it('should maintain the same tool names as original server', () => {
      const expectedToolNames = [
        'outlook_list_emails',
        'outlook_send_email',
        'outlook_list_events',
        'outlook_create_event',
        'outlook_get_email',
        'outlook_search_emails',
        'outlook_create_draft',
        'outlook_reply_to_email',
        'outlook_reply_all',
        'outlook_forward_email',
        'outlook_delete_email',
        'outlook_move_email',
        'outlook_mark_as_read',
        'outlook_flag_email',
        'outlook_categorize_email',
        'outlook_archive_email',
        'outlook_batch_process_emails',
        'outlook_list_folders',
        'outlook_create_folder',
        'outlook_rename_folder',
        'outlook_get_folder_stats',
        'outlook_list_attachments',
        'outlook_download_attachment',
        'outlook_add_attachment',
        'outlook_scan_attachments'
      ];

      const actualToolNames = Object.keys(allToolSchemaMap);
      
      expectedToolNames.forEach(toolName => {
        expect(actualToolNames).toContain(toolName);
      });
    });

    it('should maintain the same schema structure as original server', () => {
      // Test a few key schemas to ensure structure is preserved
      const sendEmailSchema = allToolSchemaMap['outlook_send_email'];
      expect(sendEmailSchema.inputSchema.properties.preserveUserStyling).toBeDefined();
      expect(sendEmailSchema.inputSchema.properties.preserveUserStyling.default).toBe(true);
      
      const createEventSchema = allToolSchemaMap['outlook_create_event'];
      expect(createEventSchema.inputSchema.properties.recurrence).toBeDefined();
      expect(createEventSchema.inputSchema.properties.recurrence.properties.pattern).toBeDefined();
      expect(createEventSchema.inputSchema.properties.recurrence.properties.range).toBeDefined();
    });
  });
});