/**
 * Calendar-related MCP tool schemas
 * 
 * This module contains all JSON schemas for calendar operations in the Outlook MCP server.
 * Includes support for event creation, recurring meetings, and Teams integration.
 */

export const listEventsSchema = {
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
};

export const createEventSchema = {
  name: 'outlook_create_event',
  description: 'Create a new calendar event in Outlook with optional Teams meeting integration',
  inputSchema: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'Event subject/title',
      },
      start: {
        type: 'object',
        description: 'Event start date and time configuration',
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
        description: 'Event end date and time configuration',
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
      onlineMeetingProvider: {
        type: 'string',
        enum: ['teamsForBusiness', 'skypeForBusiness'],
        description: 'Online meeting provider (default: "teamsForBusiness")',
      },
      recurrence: {
        type: 'object',
        description: 'Recurrence pattern for recurring meetings',
        properties: {
          pattern: {
            type: 'object',
            description: 'The recurrence pattern',
            properties: {
              type: {
                type: 'string',
                enum: ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'],
                description: 'The recurrence pattern type',
              },
              interval: {
                type: 'integer',
                minimum: 1,
                description: 'Number of units between occurrences',
              },
              daysOfWeek: {
                type: 'array',
                items: {
                  type: 'string',
                  enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
                },
                description: 'Days of the week (required for weekly, relativeMonthly, relativeYearly)',
              },
              dayOfMonth: {
                type: 'integer',
                minimum: 1,
                maximum: 31,
                description: 'Day of the month (required for absoluteMonthly, absoluteYearly)',
              },
              month: {
                type: 'integer',
                minimum: 1,
                maximum: 12,
                description: 'Month of the year (required for absoluteYearly, relativeYearly)',
              },
              index: {
                type: 'string',
                enum: ['first', 'second', 'third', 'fourth', 'last'],
                description: 'Instance of the allowed days (for relativeMonthly, relativeYearly)',
              },
              firstDayOfWeek: {
                type: 'string',
                enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
                description: 'First day of the week (for weekly patterns, default: sunday)',
              },
            },
            required: ['type', 'interval'],
          },
          range: {
            type: 'object',
            description: 'The recurrence range',
            properties: {
              type: {
                type: 'string',
                enum: ['numbered', 'endDate', 'noEnd'],
                description: 'The recurrence range type',
              },
              startDate: {
                type: 'string',
                format: 'date',
                description: 'Start date of the recurrence (YYYY-MM-DD)',
              },
              endDate: {
                type: 'string',
                format: 'date',
                description: 'End date of the recurrence (YYYY-MM-DD, required for endDate type)',
              },
              numberOfOccurrences: {
                type: 'integer',
                minimum: 1,
                description: 'Number of occurrences (required for numbered type)',
              },
            },
            required: ['type', 'startDate'],
          },
        },
        required: ['pattern', 'range'],
      },
    },
    required: ['subject', 'start', 'end'],
  },
};

// Export all calendar schemas as an array for easy iteration
export const calendarSchemas = [
  listEventsSchema,
  createEventSchema,
];

// Export mapping for quick lookup
export const calendarSchemaMap = {
  'outlook_list_events': listEventsSchema,
  'outlook_create_event': createEventSchema,
};