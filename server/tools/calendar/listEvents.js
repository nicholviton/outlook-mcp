import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// List calendar events
export async function listEventsTool(authManager, args) {
  const { startDateTime, endDateTime, limit = 10, calendar } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendar ? `/me/calendars/${calendar}/events` : '/me/events';
    const options = {
      select: 'subject,start,end,location,attendees,bodyPreview',
      top: limit,
      orderby: 'start/dateTime',
    };

    if (startDateTime && endDateTime) {
      options.filter = `start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`;
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
      preview: event.bodyPreview,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ events, count: events.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list events');
  }
}