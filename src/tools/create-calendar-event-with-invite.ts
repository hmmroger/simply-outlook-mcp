import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { GraphService } from "../simply-outlook/graph-service.js";
import { getErrorToolResult, textToolResult, toCalendarEventResult } from "./tool-utils.js";

export const CREATE_CALENDAR_EVENT_WITH_INVITE_TOOL_NAME = "create-calendar-event-with-invite";

const DEFAULT_EVENT_DURATION_MINUTES = 30;

export const registerCreateCalendarEventWithInviteTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    CREATE_CALENDAR_EVENT_WITH_INVITE_TOOL_NAME,
    "Create a calendar event in Outlook and send invitations to specified attendees. Use this tool when you need to invite other people to the event.",
    {
      subject: z.string().describe("The title/subject of the calendar event"),
      content: z
        .string()
        .optional()
        .describe("Optional description or body content for the event. Must be in markdown or plain text format."),
      startDateTime: z
        .string()
        .describe(
          "The event start date and time in ISO format using local time zone. Format: 'YYYY-MM-DDTHH:mm:ss' (e.g., '2025-12-25T14:30:00')"
        ),
      endDateTime: z
        .string()
        .optional()
        .describe(
          `The event end date and time in ISO format using local time zone. Optional - if not provided, the event will last ${DEFAULT_EVENT_DURATION_MINUTES} minutes. Format: 'YYYY-MM-DDTHH:mm:ss' (e.g., '2025-12-25T15:00:00')`
        ),
      location: z
        .string()
        .optional()
        .describe("Optional location or venue for the event (e.g., 'Conference Room A', 'Zoom Meeting', 'Central Park')"),
      userEmails: z
        .string()
        .array()
        .describe(
          "Array of email addresses to invite as attendees. Each attendee will receive a calendar invitation. Use empty array [] if no attendees should be invited."
        ),
      isMeeting: z
        .boolean()
        .optional()
        .describe(
          "Optional flag to mark this event as a meeting. When true, this enables meeting-specific features like online meeting links."
        ),
    },
    async ({ subject, content, startDateTime, endDateTime, location, userEmails, isMeeting }) => {
      try {
        const startDateTimeUtc = new Date(startDateTime).toISOString();

        let calculatedEndDateTime = endDateTime;
        if (!calculatedEndDateTime) {
          const startDate = new Date(startDateTime);
          startDate.setMinutes(startDate.getMinutes() + DEFAULT_EVENT_DURATION_MINUTES);
          calculatedEndDateTime = startDate.toISOString();
        }

        const endDateTimeUtc = new Date(calculatedEndDateTime).toISOString();
        const eventData = await graphService.createCalendarEvent(
          subject,
          content || "",
          startDateTimeUtc,
          endDateTimeUtc,
          userEmails,
          location,
          isMeeting
        );
        return textToolResult([`Event created successfully:`, JSON.stringify(toCalendarEventResult(eventData))]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to create calendar event.");
      }
    }
  );
};
