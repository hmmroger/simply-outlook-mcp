import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { GraphService } from "../simply-outlook/graph-service.js";
import { getErrorToolResult, textToolResult, toCalendarEventResult } from "./tool-utils.js";

export const CREATE_CALENDAR_EVENT_TOOL_NAME = "create-calendar-event";

const DEFAULT_EVENT_DURATION_MINUTES = 30;

export const registerCreateCalendarEventTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    CREATE_CALENDAR_EVENT_TOOL_NAME,
    "Create a personal calendar event in Outlook without sending invitations to other attendees. This creates a private event only on the user's calendar.",
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
        .describe("Optional location or venue for the event (e.g., 'Conference Room A', 'Airport', 'Central Park')"),
    },
    async ({ subject, content, startDateTime, endDateTime, location }) => {
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
          undefined,
          location
        );
        return textToolResult([`Event created successfully:`, JSON.stringify(toCalendarEventResult(eventData))]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to create calendar event.");
      }
    }
  );
};
