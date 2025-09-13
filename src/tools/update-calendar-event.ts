import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { GraphService } from "../simply-outlook/graph-service.js";
import { getErrorToolResult, textToolResult, toCalendarEventResult } from "./tool-utils.js";

export const UPDATE_CALENDAR_EVENT_TOOL_NAME = "update-calendar-event";

export const registerUpdateCalendarEventTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    UPDATE_CALENDAR_EVENT_TOOL_NAME,
    "Update an existing calendar event in Outlook. You can modify the subject, content, start/end times, or location. At least one field must be provided to update.",
    {
      id: z.string().describe("The unique identifier of the calendar event to update"),
      subject: z.string().optional().describe("Optional new title/subject for the calendar event"),
      content: z
        .string()
        .optional()
        .describe("Optional new description or body content for the event. Must be in markdown or plain text format."),
      startDateTime: z
        .string()
        .optional()
        .describe(
          "Optional new start date and time in ISO format using local time zone. Format: 'YYYY-MM-DDTHH:mm:ss' (e.g., '2025-12-25T14:30:00')"
        ),
      endDateTime: z
        .string()
        .optional()
        .describe(
          "Optional new end date and time in ISO format using local time zone. Format: 'YYYY-MM-DDTHH:mm:ss' (e.g., '2025-12-25T15:00:00')"
        ),
      location: z
        .string()
        .optional()
        .describe(
          "Optional new location or venue for the event (e.g., 'Conference Room A', 'Airport', 'Central Park'). Use empty string to remove location."
        ),
    },
    async ({ id, subject, content, startDateTime, endDateTime, location }) => {
      try {
        const startDateTimeUtc = startDateTime ? new Date(startDateTime).toISOString() : undefined;
        const endDateTimeUtc = endDateTime ? new Date(endDateTime).toISOString() : undefined;

        const eventData = await graphService.updateCalendarEvent(id, content, subject, startDateTimeUtc, endDateTimeUtc, location);

        return textToolResult([`Calendar event updated successfully:`, JSON.stringify(toCalendarEventResult(eventData))]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to update calendar event.");
      }
    }
  );
};
