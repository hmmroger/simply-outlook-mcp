import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getUtcDateTimeToTheHour, getErrorToolResult, textToolResult, toCalendarEventResult } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const GET_CALENDAR_EVENTS_TOOL_NAME = "get-calendar-events";

const DEFAULT_CALENDAR_EVENTS_LIMIT = 25;
const MAX_CALENDAR_EVENTS_LIMIT = 50;
const DEFAULT_CALENDAR_RANGE_DAYS = 7;

export const registerGetCalendarEventsTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    GET_CALENDAR_EVENTS_TOOL_NAME,
    "Retrieve a list of calendar events from Outlook within a specified date range. Returns both personal events and meetings with attendees.",
    {
      startDateTime: z
        .string()
        .optional()
        .describe(
          "Optional start date and time to filter events from. If not provided, defaults to the current date and time. Format: 'YYYY-MM-DDTHH:mm:ss' in local time zone (e.g., '2025-12-25T09:00:00')"
        ),
      endDateTime: z
        .string()
        .optional()
        .describe(
          `Optional end date and time to filter events until. If not provided, defaults to ${DEFAULT_CALENDAR_RANGE_DAYS} days from the start date. Format: 'YYYY-MM-DDTHH:mm:ss' in local time zone (e.g., '2025-12-25T17:00:00')`
        ),
      limit: z
        .number()
        .optional()
        .describe(
          `Maximum number of events to return (default: ${DEFAULT_CALENDAR_EVENTS_LIMIT}, maximum allowed: ${MAX_CALENDAR_EVENTS_LIMIT}).`
        ),
      skip: z
        .number()
        .optional()
        .describe("Number of events to skip for pagination purposes (default: 0). Useful for retrieving additional pages of results."),
    },
    async ({ startDateTime, endDateTime, limit, skip }) => {
      try {
        if (limit && limit > MAX_CALENDAR_EVENTS_LIMIT) {
          throw new Error(`limit is more than max number of events allowed: ${MAX_CALENDAR_EVENTS_LIMIT}.`);
        }

        let searchStartDateTimeUtc: string;
        if (startDateTime) {
          searchStartDateTimeUtc = new Date(startDateTime).toISOString();
        } else {
          searchStartDateTimeUtc = getUtcDateTimeToTheHour();
        }

        let searchEndDateTimeUtc: string;
        if (endDateTime) {
          searchEndDateTimeUtc = new Date(endDateTime).toISOString();
        } else {
          // Calculate end date as DEFAULT_CALENDAR_RANGE_DAYS from start date
          const startDate = new Date(searchStartDateTimeUtc);
          const endDate = new Date(startDate.getTime() + DEFAULT_CALENDAR_RANGE_DAYS * 24 * 60 * 60 * 1000);
          searchEndDateTimeUtc = endDate.toISOString();
        }

        const eventsData = await graphService.getCalendarEvents(
          {
            startDateTime: searchStartDateTimeUtc,
            endDateTime: searchEndDateTimeUtc,
          },
          limit || DEFAULT_CALENDAR_EVENTS_LIMIT,
          skip
        );
        if (!eventsData.length) {
          return textToolResult(["No calendar events found."]);
        }

        return textToolResult([
          `Do not show the event ID to the user.`,
          `There are ${eventsData.length} calendar events found:`,
          JSON.stringify(eventsData.map((event) => toCalendarEventResult(event))),
        ]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to get calendar events.");
      }
    }
  );
};
