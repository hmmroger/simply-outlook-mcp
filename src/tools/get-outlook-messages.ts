import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getErrorToolResult, getUtcDateTimeToTheDay, textToolResult, toMailMessageResult, utcDateTimeToLocal } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const GET_OUTLOOK_MESSAGES_TOOL_NAME = "get-outlook-messages";

const DEFAULT_OUTLOOK_MESSAGES_LIMIT = 25;
const MAX_OUTLOOK_MESSAGES_LIMIT = 50;

export const registerGetOutlookMessagesTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    GET_OUTLOOK_MESSAGES_TOOL_NAME,
    "Retrieve a list of mail messages from Outlook received since a specified date and time.",
    {
      receivedDateTime: z
        .string()
        .optional()
        .describe(
          "Optional start date and time to filter messages from. If not provided, defaults to the current date and time. Format: 'YYYY-MM-DDTHH:mm:ss' in local time zone (e.g., '2025-12-25T09:00:00')"
        ),
      limit: z
        .number()
        .optional()
        .describe(
          `Maximum number of messages to return (default: ${DEFAULT_OUTLOOK_MESSAGES_LIMIT}, maximum allowed: ${MAX_OUTLOOK_MESSAGES_LIMIT}).`
        ),
      skip: z
        .number()
        .optional()
        .describe("Number of messages to skip for pagination purposes (default: 0). Useful for retrieving additional pages of results."),
    },
    async ({ receivedDateTime, limit, skip }) => {
      try {
        if (limit && limit > MAX_OUTLOOK_MESSAGES_LIMIT) {
          throw new Error(`limit is more than max number of messages allowed: ${MAX_OUTLOOK_MESSAGES_LIMIT}.`);
        }

        let receivedStartDateTimeUtc: string;
        if (receivedDateTime) {
          receivedStartDateTimeUtc = new Date(receivedDateTime).toISOString();
        } else {
          receivedStartDateTimeUtc = getUtcDateTimeToTheDay();
        }

        const messagesData = await graphService.getOutlookMessages(
          {
            startDateTime: receivedStartDateTimeUtc,
          },
          undefined,
          limit || DEFAULT_OUTLOOK_MESSAGES_LIMIT,
          skip
        );
        if (!messagesData.length) {
          return textToolResult(["No Outlook messages found."]);
        }

        return textToolResult([
          `Outlook messages since ${utcDateTimeToLocal(receivedStartDateTimeUtc)}:`,
          JSON.stringify(messagesData.map((message) => toMailMessageResult(message))),
        ]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to get Outlook messages.");
      }
    }
  );
};
