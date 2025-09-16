import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getErrorToolResult, textToolResult, toMailMessageResult } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const SEARCH_OUTLOOK_MESSAGES_TOOL_NAME = "search-outlook-messages";

const DEFAULT_OUTLOOK_MESSAGES_LIMIT = 25;
const MAX_OUTLOOK_MESSAGES_LIMIT = 50;

export const registerSearchOutlookMessagesTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    SEARCH_OUTLOOK_MESSAGES_TOOL_NAME,
    "Search for mail messages in Outlook based on keywords.",
    {
      keywords: z.string().describe("Search keywords to find messages. Can include sender names, subject text, or message content."),
      limit: z
        .number()
        .optional()
        .describe(
          `Maximum number of messages to return (default: ${DEFAULT_OUTLOOK_MESSAGES_LIMIT}, maximum allowed: ${MAX_OUTLOOK_MESSAGES_LIMIT}).`
        ),
    },
    async ({ keywords, limit }) => {
      try {
        if (!keywords) {
          throw new Error("Provide 'keywords' to search messages.");
        }

        if (limit && limit > MAX_OUTLOOK_MESSAGES_LIMIT) {
          throw new Error(`limit is more than max number of messages allowed: ${MAX_OUTLOOK_MESSAGES_LIMIT}.`);
        }

        const messagesData = await graphService.getOutlookMessages(undefined, keywords, limit || DEFAULT_OUTLOOK_MESSAGES_LIMIT);
        if (!messagesData.length) {
          return textToolResult(["No Outlook messages found."]);
        }

        return textToolResult([
          `Do not show the message ID to the user.`,
          `There are ${messagesData.length} Outlook messages matching keywords [${keywords}]:`,
          JSON.stringify(messagesData.map((message) => toMailMessageResult(message))),
        ]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to search Outlook messages.");
      }
    }
  );
};
