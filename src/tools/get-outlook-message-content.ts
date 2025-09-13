import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getErrorToolResult, getUtcDateTimeToTheDay, textToolResult, toMailMessageResult, utcDateTimeToLocal } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const GET_OUTLOOK_MESSAGE_CONTENT_TOOL_NAME = "get-outlook-message-content";

export const registerGetOutlookMessageContentTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    GET_OUTLOOK_MESSAGE_CONTENT_TOOL_NAME,
    "Retrieve the full content of a specific Outlook mail message by its ID.",
    {
      id: z.string().describe("The unique identifier of the mail message to retrieve the content for."),
    },
    async ({ id }) => {
      try {
        const messageData = await graphService.getOutlookMessageById(id);
        const messageResult = toMailMessageResult(messageData);

        return textToolResult(
          [
            `Subject: ${messageResult.subject}`,
            `From: ${messageResult.from.name} <${messageResult.from.email}>`,
            `Received: ${messageResult.receivedDateTime}`,
            `Importance: ${messageResult.importance}`,
            `Read: ${messageResult.isRead ? "Yes" : "No"}`,
            `Draft: ${messageResult.isDraft ? "Yes" : "No"}`,
            messageResult.toRecipients && messageResult.toRecipients.length > 0
              ? `To: ${messageResult.toRecipients.map((r) => `${r.name} <${r.email}>`).join(", ")}`
              : "",
            "",
            "Content:",
            messageResult.content || "No content available",
          ].filter((line) => line !== "")
        );
      } catch (error) {
        return getErrorToolResult(error, "Failed to get Outlook message content.");
      }
    }
  );
};
