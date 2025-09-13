import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getErrorToolResult, textToolResult } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const REPLY_OUTLOOK_MESSAGE_TOOL_NAME = "reply-outlook-message";

export const registerReplyOutlookMessageTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    REPLY_OUTLOOK_MESSAGE_TOOL_NAME,
    "Reply to an existing Outlook mail message with new content.",
    {
      messageId: z.string().describe("The unique identifier of the mail message to reply to."),
      content: z
        .string()
        .describe("The reply content/body of the email message. Supports Markdown formatting which will be converted to HTML."),
    },
    async ({ messageId, content }) => {
      try {
        if (!messageId) {
          throw new Error("Message ID is required to reply to a message.");
        }

        if (!content || content.trim().length === 0) {
          throw new Error("Reply content cannot be empty.");
        }

        await graphService.replyOutlookMessage(messageId, content);

        return textToolResult([`Successfully sent reply to Outlook message with ID: ${messageId}`]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to reply to Outlook message.");
      }
    }
  );
};
