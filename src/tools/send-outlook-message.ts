import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getErrorToolResult, textToolResult } from "./tool-utils.js";
import { GraphService } from "../simply-outlook/graph-service.js";

export const SEND_OUTLOOK_MESSAGE_TOOL_NAME = "send-outlook-message";

export const registerSendOutlookMessageTool = async (server: McpServer, graphService: GraphService) => {
  server.tool(
    SEND_OUTLOOK_MESSAGE_TOOL_NAME,
    "Send a new mail message through Outlook to specified recipients.",
    {
      subject: z.string().describe("The subject line of the email message."),
      content: z.string().describe("The content/body of the email message. Must be in markdown or plain text format."),
      recipientEmails: z.string().array().describe("Array of email addresses to send the message to."),
    },
    async ({ subject, content, recipientEmails }) => {
      try {
        if (!recipientEmails || recipientEmails.length === 0) {
          throw new Error("At least one recipient email address is required.");
        }

        await graphService.sendOutlookMessage(subject, content, recipientEmails);

        return textToolResult([`Successfully sent Outlook message.`]);
      } catch (error) {
        return getErrorToolResult(error, "Failed to send Outlook message.");
      }
    }
  );
};
