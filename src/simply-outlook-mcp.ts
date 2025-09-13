import { TokenCredential } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ConsoleLogger } from "./common/console-logger.js";
import { VERSION } from "./version.js";
import { GraphService } from "./simply-outlook/graph-service.js";
import { SimplyOutlookMcpEnvs } from "./simply-outlook-mcp.types.js";
import { registerGetCalendarEventsTool, GET_CALENDAR_EVENTS_TOOL_NAME } from "./tools/get-calendar-events.js";
import { registerCreateCalendarEventTool, CREATE_CALENDAR_EVENT_TOOL_NAME } from "./tools/create-calendar-event.js";
import { registerUpdateCalendarEventTool, UPDATE_CALENDAR_EVENT_TOOL_NAME } from "./tools/update-calendar-event.js";
import {
  registerCreateCalendarEventWithInviteTool,
  CREATE_CALENDAR_EVENT_WITH_INVITE_TOOL_NAME,
} from "./tools/create-calendar-event-with-invite.js";
import { registerGetOutlookMessagesTool, GET_OUTLOOK_MESSAGES_TOOL_NAME } from "./tools/get-outlook-messages.js";
import { registerSearchOutlookMessagesTool, SEARCH_OUTLOOK_MESSAGES_TOOL_NAME } from "./tools/search-outlook-messages.js";
import { registerGetOutlookMessageContentTool, GET_OUTLOOK_MESSAGE_CONTENT_TOOL_NAME } from "./tools/get-outlook-message-content.js";
import { registerSendOutlookMessageTool, SEND_OUTLOOK_MESSAGE_TOOL_NAME } from "./tools/send-outlook-message.js";
import { registerReplyOutlookMessageTool, REPLY_OUTLOOK_MESSAGE_TOOL_NAME } from "./tools/reply-outlook-message.js";

export const SIMPLY_OUTLOOK_MCP_SCOPES = ["Calendars.ReadWrite", "Mail.Read", "Mail.Send", "User.Read"];

type ToolRegistration = (mcpServer: McpServer, graphService: GraphService) => Promise<void>;

const TOOL_DEFS: { name: string; tool: ToolRegistration }[] = [
  { name: GET_CALENDAR_EVENTS_TOOL_NAME, tool: registerGetCalendarEventsTool },
  { name: CREATE_CALENDAR_EVENT_TOOL_NAME, tool: registerCreateCalendarEventTool },
  { name: UPDATE_CALENDAR_EVENT_TOOL_NAME, tool: registerUpdateCalendarEventTool },
  { name: CREATE_CALENDAR_EVENT_WITH_INVITE_TOOL_NAME, tool: registerCreateCalendarEventWithInviteTool },
  { name: GET_OUTLOOK_MESSAGES_TOOL_NAME, tool: registerGetOutlookMessagesTool },
  { name: SEARCH_OUTLOOK_MESSAGES_TOOL_NAME, tool: registerSearchOutlookMessagesTool },
  { name: GET_OUTLOOK_MESSAGE_CONTENT_TOOL_NAME, tool: registerGetOutlookMessageContentTool },
  { name: SEND_OUTLOOK_MESSAGE_TOOL_NAME, tool: registerSendOutlookMessageTool },
  { name: REPLY_OUTLOOK_MESSAGE_TOOL_NAME, tool: registerReplyOutlookMessageTool },
];

export const createMcpServer = async (credential: TokenCredential): Promise<McpServer> => {
  const disabledToolsStr = process.env[SimplyOutlookMcpEnvs.SIMPLY_OUTLOOK_MCP_DISABLED_TOOLS];
  const disabledTools = new Set<string>(
    disabledToolsStr
      ? disabledToolsStr
          .split(",")
          .map((tool) => tool.trim().toLowerCase())
          .filter((tool) => !!tool)
      : []
  );

  const graphService = new GraphService(new ConsoleLogger("GraphService", true), credential, SIMPLY_OUTLOOK_MCP_SCOPES);
  if (!(await graphService.isAuthenticated())) {
    throw new Error("Please run 'npx simply-outlook-mcp --auth --client_id <CLIENT ID>' before using for the first time.");
  }

  const mcpServer = new McpServer({
    name: "simply-outlook-mcp",
    version: VERSION,
  });

  for (const toolDef of TOOL_DEFS) {
    if (!disabledTools.has(toolDef.name)) {
      await toolDef.tool(mcpServer, graphService);
    }
  }

  return mcpServer;
};
