#!/usr/bin/env node

import { config } from "dotenv";
import yargs from "yargs/yargs";
import { hideBin } from "yargs/helpers";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createMcpServer, SIMPLY_OUTLOOK_MCP_SCOPES } from "./simply-outlook-mcp.js";
import { SimplyOutlookMcpEnvs } from "./simply-outlook-mcp.types.js";
import { authenticate, getDeviceCredential } from "./auth-utils.js";

config({ quiet: true });

const main = async () => {
  const argv = yargs(hideBin(process.argv))
    .option("auth", {
      type: "boolean",
      description: "Initialize and save authentication record.",
    })
    .option("force", {
      type: "boolean",
      description: "Re-initialize authentication to switch account.",
      implies: ["auth"],
    })
    .option("client_id", {
      type: "string",
      description: "Application ID created in Azure portal.",
    })
    .option("tenant_id", {
      type: "string",
      description: "Application's tenant ID.",
    })
    .parseSync();

  const clientId = argv.client_id || process.env[SimplyOutlookMcpEnvs.SIMPLY_OUTLOOK_MCP_CLIENT_ID];
  if (!clientId) {
    throw new Error("Missing client ID.");
  }

  const tenantId = argv.tenant_id || process.env[SimplyOutlookMcpEnvs.SIMPLY_OUTLOOK_MCP_TENANT_ID];
  if (argv.auth) {
    await authenticate(clientId, SIMPLY_OUTLOOK_MCP_SCOPES, tenantId, argv.force);
    return;
  }

  const credential = getDeviceCredential(clientId, tenantId, argv.force);
  const server = await createMcpServer(credential);
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Simply Outlook MCP Server running on stdio.");
};

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
