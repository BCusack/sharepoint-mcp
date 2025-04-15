import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import dotenv from "dotenv";

import { register } from "microsoft-graph/services/context";
import { getEnvironmentVariable } from "microsoft-graph/services/environmentVariable";
import { TenantId } from "microsoft-graph/models/TenantId";
import { ClientId } from "microsoft-graph/models/ClientId";
import { SiteId } from "microsoft-graph/models/SiteId";
import { ClientSecret } from "microsoft-graph/models/ClientSecret";

import { createDriveRef } from "microsoft-graph/services/drive";
import { createSiteRef } from "microsoft-graph/services/site";
import { DriveId } from "microsoft-graph/models/DriveId";
import listDriveItems from "microsoft-graph/operations/driveItem/listDriveItems";
import { createDriveItemRef } from "microsoft-graph/services/driveItem";
import { DriveItemId } from "microsoft-graph/models/DriveItemId";
import getDriveItem from "microsoft-graph/operations/driveItem/getDriveItem";
dotenv.config();

// Environment/config
const tenantId = getEnvironmentVariable("TENANT_ID") as TenantId;
const clientId = getEnvironmentVariable("CLIENT_ID") as ClientId;
const clientSecret = getEnvironmentVariable("CLIENT_SECRET") as ClientSecret;

const driveId = getEnvironmentVariable("DRIVE_ID") as DriveId;
const siteId = getEnvironmentVariable("SITE_ID") as SiteId;

const contextRef = register(tenantId, clientId, clientSecret);
const siteRef = createSiteRef(contextRef, siteId);
const driveRef = createDriveRef(siteRef, driveId);

// Create MCP server
const server = new McpServer({
  name: "SharePoint MCP",
  version: "1.0.0"
});

// Resource: List Drive Items (files/folders)
server.resource(
  "sharepoint.drive.items",
  new ResourceTemplate("sharepoint://drive/{folderId?}", { list: undefined }),
  async (uri, { folderId }) => {
    const ref = folderId ? createDriveRef(siteRef, folderId as DriveId) : driveRef;
    const items = await listDriveItems(ref);
    return {
      contents: items.map(item => ({
        uri: `sharepoint://drive/${item.id}`,
        text: item.name,
        extra: item
      }))
    };
  }
);

// Resource: List Sites
server.resource(
  "sharepoint.sites",
  new ResourceTemplate("sharepoint://sites", { list: undefined }),
  async (uri) => {
    const sites = await listSites(contextRef);
    return {
      contents: sites.map((site: { id: string; name: string }) => ({
        uri: `sharepoint://sites/${site.id}`,
        text: site.name,
        extra: site
      }))
    };
  }
);

// Resource: List Lists
server.resource(
  "sharepoint.lists",
  new ResourceTemplate("sharepoint://lists", { list: undefined }),
  async (uri) => {
    const lists = await listLists(siteRef);
    return {
      contents: lists.map(list => ({
        uri: `sharepoint://lists/${list.id}`,
        text: list.name,
        extra: list
      }))
    };
  }
);

// Tool: Search Files
server.tool(
  "sharepoint.searchFiles",
  { query: z.string() },
  async ({ query }) => {
    const results = await searchDriveItems(driveRef, query);
    return {
      content: results.map(item => ({
        type: "text",
        text: item.name,
        extra: item
      }))
    };
  }
);

// Tool: Create Folder
server.tool(
  "sharepoint.createFolder",
  { folderName: z.string() },
  async ({ folderName }) => {
    const folder = await createFolder(driveRef, folderName);
    return {
      content: [{
        type: "text",
        text: `Created folder: ${folder.name}`,
        extra: folder
      }]
    };
  }
);

// Start MCP server using stdio transport (suitable for LLM plugins, etc)
const transport = new StdioServerTransport();
await server.connect(transport);