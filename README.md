# SharePoint MCP Server

A Model Context Protocol (MCP) server exposing SharePoint resources and tools for use by LLMs or MCP clients.

## Setup

1. Copy `.env.example` to `.env` and fill in your Azure/SharePoint credentials.
2. Install dependencies:

    ```sh
    pnpm install
    ```

3. Run in development mode:

    ```sh
    pnpm dev
    ```

## Reference

- [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- [Microsoft Graph JS](https://www.npmjs.com/package/microsoft-graph)