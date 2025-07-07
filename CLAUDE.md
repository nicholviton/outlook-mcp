# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This project implements a Microsoft Outlook MCP (Model Context Protocol) server that enables AI assistants to interact with Outlook email, calendar, contacts, and tasks through Microsoft Graph API. The implementation follows DXT extension format for Claude Desktop deployment.

## Project Structure

This is a documentation-first project in early planning phase:

- `docs/` - Complete project documentation and planning
  - `outlook-mcp-plan.md` - Detailed development plan with 6 phases
  - `PRD.md` - Comprehensive Product Requirements Document defining all 45+ MCP tools
  - `graph_api.md` - Microsoft Graph API integration reference guide
  - `DXT_format.md` - Claude Desktop Extension format specification

The actual implementation has not yet been started - this repository currently contains only planning documentation.

## Development Workflow

Since this is a planning-phase project, no build commands exist yet. When implementation begins, the plan calls for:

1. **Project Initialization**: `dxt init` to create DXT extension structure
2. **Dependencies**: `npm install @modelcontextprotocol/sdk @azure/identity @microsoft/microsoft-graph-client`
3. **Testing**: `npx @modelcontextprotocol/inspector` for MCP server testing
4. **Packaging**: `dxt pack . outlook-mcp.dxt` to create extension

## Key Implementation Points

### Authentication Architecture
- OAuth 2.0 authorization code flow with PKCE
- Azure AD app registration required with specific permissions
- Token management with 60-minute access tokens and 90-day refresh tokens
- Secure credential storage using OS keychain

### API Integration Requirements  
- Microsoft Graph API v1.0 endpoint
- Rate limiting: 4 concurrent requests per mailbox maximum
- Exponential backoff for 429 throttling responses
- Support for batch requests (up to 20 operations)

### MCP Server Structure
Based on the PRD, the server will implement 45+ tools across categories:
- Email operations (16 tools)
- Calendar operations (11 tools) 
- Contact operations (8 tools)
- Task operations (6 tools)
- Utility tools (4 tools)

### Security Considerations
- All sensitive data (API keys, tokens) stored in OS credential store
- Input validation and sanitization for all user inputs
- RBAC compliance with Azure AD permissions
- Audit logging for enterprise compliance

## Next Steps for Implementation

1. **Phase 1**: Azure AD app registration and authentication setup
2. **Phase 2**: Basic MCP server with core email tools
3. **Phase 3**: Calendar and contact integration
4. **Phase 4**: Advanced features and enterprise compliance
5. **Phase 5**: Testing and optimization
6. **Phase 6**: Packaging and deployment as DXT extension

Refer to `docs/outlook-mcp-plan.md` for detailed implementation timeline and `docs/PRD.md` for complete tool specifications.