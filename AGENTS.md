# Task Compilation Requirements

This app should support all its functionality through the tcp connection and mcp server.
Make sure to sync the feature changes with the tcp and mcp servers.

## Code changes

For any code changes, if you can add unit tests without breaking the architecture add them.
Run code formatter, linter and typescript compilation, and make sure they don't fail ort create warnings/hints.


# Architectural Principles

- The main process owns workbook truth. The renderer never owns authoritative workbook state.
- All workbook writes should flow through the controller and, ideally, through transaction-shaped operations rather than ad hoc mutations.
- Persist raw user input; derive display/computed views separately. The formula work already follows this well.
- TCP is the canonical automation contract. MCP should remain a typed/documented adapter over - TCP, not a second business-logic implementation.
Transport layers stay thin. No workbook rules in TCP handlers, MCP tools, preload, or React components.
- Shared contracts live in workbook-core.ts so UI, TCP, and MCP speak the same language.
- Workbook changes must preserve summary/version/event behavior so UI and external clients stay in sync.