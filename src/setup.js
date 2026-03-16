#!/usr/bin/env node

/**
 * Post-install script: registers sharepoint-online-mcp in Claude Desktop config.
 * 
 * Runs automatically after `npm install -g` or `npx` first-run.
 * - Finds Claude Desktop config (macOS, Windows, Linux)
 * - Adds the "sharepoint" MCP server entry if not already present
 * - Never overwrites an existing "sharepoint" entry
 * - Fails silently — install must not break if Claude Desktop isn't installed
 */

import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir, platform } from "node:os";

function getConfigPath() {
  const home = homedir();
  switch (platform()) {
    case "darwin":
      return join(home, "Library", "Application Support", "Claude", "claude_desktop_config.json");
    case "win32":
      return join(process.env.APPDATA || join(home, "AppData", "Roaming"), "Claude", "claude_desktop_config.json");
    case "linux":
      return join(home, ".config", "Claude", "claude_desktop_config.json");
    default:
      return null;
  }
}

function install() {
  const configPath = getConfigPath();
  if (!configPath) return;

  let config = {};
  if (existsSync(configPath)) {
    try {
      config = JSON.parse(readFileSync(configPath, "utf-8"));
    } catch {
      // Corrupt config — don't touch it
      return;
    }
  } else {
    // Create the directory if needed
    try {
      mkdirSync(join(configPath, ".."), { recursive: true });
    } catch {
      return;
    }
  }

  if (!config.mcpServers) {
    config.mcpServers = {};
  }

  // Don't overwrite if user already has a sharepoint entry
  if (config.mcpServers.sharepoint) {
    process.stderr.write("✅ SharePoint MCP already registered in Claude Desktop.\n");
    return;
  }

  config.mcpServers.sharepoint = {
    command: "npx",
    args: ["-y", "sharepoint-online-mcp"],
  };

  try {
    writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n", "utf-8");
    process.stderr.write("\n");
    process.stderr.write("✅ SharePoint MCP registered in Claude Desktop.\n");
    process.stderr.write("   Restart Claude Desktop to activate.\n");
    process.stderr.write("\n");
  } catch {
    // Permission denied, read-only fs, etc. — fail silently
  }
}

function uninstall() {
  const configPath = getConfigPath();
  if (!configPath || !existsSync(configPath)) return;

  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf-8"));
  } catch {
    return;
  }

  if (!config.mcpServers?.sharepoint) return;

  delete config.mcpServers.sharepoint;

  try {
    writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n", "utf-8");
    process.stderr.write("🗑️  SharePoint MCP removed from Claude Desktop config.\n");
  } catch {
    // fail silently
  }
}

// Called as: node setup.js [install|uninstall]
const action = process.argv[2] || "install";
if (action === "uninstall") {
  uninstall();
} else {
  install();
}
