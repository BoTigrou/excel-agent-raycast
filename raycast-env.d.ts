/// <reference types="@raycast/api">

/* 🚧 🚧 🚧
 * This file is auto-generated from the extension's manifest.
 * Do not modify manually. Instead, update the `package.json` file.
 * 🚧 🚧 🚧 */

/* eslint-disable @typescript-eslint/ban-types */

type ExtensionPreferences = {
  /** AI Provider - Choose which AI provider to use for generating scripts */
  "aiProvider": "raycast" | "openai" | "gemini" | "claude",
  /** OpenAI API Key - Your OpenAI API key (required if using OpenAI) */
  "openaiApiKey"?: string,
  /** Google Gemini API Key - Your Google Gemini API key (required if using Gemini) */
  "geminiApiKey"?: string,
  /** Anthropic Claude API Key - Your Anthropic API key (required if using Claude) */
  "claudeApiKey"?: string
}

/** Preferences accessible in all the extension's commands */
declare type Preferences = ExtensionPreferences

declare namespace Preferences {
  /** Preferences accessible in the `ask-excel` command */
  export type AskExcel = ExtensionPreferences & {}
  /** Preferences accessible in the `read-excel` command */
  export type ReadExcel = ExtensionPreferences & {}
}

declare namespace Arguments {
  /** Arguments passed to the `ask-excel` command */
  export type AskExcel = {}
  /** Arguments passed to the `read-excel` command */
  export type ReadExcel = {}
}

