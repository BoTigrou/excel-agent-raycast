# Excel Agent for Raycast

Control Microsoft Excel with natural language. Ask questions, format cells, create tables, and more — all from Raycast.

![Excel Agent Demo](https://raw.githubusercontent.com/julien-archambault/excel-agent-raycast/main/assets/demo.png)

## Features

- **Natural Language Control** — Just type what you want to do in Excel
- **Quick Actions** — One-click formatting for common tasks
- **AI-Powered** — Uses AI to generate AppleScript for complex operations
- **Multiple AI Providers** — Works with Raycast AI, OpenAI, Google Gemini, or Anthropic Claude
- **Smart Caching** — Built-in scripts for instant common operations

## Requirements

- **macOS** — Uses AppleScript to control Excel
- **Microsoft Excel** — Must be installed and running
- **Raycast v1.104.3+**

## Setup

### 1. Install the Extension

Install from the Raycast Store or build from source.

### 2. Grant Accessibility Permissions

The first time you run a command, macOS will ask you to grant Raycast Accessibility permissions:

1. Open **System Settings** → **Privacy & Security** → **Accessibility**
2. Add and enable **Raycast**

### 3. Configure AI Provider

Open the extension preferences and choose your AI provider:

| Provider | Setup Required |
|----------|---------------|
| **Raycast AI** | Raycast Pro subscription |
| **OpenAI** | API key from [platform.openai.com](https://platform.openai.com) |
| **Google Gemini** | API key from [Google AI Studio](https://makersuite.google.com) |
| **Anthropic Claude** | API key from [console.anthropic.com](https://console.anthropic.com) |

## Usage

### "Ask Excel" Command

Open Raycast and type **"Ask Excel"** to see:

- **Quick Actions** — Instant formatting (Bold Headers, Currency, Borders, etc.)
- **Recent Commands** — Your history of custom commands
- **Custom Command** — Type anything you want Excel to do

#### Example Commands

- *"Create a table with headers: Date, Product, Revenue"*
- *"Make row 1 bold with blue background"*
- *"Format column B as currency"*
- *"Add sum formula in C10"*
- *"Apply financial style (blue inputs, black formulas)"*
- *"Auto-fit all columns"*
- *"Freeze the top row"*

### "Read Excel" Command

Quickly view the current selection and sheet information without leaving Raycast.

## Tips

- **Be specific** — Instead of "make it bold", try "make row 1 bold"
- **Combine actions** — "Bold headers and add borders to used range"
- **Use Quick Actions** — They're instant and don't require AI
- **Check Excel is running** — The extension will remind you if Excel isn't open

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Excel not running" | Open Microsoft Excel with a workbook |
| "Accessibility permissions" | Add Raycast to System Settings → Privacy & Security → Accessibility |
| "No workbook open" | Open or create an Excel file |
| AI not responding | Check your API key in preferences |

## Privacy

- Your Excel data stays on your Mac
- AI providers only receive the minimal context needed (sheet name, selection address)
- No data is stored on external servers by this extension

## License

MIT License — See [LICENSE](LICENSE) for details.

---

Made with ❤️ for Excel power users.
