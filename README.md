# D365 F&O Drill Engine

A Model Context Protocol (MCP) server that turns Dynamics 365 Finance & Operations 
certification prep into a conversation. Runs inside Claude Desktop. Drill by module 
(Inventory, AR, GL, WMS, Production Control), track your accuracy, review the ones 
you get wrong.

## What it does

Ten MCP tools across four categories:
- **Drill / Practice** — start a drill session, answer, get scored, review
- **Lookup** — pull specific concepts on demand
- **Reference** — quick reference cards for exam-relevant topics
- **Discovery** — see what categories and how many cards are available

## Who it's for

Anyone preparing for MB-330, MB-500, MB-700, MB-800, or MB-820. Also useful as a
day-to-day reference for practitioners who want to sharpen their memory on specific
F&O topics without leaving Claude Desktop.

## Installation

**Prerequisites:**
- Python 3.10+
- Claude Desktop
- (Optional) A Dataverse tenant if you want to store your progress remotely.
  Otherwise, local storage is used by default.

**Setup:**

1. Clone the repo:
   \`\`\`bash
   git clone https://github.com/AntonTjiptadi/d365fo-drill-engine.git
   cd d365fo-drill-engine
   \`\`\`

2. Create a virtual environment and install dependencies:
   \`\`\`bash
   python -m venv .venv
   source .venv/bin/activate  # or .venv\Scripts\activate on Windows
   pip install -r requirements.txt
   \`\`\`

3. Copy `.env.example` to `.env` and fill in your values (if using Dataverse).

4. Add the MCP server to your Claude Desktop config
   (`~/Library/Application Support/Claude/claude_desktop_config.json` on Mac,
   `%APPDATA%\Claude\claude_desktop_config.json` on Windows):

   \`\`\`json
   {
     "mcpServers": {
       "d365fo-drill": {
         "command": "python",
         "args": ["/absolute/path/to/mcp_server/server.py"]
       }
     }
   }
   \`\`\`

5. Restart Claude Desktop. In a new conversation, you should see the drill tools 
   available.

## Usage

In Claude Desktop:

> "Start a 10-question drill on Inventory Management."

> "Show me what I got wrong in my last session."

> "Give me a quick reference card on cost groups."

## Content

Drill cards are authored from practitioner experience across 15+ years of D365 F&O
delivery and enriched with public Microsoft Learn documentation. Cards are stored 
locally in `content/` by default. Contribution guidelines below.

## Contributing

Cards live as structured JSON in `content/<category>/`. To add or improve a card:
- Fork the repo
- Add or edit the card following the schema in `content/SCHEMA.md`
- Open a PR with a brief description of what the card covers and why it matters

## Limitations

This tool is a study aid, not a replacement for hands-on practice or official 
Microsoft Learn paths. It reflects one practitioner's synthesis of the exam surface 
and may be incomplete or biased toward areas I've worked in most.

## About

Built by [Anton Tjiptadi](https://linkedin.com/in/anton-tjiptadi-18b3aa21/), Solution Architect
working on D365 F&O and agentic AI. If you find this useful, connect on LinkedIn or 
open a discussion.

## License

MIT. See LICENSE.