# gantt-generator
Gantt Generator is an AI-powered project management engine built directly into Google Sheets. It can turn raw natural language notes into live, cascading, business-day-aware Gantt charts in seconds using the Groq API (Llama 3).

# 📊 Gantt Generator

Gantt Generator is a lightweight, AI-driven Project Management engine that lives entirely inside Google Sheets. It leverages **Groq API (Llama-3.3-70b-versatile)** to parse raw, natural language project notes and instantly generate a live, cascading Gantt chart. 

Stop manually dragging cells, fighting with Smartsheet, or writing complex date formulas. Drop your raw notes into the sidebar, and let the AI build your timeline.

## Core Features

* **Natural Language Parsing:** Paste meeting notes, Slack messages, or rough timelines. The AI extracts task names, durations, and parallel/sequential dependencies automatically.
* **Business-Day Engine:** Built-in logic automatically skips weekends (Saturdays and Sundays). A "1 Week" task perfectly spans 5 working days. Fractional durations (e.g., `0.29` weeks) are fully supported.
* **Live Cascading Updates:** The generated spreadsheet is fully interactive. If you manually change a task's duration in the sheet, the script instantly recalculates the End Date and redraws the timeline bars on the fly.
* **Groq:** Powered by Groq, generating complex schedules takes mere seconds.

---

## Quickstart Installation

You do not need to install any external libraries. This runs natively on Google Apps Script.

### Step 1: Get a Groq API Key
1. Go to [GroqCloud](https://console.groq.com/).
2. Create a free account and generate an API key (`gsk_...`).

### Step 2: Set up Google Sheets
1. Create a new, blank Google Sheet.
2. Click on **Extensions > Apps Script** in the top menu.
3. Delete any code in the editor and create two files: `Code.gs` and `Sidebar.html`.

### Step 3: Add the Code
1. Copy the contents of `Code.gs` from this repository and paste it into your `Code.gs` file.
2. **CRITICAL:** Paste your Groq API key into line 17 of `Code.gs`:
   ```javascript
   const GROQ_API_KEY = "gsk_YOUR_API_KEY_HERE";
