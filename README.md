# addie

Google Apps Script for cold job-hunting emails. You add leads in a Sheet, pick a first-email template (or generate one with Groq), and the script schedules sends, then follows up twice if nobody replies.

The runnable script is [`apps-script/Code.gs`](apps-script/Code.gs). Paste that one file into a container-bound Apps Script on your leads spreadsheet (Extensions → Apps Script).

## What changed from the original script

| You asked for | How it works now |
| --- | --- |
| Multiple first-email templates | One Google Doc with several `Template:` blocks. Sheet column **Template** picks which block to use for email 1. |
| Auto sequence if no reply | Email 1 → wait → `followup_1` → wait → `followup_2`. A reply from a To/CC address sets **Reply** to Yes and stops. Follow-ups reply on the same Gmail thread. |
| Multiple To / CC / BCC | **Email**, **CC**, and **BCC** accept comma- or semicolon-separated addresses. |
| Runtime generation with Groq | Fill **Generate Context** (for example `write this using my experience at XYZ`). Groq drafts email 1; **Subject Choice** still picks among the subjects it returns. |
| Multiple subject lines | Repeat `Subject:` lines in the doc. Sheet **Subject Choice** is `1`, `2`, `3`, or the exact subject text. Blank uses the first. |

Scheduling is unchanged: new/pending rows get a send slot 12 minutes after the last one, inside your send window. Triggers still run `scheduleEmails` and `sendScheduledEmails`.

## One-time setup

1. Copy [`docs/sheet-headers.csv`](docs/sheet-headers.csv) into row 1 of `Sheet1` (or use the menu later to append missing columns).
2. Copy [`docs/google-doc-template.md`](docs/google-doc-template.md) into your Google Doc. Keep the `Template:` / `Subject:` / `Body:` markers. Put the doc ID in `CONFIG.DOC_ID` inside `Code.gs`.
3. In the Apps Script editor: Services (+) → add **Gmail API**.
4. File → Project settings → set the script timezone to match the spreadsheet.
5. Reload the Sheet. Menu **Cold DM Automation** → **Add missing sheet columns**.
6. If you want Groq: menu → **Save Groq API key** (stored in Script Properties, not in the file).
7. In the editor, select `setupTriggers` → Run → authorize.

After that, add rows and leave **Status** blank or `Pending`. Everything else is automatic. Use **Force-send test email now** to send one row immediately without changing Status.

## Sheet columns

| Column | Used for |
| --- | --- |
| Email | To: one or more addresses |
| CC / BCC | Optional extra recipients |
| Name / Company / Role | `{{first_name}}`, `{{name}}`, `{{company}}`, `{{role}}` |
| Template | First-email template id (`default`, `xyz_experience`, …). Blank = `default`, or the first non-follow-up template |
| Subject Choice | `1` / `2` / `3` or exact subject text |
| Generate Context | If set, Groq writes email 1. Leave blank to send the doc template as-is |
| Status | `Pending` → `Scheduled` → `Sent` → `Replied` or `Completed` (or `Error`) |
| Sequence Step | `1` first email, `2` followup_1, `3` followup_2 |
| Scheduled Time / Sent Time | Filled by the script |
| Next Follow-up | When the next sequence step may be queued |
| Thread ID / Original Subject | Used to reply in-thread and detect replies |
| Reply | `Yes` when a recipient replies |
| Error | Last send error |

Do not type `Scheduled` / `Sent` yourself. Leave new leads `Pending`.

### Example rows

- First email from the `default` template, subject line 2: Template=`default`, Subject Choice=`2`
- First email generated from context: Generate Context=`use my experience at XYZ; keep it under 120 words`
- Two recipients plus a CC: Email=`a@co.com, b@co.com`, CC=`recruiter@co.com`

## Google Doc shape

```
Template: default
Subject: {{first_name}}, quick question about {{company}}
Subject: Idea for {{company}}
Body:
Hi {{first_name}},
...

Template: followup_1
Subject: checking in
Body:
...

Template: followup_2
Subject: closing the loop
Body:
...
```

`followup_1` and `followup_2` names must match `CONFIG.FOLLOWUP_TEMPLATE_IDS`. You can add as many other first-email templates as you want.

## Config you may want to edit

In `CONFIG` at the top of `Code.gs`:

- `SEND_DAYS`, `START_HOUR`, `END_HOUR`, `INTERVAL_MINUTES` — send window and spacing
- `FOLLOWUP_DELAY_DAYS` — `[3, 7]` means follow-up 1 three days after email 1, follow-up 2 seven days after follow-up 1
- `GROQ_MODEL` — default `llama-3.3-70b-versatile`

## Status values

- **Pending** — waiting for a send slot (new lead or a due follow-up)
- **Scheduled** — has a send time
- **Sent** — last message went out; waiting for a reply or the next follow-up delay
- **Replied** — a To/CC address wrote back; sequence stopped
- **Completed** — both follow-ups sent with no reply
- **Error** — send failed; use **Reset error rows** to put them back to Pending
