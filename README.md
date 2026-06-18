# Word Data Merger — Office Add-in

A Microsoft Word task-pane add-in that merges data from an uploaded Excel file into content controls in the open document. Handles both flat (single-record) fields and repeating table rows for line items.

---

## Quick Start

### 1. Install dependencies

```bash
npm install
```

### 2. Install dev certificates (first time only)

The add-in is served over HTTPS from localhost. Office requires this.

```bash
npm run install-certs
```

When prompted, allow the certificate to be trusted (macOS: enter your password).

### 3. Start the server

```bash
npm start
```

This starts an HTTPS server at `https://localhost:3000`. Keep this terminal open.

### 4. Open the sample template in Word

Open `samples/sample-template.docx` in Microsoft Word for Mac or Windows.

### 5. Sideload the add-in into Word

**Word for Mac / Windows (desktop):**

1. In Word, go to **Insert → Add-ins → My Add-ins**.
2. Click **Upload My Add-in** (or "Manage My Add-ins" → "Upload My Add-in").
3. Browse to and select `manifest.xml` from this project folder.
4. Click **Upload**.

The "Open Merger" button will appear in the **Home** ribbon tab under "Data Merger".

**Word on the web:**

1. Go to **Insert → Office Add-ins**.
2. Choose **Upload My Add-in** and select `manifest.xml`.

---

## Full Test Walkthrough

With `sample-template.docx` open in Word and the add-in sideloaded:

1. **Open the task pane** — click "Open Merger" in the Home ribbon.

2. **Upload data** — click "Choose .xlsx file…" and select `samples/sample-data.xlsx`.
   A preview table will appear showing all rows.

3. **Set flat data row** — the add-in defaults to row 2. Row 2 in the spreadsheet contains the invoice header data (`client_name`, `invoice_number`, `invoice_date`, `total_due`). Leave it as `2`.

4. **Set line-item rows** — defaults to rows 3–7. Rows 3–7 contain the five invoice line items. Leave as-is.

5. **Scan the document** — click "Scan Document & Map Columns". The add-in reads all content control tags from the open Word document and auto-maps them to matching Excel column headers.
   - Flat fields panel: `client_name`, `invoice_number`, `invoice_date`, `total_due`
   - Line-item fields panel: `item_description`, `item_qty`, `item_unit_price`, `item_total`
   - All should show green "auto" badges (exact name matches).

6. **Merge** — click "Merge into Document".

7. **Verify results:**
   - "Client:" → Globex Corporation
   - "Invoice #:" → INV-2026-0042
   - "Invoice Date:" → 2026-06-18
   - "Total Due:" → $4,725.00
   - The table should now have five data rows (template row is replaced):
     - Strategy Consulting — 10 hrs | 10 | $250.00 | $2,500.00
     - Market Research Report | 1 | $850.00 | $850.00
     - Brand Identity Package | 1 | $600.00 | $600.00
     - Social Media Setup | 3 | $95.00 | $285.00
     - Project Management Fee | 1 | $490.00 | $490.00

---

## How to Build Your Own Template

### Inserting a content control in Word

1. Enable the Developer tab: **Word → Preferences → Ribbon & Toolbar → Developer** (Mac) or **File → Options → Customize Ribbon → Developer** (Windows).
2. Place your cursor where you want a field.
3. In the **Developer** tab, click **Plain Text Content Control** (the "Aa" icon with a box).
4. With the control selected, click **Properties**.
5. Set the **Tag** field to your column name from Excel (e.g. `client_name`). The tag is what the add-in matches against.
6. Click OK.

### Flat fields

Use any tag name that matches an Excel column header (case-insensitive, spaces and underscores treated the same). Place these controls anywhere in the document.

### Repeating table rows (line items)

1. Create a table with a header row.
2. Add exactly **one** body row — this is the **template row**.
3. In each cell of the template row, insert a Plain Text Content Control.
4. Tag each control with `item_` followed by the Excel column name:
   - `item_description`, `item_qty`, `item_unit_price`, `item_total`
5. The add-in finds this template row automatically (by detecting `item_` prefixed tags), clones it for each data row, fills the clones, and deletes the original.

**Rules:**
- The table must contain at least one row with `item_` tagged controls.
- Only one table in the document should use `item_` tags.
- Each template cell should contain exactly one Plain Text Content Control.

---

## Project Structure

```
word_merger/
├── manifest.xml              # Office Add-in manifest
├── package.json
├── server.js                 # Local HTTPS dev server
├── src/
│   └── taskpane/
│       ├── taskpane.html     # Task pane UI
│       ├── taskpane.css      # Styles
│       └── taskpane.js       # All merge logic (Office.js + SheetJS)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
├── samples/
│   ├── sample-template.docx  # Invoice template with content controls
│   └── sample-data.xlsx      # Test data (flat row + 5 line items)
└── scripts/
    ├── generate_samples.py   # Regenerate sample files
    └── generate_icons.py     # Regenerate icon PNGs
```

---

## Regenerating Sample Files

```bash
npm run generate-samples
```

Requires Python 3 with `python-docx` and `openpyxl` installed:

```bash
pip3 install python-docx openpyxl
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Not trusted" SSL error in task pane | Run `npm run install-certs` and trust the cert |
| Task pane shows blank / won't load | Confirm `npm start` server is running; visit `https://localhost:3000/src/taskpane/taskpane.html` in a browser first |
| "No content control found with tag X" | Check the tag in Word Developer → Properties matches exactly (case insensitive but must have correct `item_` prefix) |
| Template row not removed / not found | Ensure only one row in the table has cells tagged `item_*`; ensure those controls are Plain Text type |
| Add-in doesn't appear in ribbon | Re-sideload: Insert → Add-ins → My Add-ins → Remove old one, upload again |
