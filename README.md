# AI Agent — Automated Food Image Collection & Processing

Reads food items from Excel, searches online for images, evaluates/selects, resizes to **1800×1200**, compresses under **10 MB**, renames by item name, saves locally, and optionally uploads to Google Drive.

## Quick start

```bash
pip install -r requirements.txt
python agent.py --excel sample_food_items.xlsx
```

Outputs:
- `output_images/` — processed JPGs named like `Dal Tadka.jpg`
- `processing_report.csv` — status per item

### Full menu (211 unique items)

```bash
python agent.py --excel input_food_items.xlsx --delay 1.5
```

Limit for a faster demo:

```bash
python agent.py --excel input_food_items.xlsx --limit 10
```

## Google Drive upload (optional)

1. Create a Google Cloud service account and enable Drive API.
2. Share your target Drive folder with the service account email (Editor).
3. Save the JSON key as `credentials.json` in this folder.
4. Run:

```bash
python agent.py --excel sample_food_items.xlsx --drive-folder-id YOUR_FOLDER_ID
```

Or set env vars: `DRIVE_FOLDER_ID`, `GOOGLE_APPLICATION_CREDENTIALS`.

Without credentials, images stay in `output_images/` (workflow still completes).

## Files

| File | Purpose |
|------|---------|
| `agent.py` | Automation agent |
| `sample_food_items.xlsx` | Small demo list |
| `input_food_items.xlsx` | Full unique list from assignment sheet |
| `processing_report.csv` | Generated report |
| `output_images/` | Processed images |

## Workflow

`Excel → Search (DuckDuckGo) → Evaluate → Resize 1800×1200 → Rename → Local/Drive → CSV report`
