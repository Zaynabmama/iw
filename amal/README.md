# Amal PDF to Excel

This project prepares a simple Streamlit workflow for:

- Uploading 2 PDF files
- One `SOB` PDF
- One `PO` or `Commercial Invoice` PDF from IBM
- Producing one Excel workbook with 2 sheets:
- `comm-inv`
- `pack_list`

The upload and export flow is ready. The field extraction and final sheet mapping are intentionally left as placeholders until the exact sheet structure is provided.

## Run

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Run the commands from inside the `amal` folder.

## Current status

- Upload UI is ready
- File-type validation is ready
- Workbook generation is ready
- `comm-inv` sheet placeholder is ready
- `pack_list` sheet placeholder is ready
- PDF text extraction skeleton is ready
- Business mapping rules are pending your final column structure

## Expected next step

Share the exact columns and sample row structure for:

- `comm-inv`
- `pack_list`

Once you send that, we can wire the extraction and export logic into the prepared modules.
