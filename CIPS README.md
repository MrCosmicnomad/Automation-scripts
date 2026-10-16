# CIPS Auto Collator

A Python automation tool for collating CIPS Create, Update, and Subscribe tickets, consolidating CSV data into Excel master files, checking scancodes against IMP through a browser bridge, and generating a pending-items workbook.

## Features

- **Automated ticket retrieval:** Fetches pending CIPS tickets matching configured ticket categories, statuses, and requester identity.
- **CSV processing:** Downloads CSV attachments and maps their data to the appropriate master worksheet columns.
- **Data consolidation:** Merges new records into the master workbook and deduplicates records by scancode.
- **Ticket management:** Creates or updates internal tracking tickets with SKU counts and resolves processed source tickets.
- **IMP verification:** Checks scancodes through a local browser bridge and classifies items as done or pending.
- **Pending workbook generation:** Creates a separate Excel workbook containing items that still require attention.
- **Caching and recovery:** Uses local cache files, a run log, and a master workbook backup to support repeated runs.

## Requirements

- Windows
- Python 3.9 or a compatible version supported by the installed dependencies
- Access to the required internal services and authorization to perform ticket operations
- A valid Midway authentication session
- The required Excel workbooks and supporting files
- A compatible browser bridge userscript for IMP checks

### Python dependencies

Install the required packages using:

```bash
python -m pip install requests openpyxl
```

## Required Files

Keep the following files in the same directory as the Python script:

| File | Purpose |
|---|---|
| `cips_collator.py` | Main automation script |
| `CIPS Template.xlsx` | Workbook template containing the expected worksheet headers |
| `MASTER Collation File.xlsx` | Master workbook used to read and consolidate records |
| `Exclude_scancodes.txt` | Optional list of scancodes to exclude from the pending workbook |
| `Run CIPS Collator.bat` | Optional Windows launcher, if provided |

The script also creates local cache and log files during execution.

## Authentication and Browser Setup

1. Initialize the required Midway authentication session using the approved internal procedure.
2. Confirm that the authentication cookie is available at the expected location in your Windows user profile.
3. Ensure you are authorized to access the required ticketing services.
4. Install and enable the compatible IMP browser bridge userscript.
5. Sign in to Grocery Central in your browser and ensure the bridge is permitted to communicate with the local service on port `5050`.

The automation depends on valid authentication and a working browser bridge. Authentication failures or missing browser results can prevent successful processing.

## How to Run

1. Close `MASTER Collation File.xlsx` in Excel to avoid file-locking issues.
2. Confirm that the template and master workbook exist in the script directory.
3. Initialize authentication.
4. Run the script using Python:

   ```bash
   python cips_collator.py
   ```

   Alternatively, double-click `Run CIPS Collator.bat` if the launcher is available and configured correctly.

5. Follow the console output to monitor ticket retrieval, processing, deduplication, IMP verification, and workbook generation.
6. Review the generated workbook and ticket-operation results before handing off pending items.

## Generated Files

| File or directory | Purpose |
|---|---|
| `Pending CIPS.xlsx` | Workbook containing items that remain pending |
| `MASTER Collation File.xlsx.bak` | Backup of the master workbook before it is overwritten |
| `.cips_cache/` | Cached ticket data used to reduce repeated downloads |
| `.cips_cache/run_log.json` | Run history and tracking information for internal tickets and resolved source tickets |

Keep the generated files available between runs when the automation relies on previously recorded state.

## Important Processing Rules

- Tickets are filtered using the configured ticket criteria and requester identity.
- Create, Update, and Subscribe records are handled as separate streams.
- Duplicate records with a scancode are removed during consolidation, retaining the first encountered occurrence.
- Items that cannot be verified through IMP are conservatively treated as pending.
- Excluded scancodes are omitted from the pending handoff workbook, but may still be included in backlog calculations.
- Source-ticket resolution is subject to a confirmation prompt.

## Troubleshooting

**Authentication failed**
- Reinitialize the required authentication session and retry.

**Template or master workbook missing**
- Check that the required workbook is in the same directory as the script.

**Master workbook is locked**
- Close the workbook in Excel and rerun the automation.

**IMP verification does not complete**
- Verify that the browser bridge is installed, the browser is authenticated, and local communication is allowed.
- Check whether the bridge returned any results before assuming the verification succeeded.

**Missing or unexpected pending records**
- Review the scancode values, master worksheet headers, exclusion list, and IMP results.

## Security and Usage

This tool is intended only for authorized environments.

- Do not commit authentication cookies, passwords, access tokens, customer data, or confidential operational data.
- Do not publish internal endpoints, ticket identifiers, usernames, or proprietary workflows without authorization.
- Review all ticket-creation, update, and resolution operations before running the automation in a production environment.
- Store logs, caches, workbooks, and backups securely.
