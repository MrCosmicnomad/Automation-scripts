# D-DART | Driver Detention Analysis & Review Tool

**Version:** 10.8  
**Author:** Sachin Vallakati  
**Platform:** Browser userscript  
**Environment:** Authorized internal logistics systems

## Overview

D-DART is a browser-based tool designed to streamline driver detention analysis and review. It evaluates order execution data, applies shipper-specific Statement of Work (SOW) rules, calculates potential detention charges, and helps users review or process detention-related actions.

The tool provides batch processing, approval workflows, lane validation, progress tracking, error handling, and a consolidated results interface.

## Features

- **Automated detention analysis:** Evaluates execution timestamps, free-time thresholds, stop types, and applicable SOW rules.
- **Batch processing:** Analyzes multiple order IDs in a single session, with configurable internal chunking and progress tracking.
- **Review and Action modes:** Review Mode displays proposed outcomes without performing detention actions; Action Mode can execute supported changes.
- **Lane validation:** Compares the contracted SMC lane with FMC execution stop codes to identify mismatches.
- **SOW configuration:** Retrieves shipper-specific rules from an authorized SharePoint configuration source.
- **Charge validation:** Identifies chargeable orders, existing charges, potential charge discrepancies, and cases requiring further review.
- **Approval workflow:** Supports approval, rejection, skipping, and authorization-number requirements where configured.
- **Hold management:** Supports applicable detention hold creation or release workflows.
- **Resilience:** Includes retry handling, rate-limit responses, circuit breakers, caching, and version-conflict handling.
- **Progress persistence:** Saves batch progress to help manage interrupted processing sessions.
- **Diagnostics:** Provides logs, health checks, performance metrics, and debug reporting.

## How It Works

1. The user opens an authorized logistics application.
2. D-DART initializes its interface and loads the available SOW configuration.
3. The user enters one or more order IDs.
4. The tool retrieves relevant order and execution information.
5. Detention eligibility, timestamps, free time, existing charges, and lane consistency are evaluated.
6. Results are presented for review or processing, depending on the selected operating mode.
7. The tool generates processing results and reports successful, pending, skipped, or failed actions.

## Installation

### Requirements

- A supported desktop browser.
- Tampermonkey or another compatible userscript manager.
- Authorized access to the internal applications and configuration sources used by the script.
- An authenticated session with the required permissions.

### Install the userscript

1. Open your browser's userscript manager.
2. Create a new userscript.
3. Open the supplied D-DART `.user.js` file in a text editor.
4. Copy the complete script into the userscript editor.
5. Save the script and ensure it is enabled.
6. Navigate to an authorized application page supported by the userscript.
7. Confirm that D-DART initializes and reports its configuration status.

Use only an approved installation method and an environment authorized for this tool.

## Usage

### 1. Start in Review Mode

D-DART defaults to **Review Mode**.

Use this mode to inspect analysis results and understand the actions the tool would propose without executing those actions.

### 2. Enter order IDs

Enter valid order IDs in the input panel. Multiple IDs can be separated by commas, spaces, or new lines.

The tool validates the input and removes duplicate IDs before processing.

The configured maximum batch size is 2,000 orders per session. Large batches are processed in chunks.

### 3. Review the results

Depending on the available data, results can include:

- Chargeable detention
- Existing charges or charge discrepancies
- Missing arrival or departure timestamps
- Orders within free time or below the applicable threshold
- Cancelled or invoiced orders
- Lane mismatches
- Missing or incomplete SOW configuration
- Pending approval or processing errors

Review the underlying order details before deciding what action is appropriate.

### 4. Use Action Mode cautiously

Action Mode enables supported processing operations, which may include updating orders, adding detention pricing lines, adding comments, or releasing holds.

**Important:** Action Mode can make changes to internal business records. Confirm that the analysis, lane, applicable SOW rules, charge amount, and required approvals are correct before using it.

### 5. Monitor batch processing

Follow the batch status and review the final results. If a batch is interrupted or encounters errors, inspect the reported state before deciding whether to retry it.

Avoid submitting overlapping batches for the same orders without first checking their current status.

## Configuration

D-DART loads shipper-specific Statement of Work settings from its configured SharePoint source.

Configuration may determine:

- Pickup and drop-off eligibility
- Live-load and drop-and-hook rules
- Free-time thresholds
- Automatic charge permissions
- Approval requirements
- Authorization-number requirements

If configuration is unavailable, disabled, or incomplete, do not assume that a charge is valid. Follow the applicable internal process to resolve the configuration issue.

## Keyboard Shortcuts

The script defines the following keyboard shortcuts:

| Key | Function |
|---|---|
| `Escape` | Minimize or toggle the panel |
| `Enter` | Analyze orders |
| `R` | Reset the form |
| `S` | Open settings |
| `D` | Open debug functionality |

Shortcut behavior may depend on the currently focused control or interface state.

## Troubleshooting

**Authentication errors**
- Confirm that your internal session is active.
- Refresh the application and sign in again if required.
- Verify that your account has the necessary permissions.

**SOW configuration errors**
- Confirm that the configuration source is accessible.
- Check whether the shipper's configuration is enabled and complete.
- Retry loading the configuration after resolving any access issue.

**Missing execution data**
- Verify the order ID and available execution records.
- Review missing timestamp or execution-leg warnings.
- Investigate the order through the approved internal workflow.

**Rate limits or service errors**
- Allow the tool's retry and cooldown handling to complete.
- Avoid repeatedly submitting the same batch while requests are still processing.
- Check the debug report for the reported failure.

**Unexpected charge or lane result**
- Return to Review Mode.
- Verify the contracted lane, execution stop codes, timestamps, and applicable SOW configuration.
- Do not rely on the automated result as a substitute for required business review.

## Technical Notes

The script uses browser userscript APIs for authenticated HTTP requests, local userscript storage, styling, and clipboard operations. It also includes internal caching, batch-state management, telemetry metrics, and diagnostic logging.

The configured maximum batch size is 2,000 orders per session, with a chunk size of 50 orders.

## Security and Responsible Use

- Use this tool only with explicit authorization for the relevant systems and workflows.
- Do not publish internal endpoints, configuration details, operational workflows, or sensitive business data without approval.
- Never commit credentials, session cookies, authentication tokens, customer information, order identifiers, or unredacted debug reports.
- Review the script's network permissions and data handling before installing or distributing it.
- Keep Action Mode restricted to appropriately authorized use.
- Review changes to charges and holds according to the applicable business procedures.

## License

The source file declares an MIT license. Confirm that you have permission to distribute the code and that the license is approved for the intended repository before publishing it.

---

*D-DART v10.8 | Driver Detention Analysis & Review Tool*
