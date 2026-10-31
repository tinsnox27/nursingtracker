# Technical notes

## Source and purpose

This is a self-contained recreation for portfolio review. The original workbook is unavailable. The generated workbook, sample messages, and local demonstration give a reviewer enough context to inspect the workflow without employer records or account access.

The implementation includes improvements made during reconstruction: keys distinguish renewal dates and recipient roles, sends use a script lock, unresolved delivery states remain reviewable, and a missed scheduled run can catch up within the current reminder band. These are features of this recreation and are not asserted as exact historical details of the original system.

## Source schema

The three cohort tabs use the same first five columns:

| Column | Field | Requirement |
| --- | --- | --- |
| A | Nurse name | Display name; defaults to Colleague if empty. |
| B | Nurse email | Required single address. |
| C | Supervisor email | Optional single address. |
| D | Certification type | Required nonempty label. |
| E | Expiration date | Required actual Google Sheets date value. |

Email values are trimmed and normalized to lowercase. Credential values are trimmed; delivery keys also normalize case. A changed address, credential label, or expiration date creates a different key. An actual implementation needs a record-identity policy if those fields change frequently.

The exact configured red backgrounds suppress that individual recipient. Invalid required nurse data causes the record to be counted as invalid and skipped; a malformed optional supervisor address does not stop an otherwise valid nurse message. No automatic bounce processing is included.

## Delivery schema

`Reminder Delivery Log v2` is created only in sending mode. Fields: delivery key, status, update timestamp, recipient, role, credential, expiration, threshold, source tab and row, and detail.

Preview mode reads an existing log if present but never marks a message delivered. It deduplicates preview messages within the current run. Previewing twice intentionally shows the same unsent messages again.

`SENT` suppresses future delivery of that exact key. `PENDING` and `ERROR` require review; the workflow does not assume a failed API call proves no message was delivered. A script lock serializes runs within this project, not across independent copies of the script.

Email delivery and Sheets persistence are not a transaction. If Gmail succeeds and the status update fails, the persistent `PENDING` entry prevents an automatic resend. The owner must reconcile the record. Check actual delivery history before marking it `SENT` or removing a confirmed-unsent record for retry. Retain the audit history when doing so. Exactly-once delivery is not guaranteed.

The old four-column send log in the earlier gist did not include expiration dates. This recreation uses a separate log rather than guessing a migration. Its demonstration begins with synthetic data, so no migration is needed.

## Optional sending configuration

Sending is not needed to review this project. If adapting it for actual use, test with controlled inboxes and appropriate data access first. `DRY_RUN = false` enables the main function's Gmail calls. Running `createDailyTrigger` in that mode creates a daily trigger in the 7 a.m. hour of the project timezone. It does not run at an exact minute. `runDemo` and `previewDemoReminders` refuse to run while dry run is off.

The saved script property `TRACKER_SPREADSHEET_ID` selects the workbook created by the generator; otherwise the main workflow falls back to an active bound spreadsheet. Change that property deliberately if using a different workbook. The spreadsheet timezone controls date comparisons, while the project timezone controls scheduling.

`INFOGRAPHICS` may hold Drive file IDs for optional attachments. Missing or inaccessible attachments do not block a reminder. No attachment files or issuer-specific renewal instructions are required or supplied for the portfolio demo.

## Tests and boundaries

Run `npm test`. The fourteen original behavior checks and five end-to-end demo checks use Node's built-in modules and in-memory Google service adapters. No dependency installation is needed. The local adapters implement only methods used by this sample and are not general replacements for Google APIs.

The integration test invokes the same workbook generator and preview function supplied for Apps Script. Seven unique recipient messages are expected from eight synthetic source rows. It checks that a red-flagged nurse is suppressed while the supervisor remains eligible, a duplicate row is collapsed, and out-of-window or invalid records are excluded. The simulated sending run writes seven delivery rows; a second run writes none.

Tests do not validate Google consent, real email delivery, live quota behavior, actual formatting or conditional-formatting behavior, attachments, trigger execution, or the Apps Script editor. A real service deployment needs those checks. Large datasets may need incremental processing rather than a whole-workbook scan.

## Updating the GitHub project

Upload the contents of this repository folder, including `README.md`, `src`, `test`, `examples`, `docs`, and `package.json`. GitHub will render the README and its workflow diagram automatically. The included generated HTML email can be downloaded and opened locally; GitHub's code view displays its source.

No license was selected for you. Add a license if you decide to make reuse permissions explicit.
