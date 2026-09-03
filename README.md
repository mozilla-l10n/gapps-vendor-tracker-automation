# gapps-vendor-tracker-automation

Google Apps Script automation for the Mozilla l10n vendor spend tracker (Google Sheets). Two scripts connect the tracker to Jira and to the vendor's invoice PDFs, so request intake and invoice reconciliation no longer rely on manual copying.

Both scripts run from a custom "L10n Tools" menu inside the tracker spreadsheet. No local tooling is needed to use them; this repository is the source of truth for the code and its change history.

## Scripts

### `sync_from_jira.gs`

Pulls open and recently created vendor intake tickets from Jira into the tracker's `AllJobs` tab.

- New tickets become new rows: job name, cost center, target locales, dates, status, Jira link, PM.
- For tickets already tracked, the Jira-owned fields (deadline, locales, PM, cost center) are updated when the ticket changed in Jira. Every change is listed in the end-of-run popup as "old → new"; nothing changes silently.
- Sheet-owned fields (status, word count, estimate, invoiced amount, notes) are never touched. Job names are flagged when they differ, never auto-changed.
- Tickets for the other vendor (matched by component) and tickets closed without work are skipped and reported.

### `reconcile_invoices.gs`

Monthly invoice reconciliation in two steps, with a human review in between.

1. **Reconcile invoices…** asks for a month, reads that month's invoice PDFs from the shared Drive folder, extracts subject and total from each, and matches them to tracker rows. It writes a `Reconciliation` report tab with one line per invoice and a checkbox: pre-ticked for clean matches, unticked with a written reason for anything doubtful.
2. **Apply reconciliation** acts only on ticked lines: it fills empty invoiced cells in the tracker (never overwrites) and writes the amount to the Jira ticket's invoiced field, filling empty tickets, verifying matching ones, and flagging (not changing) any that disagree.

On months that were already filled in by hand, the same two steps verify PDF ↔ sheet ↔ Jira and surface mismatches.

## Repository layout

```
scripts/
  config.dist.gs        configuration template (copy to config.gs and fill in)
  config.gs             local deployment values, gitignored, never committed
  sync_from_jira.gs     Jira tickets → tracker rows
  reconcile_invoices.gs invoice PDFs → tracker + Jira amounts
```

## Deployment

The scripts live in the Apps Script project bound to the tracker spreadsheet (Extensions → Apps Script).

1. Copy `scripts/config.dist.gs` to `scripts/config.gs` and fill in the values (Jira URL, project key, field IDs, invoice folder ID). `config.gs` is gitignored and must never be committed.
2. In the Apps Script project, create one file per script (`config`, `sync_from_jira`, `reconcile_invoices`) and paste the contents.
3. Add the Jira credentials under Project Settings → Script properties: `JIRA_EMAIL` (the account email) and `JIRA_TOKEN` (a Jira API token, created at id.atlassian.com → Security → API tokens). Credentials live only in Script properties, never in code or in this repository.
4. Enable the Advanced Drive Service: left sidebar → Services (+) → Drive API → version v2 → Add. This is used to convert invoice PDFs to text.
5. Reload the spreadsheet. The "L10n Tools" menu appears with all three commands.

## Making changes

Edit the code in this repository first, review, then paste the updated file into the Apps Script project. Keep the two in sync: the Apps Script project is the deployment, this repository is the history.

Before committing, check that no deployment values or credentials have crept into the scripts: everything instance-specific belongs in `config.gs`.
