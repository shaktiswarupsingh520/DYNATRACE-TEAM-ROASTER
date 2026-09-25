# Dynatrace Team Roster

Shared Axis Bank Dynatrace support roster designed from the existing spreadsheet-style monthly roster.

## Access model
The roster is stored as a Dynatrace Document using external ID `dynatrace-team-roster`.
- The owner creates and updates the document.
- The document is made public/read-only, which gives read access to users in the environment.
- The UI shows **Owner edit access** only when the current user's document permissions include `write`.
- Other users get **View only** and cannot update the document.
- Optimistic locking and snapshots protect concurrent updates and provide change recovery.

Dynatrace's Document service supports public documents for tenant-wide read access while write access remains controlled by document permissions/ownership. See the official SDK documentation.

## First-time setup
Open the app as the roster owner and select **Initialize Shared Roster** once. This creates the document and switches it to tenant-wide read access.

Required app scopes:
- `document:documents:read`
- `document:documents:write`

## Layout
The app follows the current roster pattern: team members as rows, calendar dates as columns, and shift codes G/E/M/W/L/H with a legend. The month is September 2026 and can be generalized to a month selector in the next iteration.

## Deployment
```bash
npm install
npm run typecheck
npm run build
npx dt-app analyze
npx dt-app deploy --dry-run
npx dt-app deploy
```
