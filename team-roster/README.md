# Dynatrace Team Roster

Dynatrace AppEngine-ready React app for maintaining the Axis Bank Dynatrace operations team roster.

Features:
- Add, edit and remove team members
- Search and status filtering
- Team summary cards
- CSV export
- Responsive UI
- Browser persistence for the MVP

Axis environment: https://axis-prod.apps.dynatrace.com/

From the team-roster directory:
npm install
npm run typecheck
npm run build
npx dt-app analyze
npx dt-app deploy --dry-run
npx dt-app deploy

Do not commit Dynatrace credentials. A platform token or OAuth credentials with the required AppEngine deployment permissions must be available to the CLI.

For shared multi-user persistence, the next iteration should move roster records to the Dynatrace Document service with document:documents:read/write permissions.