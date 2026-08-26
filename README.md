# Grounds Maintenance Hub

SharePoint Framework web part for the Waverley Grounds Maintenance team.

## SharePoint site
https://wavind.sharepoint.com/sites/Waverley

## Lists used
- GM Customers
- GM Sites
- GM Crews
- GM Jobs

## Build in GitHub (no Node needed on your PC)
1. Create a new GitHub repository.
2. Upload all files/folders from this ZIP to the root of the repo.
3. Open **Actions**.
4. Open **Build SPFx Package**.
5. Click **Run workflow**.
6. When it finishes, open the completed workflow run.
7. Download the artifact named **grounds-maintenance-hub-sppkg**.
8. Inside is `grounds-maintenance-hub.sppkg`.
9. Send that `.sppkg` to your SharePoint admin for App Catalog deployment.

## What the app does
- Reads Customers, Sites and Crews from SharePoint.
- Reads/writes Jobs directly to GM Jobs.
- 14-day schedule across active crews.
- Job status: Scheduled / Completed / Invoiced / Cancelled.
- Automatically stamps Completed Date / Invoiced Date when status changes.
- Dashboard revenue includes monthly flat contracts once and variable jobs separately.

## Build stack
- SPFx 1.23.2
- Node.js 22
- Heft toolchain
