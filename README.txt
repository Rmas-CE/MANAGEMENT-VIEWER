VAULTSYNC — MANAGEMENT VIEWER (view only, no login)
=====================================================

The viewer reads from its OWN Google Sheet ("VaultSync Viewer Mirror").
That sheet copies the Inventory data every 5 minutes. Bosses never connect to the
original Inventory Google Sheet, and the original sheet / Code.gs is NOT changed.

   Original Inventory Sheet  --(copy every 5 min, read only)-->  Viewer Mirror Sheet  <-- bosses' viewer

Files
  index.html       VaultSync hub (view only)
  inventory.html   POSM Inventory + Requests, Metrics, Beginning Stock, Movement — Excel / PDF downloads
  scorecard.html   TS POSM KPI Grading — Download Excel / Download PDF
  ViewerMirror.gs  Script for the new Viewer Mirror Google Sheet

SETUP
1) New Google Sheet for the viewer
   a. Create a new Google Sheet: "VaultSync Viewer Mirror"
   b. Extensions › Apps Script › delete the sample code › paste ViewerMirror.gs › Save
   c. Reload the sheet › menu "VaultSync Viewer" › Set up (first time)
      Allow permissions › type the Inventory Tool Admin password
      (stored only inside this script — never on the viewer page)
   d. Deploy › New deployment › Web app › Execute as: Me › Who has access: Anyone › Deploy
   e. Copy the Web app URL (ends in /exec)

2) inventory.html
   Open in Notepad, search for  PASTE_VIEWER_MIRROR_WEB_APP_URL_HERE  and replace it with the URL from 1e. Save.

3) GitHub (account Rmas-CE)
   New public repo e.g. VS-BOSS-VIEW › upload index.html, inventory.html, scorecard.html
   Settings › Pages › main / root › Save
   Link for the bosses: https://rmas-ce.github.io/VS-BOSS-VIEW/

Everyday
- The "Status" tab in the Viewer Mirror sheet shows the last refresh and any error.
- The viewer shows "View only · data as of <time>".
- Need it now? Viewer Mirror sheet › menu › Refresh now.
- Admin password changed? Viewer Mirror sheet › menu › Set / change Admin password.
  (Until then the viewer keeps showing the last good copy.)
- Proof photos are fetched only when a boss opens one, then kept for 6 hours.

What bosses can NOT do: stock in/out, scan, transfers, new requests, approvals, status changes,
proof verify/reject, add/edit/delete POSM/TS/stores, settings, passwords, reset, import.
The Viewer Mirror refuses all of these too.

Whoever has the link can view. Share it with management only.
