MQR & DCT Consult — Draft 0.7

v0.7 changes:
- Added the farmer-facing Clinical mastitis — monthly & seasonal trend graph based on the Koroa end-of-season Excel report.
- Graph runs June to May and shows:
    Heifer % (heifer cases / first calvers)
    MA cows % (mature-age cases / [peak cows - first calvers])
    Monthly % (all clinical case events / peak cows)
    Monthly SmartSAMM-style trigger line
    Cumulative Season % on a second Y-axis
- Retains the spreadsheet trigger pattern: Jun 1.0%, Jul 2.5%, Aug 4.0%, Sep 2.5%, Oct-May 1.0%; season reference 14%.
- Hover/tap/focus on a month shows the underlying figures.
- Graph is inline SVG so it remains sharp in Print / Save PDF and does not require an online chart library.
- Existing monthly detail table remains directly below the graph.
- SCC cure/new/retained logic and DCT recommendation logic are unchanged from 0.6.4.
- Service-worker cache bumped to v070.

Carried forward from 0.6.4:
- DCT recommendation checks ALL current-season herd tests.
- Highest current-season SCC sets the SCC DCT band.
- Dated MINDA imports may auto-detect up to 10 current-season herd tests.
- Prescription details remain visible and do not block report generation when incomplete.
- Planned start of calving auto-fills to 1 August of the year after the season start year and remains editable.

Testing / deployment:
Publish the folder contents to GitHub Pages and hard refresh once after upload so the v070 service-worker cache replaces the prior version.
