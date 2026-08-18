MQR & DCT Consult — Draft 0.7.7


v0.7.7 changes:
- Added an on-screen legend above Individual cow recommendations explaining Data completeness terms and the dry-off timing method.
- Dry-off advice still uses the established BCS-adjusted calculation when BCS + expected calving are both available.
- If BCS is missing but expected calving is available, dry-off advice now falls back to a standard 50-day dry period.
- Renamed the on-screen Data note column to Data completeness.
- Updated data-coverage wording to reflect the dry-off fallback.

v0.7.6 changes:
- Report-header Export CSV button now uses the Vetlife tan treatment and Print / Save PDF uses the dark brown treatment for stronger visual separation.
- Individual cow recommendations are excluded from Print / Save PDF; they remain available on-screen and through the dedicated CSV export.
- Data note remains a cow-level completeness flag. It reports missing pregnancy, BCS, expected calving date, or partial current-season SCC data. Zero SCC values are now treated consistently as missing for the partial-SCC flag.
v0.7.5 changes:
- Individual cow recommendations CSV export is now clearly labelled and available beside the Individual cow recommendations table as well as in the report header.
- Export includes cow tag, latest SCC, pre-dry SCC, pregnancy diagnosis, expected calving date, BCS, individual DCT recommendation, dry-off advice, dry-period SCC status, data note and Mastaplex result.
- CSV filename now ends in -individual-cow-recommendations.csv and includes a UTF-8 BOM for cleaner opening in Excel.
- No SCC, mastitis, DCT, selective-treatment or Mastaplex calculations were changed.
- Service-worker cache bumped to v074.

v0.7.3 changes:
- Styling: Mastaplex cultured growths now spans the full report width.
- Gram-positive / Gram-negative grouping and all calculations are unchanged.

v0.7.2 changes:
- Mastaplex cultured-growth summary is split into Gram-positive and Gram-negative organism groups.
- No-growth, mixed/contaminated and unclassified Mastaplex results are kept separate rather than forced into a Gram category.
- Original organism/result names and counts remain visible within each group.
- Gram percentages use classified bacterial growths as the denominator.
- All v0.7.1 report ordering, selective DCT outcome calculations, SCC logic and DCT recommendation logic are unchanged.

v0.7.1 changes:
- Reordered the report immediately after the monthly mastitis summary.
- Mastitis treatments used is now first, with Quarter distribution beside it.
- Previous DCT / ITS treatment now sits directly underneath as a full-width selective-DCT review section.
- Added current-season mastitis outcome calculations for cows previously treated with DCT + sealant combo versus sealant only.
- Comparison denominators include only previous-treatment cows matched to the current herd file, so cows no longer in the herd do not artificially lower the mastitis rate.
- If there are no previous sealant-only cows, the selective comparison is shown as not available rather than zero.
- Existing SCC transition logic, current-season DCT recommendation logic, mastitis event grouping and v0.7 graph calculations are unchanged.
- Service-worker cache bumped to v071.

Selective DCT outcome definitions:
- Previous DCT + sealant combo = a cow with at least one recognised DCT product and at least one recognised teat-sealant product in the previous DCT / ITS file.
- Previous sealant only = a cow with a recognised teat-sealant product and no recognised DCT product in the previous DCT / ITS file.
- Mastitis this season = the cow has at least one current-season mastitis case event using the same event logic as the report.
- The comparison is observational and is not labelled as proof that previous treatment caused or prevented mastitis.

Deploy:
Publish the folder contents to GitHub Pages and hard refresh once after upload so the v073 service-worker cache replaces the prior version.


v0.7.5 fix: the lower Individual cow recommendations export button now triggers the already-working report-header export control directly. This remains functional even if the page has updated before the PWA has replaced an older cached app.js.
