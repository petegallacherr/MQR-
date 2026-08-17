MQR & DCT Consult — Draft 0.6.4

Changes from 0.5.1:
- Removed the collapsible “Additional report details” section.
- DCT prescription details are always visible and highlighted.
- Missing prescription details do NOT block report generation.
- Prescription fields show a live incomplete/complete status.
- Added prescription details to the consultation facts when supplied; blanks show N/A.
- DCT recommendation now checks ALL current-season herd tests.
- Highest current-season SCC sets the SCC DCT band:
    below sealant cutoff -> Teat Sealant only
    at/above sealant cutoff but below LA cutoff -> SA DCT + Sealant
    at/above LA cutoff -> LA DCT + Sealant
- Dated MINDA imports may auto-detect up to 10 current-season herd tests.
- Service-worker cache bumped to v060.

Testing: publish the folder contents to GitHub Pages and hard refresh once after upload.

Draft 0.6.4 fixes:
- Rebased on the complete 0.6.2 form so no prescription details are lost.
- Dairy company dropdown retained; Dairy Goat Co-operative removed for cattle-only use.
- Planned start of calving auto-fills to 1 August of the year after the season start year and remains editable.
