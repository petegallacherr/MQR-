MQR & DCT Consult — Draft 0.6.2

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
