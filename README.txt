MQR & DCT Consult — Draft 0.5.1

Open index.html in Chrome or Edge.

What is new in 0.5
- Report-focused layout based on the completed consultation workbook.
- Corrected SCC dry-period classification using the same low/high threshold:
  below -> above = New Infection
  above -> below = Cured
  above -> above = Retained Infection
- Treatment rows are grouped into clinical case events by cow + treatment date.
- Clinical case timing from calving with explicit unavailable-data counts.
- Correct multi-quarter percentages (one clinical case = one denominator event).
- Data coverage panel so missing farmer data is not silently treated as zero.
- Previous DCT/ITS and Mastaplex report as N/A when not supplied.
- Updated DCT order and individual-cow report.

Important
This is a prototype decision-support build and still requires comparison against known completed consultations before clinical use.
CSV inputs work fully in-browser. XLS/XLSX support uses the SheetJS script referenced in index.html and therefore requires that library to be available when the page is first loaded.
