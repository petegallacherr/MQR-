MQR & DCT Consult — Draft 0.7.27

Version 0.7.27

0.7.27 changes:
- Added an in-app first-time user guide covering consultation setup, MINDA/Mastaplex exports, upload safeguards, and report outputs.
- Added a clear required/optional file table for Herd/SCC, full MINDA Treatment Register, previous DCT/ITS, and Mastaplex/Mastatest.
- Clarified that the full unfiltered MINDA Treatment Register is safe to upload; non-mastitis treatment rows do not need to be manually removed.
- Added guidance for season-year checks, dated vs legacy SCC headings, historic animals, missing SCC data, and Mastaplex seasonal filtering.
- Clarified that individual-cow recommendations are exported separately as CSV and intentionally excluded from the farmer PDF.
- Updated the Treatment Register upload-card wording to match the validated full-register workflow.
- No mastitis, SCC, DCT, SMARTSAMM, Mastaplex, AMR, or order-estimate calculation changes.
- Service-worker cache bumped to v0727.


0.7.26 changes:
- Rebalanced the monthly mastitis detail table for easier on-screen and PDF reading.
- Shortened the two long treatment headings to "Clinical cases" and "NSAID tx only".
- Moved "Included in trigger %" / "Not included in trigger %" onto smaller second-line sub-headings.
- Gave the two case columns matching widths so they remain visually balanced.
- No mastitis case-counting, SMARTSAMM trigger, SCC, DCT, Mastaplex, or AMR calculation changes.
- Service-worker cache bumped to v0726.


0.7.25 changes:
- Clarified the monthly mastitis trend table headings.
- "Clinical cases — included in trigger %" identifies the case events used in the SMARTSAMM monthly percentage.
- "NSAID tx only — not included in trigger %" makes it explicit that NSAID-only treatment records are displayed for context but excluded from the trigger percentage.
- No mastitis case-counting, SMARTSAMM trigger, SCC, DCT, Mastaplex, or AMR calculation changes.
- Service-worker cache bumped to v0725.


0.7.24 changes:
- Clarified mastitis cases that do not match the current herd file.
- Treatment-register records marked H in MINDA's IsHistoricAnimal column are now carried through to the grouped mastitis case event.
- The Individual cow case review now separates Historic-animal mastitis cases (H) from Other unmatched mastitis cases.
- Historic-animal cases remain included in the selected-season mastitis totals because the treatment event occurred during the season, but they remain outside current-herd DCT recommendations when they are not present in the herd file.
- Other unmatched cases are explicitly flagged for cow-ID / herd-file coverage review.
- No SCC, DCT threshold, monthly trigger, Mastaplex, or AMR calculation changes.
- Service-worker cache bumped to v0724.
