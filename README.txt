MQR & DCT Consult — Draft 0.7.24

Version 0.7.24

0.7.24 changes:
- Clarified mastitis cases that do not match the current herd file.
- Treatment-register records marked H in MINDA's IsHistoricAnimal column are now carried through to the grouped mastitis case event.
- The Individual cow case review now separates Historic-animal mastitis cases (H) from Other unmatched mastitis cases.
- Historic-animal cases remain included in the selected-season mastitis totals because the treatment event occurred during the season, but they remain outside current-herd DCT recommendations when they are not present in the herd file.
- Other unmatched cases are explicitly flagged for cow-ID / herd-file coverage review.
- No SCC, DCT threshold, monthly trigger, Mastaplex, or AMR calculation changes.
- Service-worker cache bumped to v0724.
