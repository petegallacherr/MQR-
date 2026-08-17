MQR & DCT Consult — Draft 0.7.1

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
Publish the folder contents to GitHub Pages and hard refresh once after upload so the v071 service-worker cache replaces the prior version.
