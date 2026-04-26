# RExcelBridge cleanup report

Safe cleanup pass based on the working fast-transfer build.

## Changed
- Removed `.vs/`, `bin/`, and `obj/` from the zip.
- Kept internal folder name as `RExcelBridge/`.
- Kept all public Excel functions intact.
- Bumped project version to `2026.04.26.2`.
- Preserved fast numeric, typed table, RCall auto-dispatch, and RObj support.

## Deliberately not removed
- Public functions that existing workbooks may use.
- General JSON/text fallback path for scalars, strings, lists, messages, summaries, and unsupported objects.

## Porting note
Use the same layered structure for Python and Julia:
Excel layer -> worker process -> dispatcher -> numeric binary path -> typed table path -> general fallback.
