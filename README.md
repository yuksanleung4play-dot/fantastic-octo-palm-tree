# MapOTCClientFunds VBA

Excel VBA module that maps OTC client fund data onto the OTC sheet.

## Encoding fix (v5)

Traditional Chinese VBA editors use **Big5**, while the source sheet headers are **Simplified Chinese (GBK)**. Writing Simplified strings as literals causes mojibake and header matching fails.

All Simplified Chinese header strings are built with `ChrW()` Unicode code points (via helper `U()` / `SC_*` functions), so the module is safe to paste into a Big5 VBA editor. Traditional Chinese literals (native to Big5) are unchanged.
