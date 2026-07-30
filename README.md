# MapOTCClientFunds VBA

Excel VBA module that maps OTC client fund data onto the OTC sheet.

## Encoding fix (v6)

Traditional Chinese VBA editors use **Big5**. Sheet headers may be Simplified Chinese (**GBK**), and Chinese string literals of either kind can be corrupted depending on editor/codepage.

**All Chinese strings** (Simplified headers, Traditional headers, sheet name, MsgBox text) are built with `ChrW()` Unicode code points via helpers:

- `U()` — concatenate code points
- `SC_*` — Simplified Chinese header names
- `TC_*` — Traditional Chinese header names
- `Msg_*` / `SheetOTCClientFunds` — UI / sheet name strings

Runtime code contains **zero** Chinese string literals, so the module is safe to paste into a Big5 VBA editor.
