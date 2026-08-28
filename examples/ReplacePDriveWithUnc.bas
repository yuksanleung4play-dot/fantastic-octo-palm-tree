Attribute VB_Name = "ReplacePDriveWithUnc"
' 警告：不要把 QueryTables.Connection 的 P:\ 永久改成 UNC。
' Excel TEXT 連線（TEXT;\\server\share\...\lme.yyyymmdd.dat）常 1004 找不到檔案，
' 即使檔案總管 / Python 都看得到同一個 .dat。請維持：
'
'   TEXT;P:\Dealing Department - New\1. 交易部日常工作分類\2. 期貨\LME SPAN\lme.20260827.dat
'
' Python 預設會把誤改的 UNC 改回 P:（記憶體、不存檔）。
' vba.rewrite_p_drive_to_unc: true 才會再把 P: 改成 UNC（不建議）。

Option Explicit
