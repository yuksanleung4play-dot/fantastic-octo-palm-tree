Attribute VB_Name = "ReplacePDriveWithUnc"
' 把巨集裡硬編碼的 P:\Dealing Department - New\... 改成 UNC。
' Python 端跑巨集前也會嘗試改 QueryTable.Connection / VBA 模組（記憶體、不存檔）。
' 若要永久改好：在 VBA 編輯器搜尋 P:\Dealing Department - New\
' 全部取代成 \\192.168.89.167\Dealing\Dealing Department - New\
'
' QueryTables.Connection 範例：
'   舊：TEXT;P:\Dealing Department - New\1. 交易部日常工作分類\2. 期貨\LME SPAN\lme.20260827.dat
'   新：TEXT;\\192.168.89.167\Dealing\Dealing Department - New\1. 交易部日常工作分類\2. 期貨\LME SPAN\lme.20260827.dat

Option Explicit
