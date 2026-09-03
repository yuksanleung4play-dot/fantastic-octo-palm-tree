Attribute VB_Name = "TestPDriveVisible"
' 診斷：這個 Excel 進程看不看得到對映磁碟機 P:。
' 請匯入早班_LME_reference_2024.xlsm（VBA 編輯器 → 檔案 → 匯入檔案）。
' Python 在跑 lme_main 之前會 Application.Run TestPDriveVisible。

Option Explicit

Public Function TestPDriveVisible() As String
    If Dir("P:\", vbDirectory) = "" Then
        TestPDriveVisible = "P: 磁碟機看不到！"
    Else
        TestPDriveVisible = "P: 磁碟機正常，內容：" & Dir("P:\")
    End If
End Function
