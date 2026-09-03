' VBA wrapper：讓 Python 可用 Application.Run 直接傳入兩個日期，避開 InputBox。
' 請把原巨集本體改名（例如 RunDailyLME_Core），或把下列邏輯併入現有 Sub。
'
' TODO: 將 Macro 名稱對齊 config.yaml 的 vba.macro_name（預設 RunDailyLME）

Option Explicit

Public Sub RunDailyLME(Optional ByVal PrevDate As Variant, Optional ByVal ThreeMDate As Variant)
    If IsMissing(PrevDate) Or IsEmpty(PrevDate) Or CStr(PrevDate) = "" Then
        PrevDate = InputBox("請輸入上日日期", "LME Daily")
        If CStr(PrevDate) = "" Then
            MsgBox "未輸入上日日期，已取消。", vbExclamation
            Exit Sub
        End If
    End If

    If IsMissing(ThreeMDate) Or IsEmpty(ThreeMDate) Or CStr(ThreeMDate) = "" Then
        ThreeMDate = InputBox("請輸入 3M date", "LME Daily")
        If CStr(ThreeMDate) = "" Then
            MsgBox "未輸入 3M date，已取消。", vbExclamation
            Exit Sub
        End If
    End If

    ' TODO: 呼叫原本的處理流程，例如：
    ' Call RunDailyLME_Core(CStr(PrevDate), CStr(ThreeMDate))
    Debug.Print "PrevDate=" & CStr(PrevDate) & " ThreeMDate=" & CStr(ThreeMDate)
End Sub
