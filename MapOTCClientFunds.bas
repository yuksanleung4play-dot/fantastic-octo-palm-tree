Attribute VB_Name = "MapOTCClientFunds"
' ==================================================
' MapOTCClientFunds - 完整版 v5
' 功能：
'   1. 動態識別 OTC客戶資金 表頭
'   2. 按賬號映射 Other Base 當日權益/初始保證金 至 OTC Sheet
'   3. 101758000 賬號計算匯率寫入 B1
'   4. 有 CNH 行的賬號，Excess 旁顯示 CNH/USD 可用明細
'   5. Excess 寫入公式 =Equity-Margin
'   6. 同天多次運行：Margin 低於上次則標紅
'   7. C14/H14/M14/R14 顯示上次導入 Margin 值
'   8. D14/I14/N14/S14 顯示上次導入時間
'
' 編碼說明（v5）：
'   繁體 VBA 編輯器為 Big5，無法正確保存簡體字串（GBK 表頭）。
'   所有簡體表頭比對字串一律用 ChrW() Unicode 碼位組出，繞開編碼問題。
'   繁體字串（編輯器原生支援）維持字面量寫法。
' ==================================================

' 以 Unicode 碼位組出字串（避免 Big5 編輯器破壞簡體字面量）
Private Function U(ParamArray cps() As Variant) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = LBound(cps) To UBound(cps)
        s = s & ChrW(cps(i))
    Next i
    U = s
End Function

' 簡體表頭字串（全部經 ChrW，對應 GBK Sheet 欄位名）
Private Function SC_KeHuZhangHao() As String
    ' 客户账号
    SC_KeHuZhangHao = U(&H5BA2, &H6237, &H8D26, &H53F7)
End Function

Private Function SC_ZiJinZhangHao() As String
    ' 资金账号
    SC_ZiJinZhangHao = U(&H8D44, &H91D1, &H8D26, &H53F7)
End Function

Private Function SC_KeHuZiJinZhangHao() As String
    ' 客户资金账号
    SC_KeHuZiJinZhangHao = U(&H5BA2, &H6237, &H8D44, &H91D1, &H8D26, &H53F7)
End Function

Private Function SC_BiZhongZuBie() As String
    ' 币种组别
    SC_BiZhongZuBie = U(&H5E01, &H79CD, &H7EC4, &H522B)
End Function

Private Function SC_BiZhongZu() As String
    ' 币种组
    SC_BiZhongZu = U(&H5E01, &H79CD, &H7EC4)
End Function

Private Function SC_BiZhongZuHao() As String
    ' 币种组号
    SC_BiZhongZuHao = U(&H5E01, &H79CD, &H7EC4, &H53F7)
End Function

Private Function SC_BiZhongHao() As String
    ' 币种号
    SC_BiZhongHao = U(&H5E01, &H79CD, &H53F7)
End Function

Private Function SC_BiZhong() As String
    ' 币种
    SC_BiZhong = U(&H5E01, &H79CD)
End Function

Private Function SC_DangRiQuanYi() As String
    ' 当日权益
    SC_DangRiQuanYi = U(&H5F53, &H65E5, &H6743, &H76CA)
End Function

Private Function SC_JinRiQuanYi() As String
    ' 今日权益
    SC_JinRiQuanYi = U(&H4ECA, &H65E5, &H6743, &H76CA)
End Function

Private Function SC_JinQuanYi() As String
    ' 今权益
    SC_JinQuanYi = U(&H4ECA, &H6743, &H76CA)
End Function

Private Function SC_KeHuChuShiBaoZhengJin() As String
    ' 客户初始保证金
    SC_KeHuChuShiBaoZhengJin = U(&H5BA2, &H6237, &H521D, &H59CB, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_ChuShiBaoZhengJin() As String
    ' 初始保证金
    SC_ChuShiBaoZhengJin = U(&H521D, &H59CB, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_BaoZhengJin() As String
    ' 保证金
    SC_BaoZhengJin = U(&H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_KeHuBaoZhengJin() As String
    ' 客户保证金
    SC_KeHuBaoZhengJin = U(&H5BA2, &H6237, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_KeYong1() As String
    ' 可用1
    SC_KeYong1 = U(&H53EF, &H7528) & "1"
End Function

Private Function SC_KeYong() As String
    ' 可用
    SC_KeYong = U(&H53EF, &H7528)
End Function


Sub MapOTCClientFunds()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Set wsSrc = ThisWorkbook.Sheets("OTC客戶資金")
    Set wsDst = ThisWorkbook.Sheets("OTC")

    ' ==================================================
    ' 第一步：動態查找 OTC客戶資金 表頭列號
    ' ==================================================
    Dim headerRow As Long
    headerRow = 1

    Dim colAcct       As Long
    Dim colCurrGroup  As Long
    Dim colCurrCode   As Long
    Dim colDailyEq    As Long
    Dim colInitMargin As Long
    Dim colAvail      As Long

    ' 預先組出簡體表頭（ChrW），繁體維持字面量
    Dim hKeHuZhangHao As String, hZiJinZhangHao As String, hKeHuZiJinZhangHao As String
    Dim hBiZhongZuBie As String, hBiZhongZu As String, hBiZhongZuHao As String
    Dim hBiZhongHao As String, hBiZhong As String
    Dim hDangRiQuanYi As String, hJinRiQuanYi As String, hJinQuanYi As String
    Dim hKeHuChuShiBZJ As String, hChuShiBZJ As String, hBZJ As String, hKeHuBZJ As String
    Dim hKeYong1 As String, hKeYong As String

    hKeHuZhangHao = SC_KeHuZhangHao()
    hZiJinZhangHao = SC_ZiJinZhangHao()
    hKeHuZiJinZhangHao = SC_KeHuZiJinZhangHao()
    hBiZhongZuBie = SC_BiZhongZuBie()
    hBiZhongZu = SC_BiZhongZu()
    hBiZhongZuHao = SC_BiZhongZuHao()
    hBiZhongHao = SC_BiZhongHao()
    hBiZhong = SC_BiZhong()
    hDangRiQuanYi = SC_DangRiQuanYi()
    hJinRiQuanYi = SC_JinRiQuanYi()
    hJinQuanYi = SC_JinQuanYi()
    hKeHuChuShiBZJ = SC_KeHuChuShiBaoZhengJin()
    hChuShiBZJ = SC_ChuShiBaoZhengJin()
    hBZJ = SC_BaoZhengJin()
    hKeHuBZJ = SC_KeHuBaoZhengJin()
    hKeYong1 = SC_KeYong1()
    hKeYong = SC_KeYong()

    Dim srcTotalCols As Long
    srcTotalCols = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column

    Dim h As Long
    For h = 1 To srcTotalCols
        Dim hVal As String
        hVal = Trim(wsSrc.Cells(headerRow, h).Value)

        Select Case hVal
            Case hKeHuZhangHao, "客戶賬號", _
                 hZiJinZhangHao, "資金賬號", _
                 hKeHuZiJinZhangHao, "客戶資金賬號"
                colAcct = h

            Case hBiZhongZuBie, "幣種組別", _
                 hBiZhongZu, "幣種組", _
                 hBiZhongZuHao, "幣種組號"
                colCurrGroup = h

            Case hBiZhongHao, "幣種號", _
                 hBiZhong, "幣種"
                colCurrCode = h

            Case hDangRiQuanYi, "當日權益", _
                 hJinRiQuanYi, "今日權益", _
                 hJinQuanYi, "今權益"
                colDailyEq = h

            Case hKeHuChuShiBZJ, "客戶初始保證金", _
                 hChuShiBZJ, "初始保證金", _
                 hBZJ, "保證金", _
                 hKeHuBZJ, "客戶保證金"
                colInitMargin = h

            Case hKeYong1, hKeYong
                colAvail = h
        End Select
    Next h

    ' 找不到必要欄位則報錯退出
    If colAcct = 0 Or colCurrGroup = 0 Or colCurrCode = 0 _
       Or colDailyEq = 0 Or colInitMargin = 0 Then

        Dim missingCols As String
        If colAcct = 0 Then missingCols = missingCols & "  - 客戶賬號／資金賬號" & vbCrLf
        If colCurrGroup = 0 Then missingCols = missingCols & "  - 幣種組別／幣種組" & vbCrLf
        If colCurrCode = 0 Then missingCols = missingCols & "  - 幣種號／幣種" & vbCrLf
        If colDailyEq = 0 Then missingCols = missingCols & "  - 當日權益／今日權益" & vbCrLf
        If colInitMargin = 0 Then missingCols = missingCols & "  - 客戶初始保證金／保證金" & vbCrLf

        MsgBox "OTC客戶資金 表中找不到以下欄位：" & vbCrLf & vbCrLf & _
               missingCols & vbCrLf & "請確認表頭名稱後重試。", vbCritical, "欄位識別失敗"
        Exit Sub
    End If

    ' ==================================================
    ' 第二步：確保 _MarginLog 暫存 Sheet 存在
    ' ==================================================
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("_MarginLog")
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "_MarginLog"
        wsLog.Columns(1).NumberFormat = "@"
        wsLog.Columns(4).NumberFormat = "@"
        wsLog.Cells(1, 1).Value = "Date"
        wsLog.Cells(1, 2).Value = "AcctNum"
        wsLog.Cells(1, 3).Value = "MarginBase"
        wsLog.Cells(1, 4).Value = "Time"
        wsLog.Visible = xlSheetVeryHidden
    Else
        wsLog.Columns(1).NumberFormat = "@"
        wsLog.Columns(4).NumberFormat = "@"
    End If

    ' 統一日期與時間格式
    Dim today   As String
    Dim nowTime As String
    today = Format(Date, "yyyy-mm-dd")
    nowTime = Format(Now, "hh:mm:ss")

    ' ==================================================
    ' 第三步：寫入 Previous Margin 標題（B13）
    ' ==================================================
    If wsDst.Cells(13, 2).Value = "" Then
        wsDst.Cells(13, 2).Value = "Previous Margin"
        wsDst.Cells(13, 2).Font.Bold = True
    End If

    ' ==================================================
    ' 第四步：主循環，遍歷 OTC Sheet 各賬號 Block
    ' ==================================================
    Dim srcLR As Long
    srcLR = wsSrc.Cells(wsSrc.Rows.Count, colAcct).End(xlUp).Row

    Dim blockCols As Variant
    blockCols = Array(1, 6, 11, 16, 21)   ' A/F/K/P/U 列（賬號列）

    Dim b As Integer
    For b = 0 To UBound(blockCols)

        Dim acctCol As Long
        Dim valCol  As Long
        acctCol = blockCols(b)
        valCol = acctCol + 2    ' C/H/M/R（值列）

        Dim dstAcct As String
        dstAcct = Trim(wsDst.Cells(3, acctCol).Value)
        If dstAcct = "" Then GoTo NextBlock

        ' 提取純數字賬號
        Dim acctNum As String
        Dim spPos As Long
        spPos = InStr(dstAcct, " ")
        If spPos > 0 Then
            acctNum = Trim(Left(dstAcct, spPos - 1))
        Else
            acctNum = dstAcct
        End If

        ' 初始化數據變量
        Dim valEquityBase As Double
        Dim valMarginBase As Double
        Dim valMarginUSD  As Double
        Dim valAvailCNH   As Double
        Dim valAvailUSD   As Double
        Dim hasCNH        As Boolean
        valEquityBase = 0
        valMarginBase = 0
        valMarginUSD = 0
        valAvailCNH = 0
        valAvailUSD = 0
        hasCNH = False

        ' --------------------------------------------------
        ' 在 OTC客戶資金 中查找對應賬號數據
        ' --------------------------------------------------
        Dim j As Long
        For j = headerRow + 1 To srcLR

            Dim srcAcctRaw As String
            srcAcctRaw = Trim(wsSrc.Cells(j, colAcct).Value)

            Dim srcAcctNum As String
            If InStr(srcAcctRaw, "(") > 0 Then
                srcAcctNum = Trim(Left(srcAcctRaw, InStr(srcAcctRaw, "(") - 1))
            Else
                srcAcctNum = srcAcctRaw
            End If

            If srcAcctNum = acctNum Then

                Dim currGroup As String
                Dim currCode  As String
                currGroup = Trim(wsSrc.Cells(j, colCurrGroup).Value)
                currCode = Trim(wsSrc.Cells(j, colCurrCode).Value)

                Dim isOther As Boolean
                Dim isBase  As Boolean
                Dim isUSD   As Boolean
                Dim isCNH   As Boolean
                isOther = (UCase(currGroup) = "OTHER")
                isBase = (UCase(currCode) = "BASE" Or _
                           UCase(currCode) = "OTHER_BASE" Or _
                           UCase(currCode) = "OTHER BASE")
                isUSD = (UCase(currCode) = "USD" Or _
                           UCase(currCode) = "OTHER_USD" Or _
                           UCase(currCode) = "OTHER USD")
                isCNH = (UCase(currCode) = "CNH" Or _
                           UCase(currCode) = "OTHER_CNH" Or _
                           UCase(currCode) = "OTHER CNH")

                If isOther And isBase Then
                    valEquityBase = wsSrc.Cells(j, colDailyEq).Value
                    valMarginBase = wsSrc.Cells(j, colInitMargin).Value
                End If

                If isOther And isUSD Then
                    valMarginUSD = wsSrc.Cells(j, colInitMargin).Value
                    If colAvail > 0 Then
                        valAvailUSD = wsSrc.Cells(j, colAvail).Value
                    End If
                End If

                If isOther And isCNH Then
                    hasCNH = True
                    If colAvail > 0 Then
                        valAvailCNH = wsSrc.Cells(j, colAvail).Value
                    End If
                End If

            End If
        Next j

        ' --------------------------------------------------
        ' 回填 OTC Sheet
        ' --------------------------------------------------

        ' 1. Equity（行4）
        wsDst.Cells(4, valCol).Value = valEquityBase
        wsDst.Cells(4, valCol).NumberFormat = "#,##0.00"

        ' --------------------------------------------------
        ' 2. Margin（行5）+ Previous Margin（行14）+ 標色邏輯
        '    順序：① 查舊值 → ② 寫 Previous Margin & 時間
        '          → ③ 寫新 Margin → ④ 比較標色 → ⑤ 更新 Log
        ' --------------------------------------------------

        ' ① 查找 _MarginLog 舊記錄
        Dim logLR      As Long
        Dim prevMargin As Double
        Dim prevDate   As String
        Dim prevTime   As String
        Dim logRow     As Long
        Dim foundLog   As Boolean
        logLR = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row
        prevMargin = 0
        prevDate = ""
        prevTime = ""
        logRow = 0
        foundLog = False

        Dim lj As Long
        For lj = 2 To logLR
            If Trim(wsLog.Cells(lj, 2).Value) = acctNum Then
                Dim rawDate As String
                rawDate = Trim(wsLog.Cells(lj, 1).Value)
                On Error Resume Next
                prevDate = Format(CDate(rawDate), "yyyy-mm-dd")
                On Error GoTo 0
                If prevDate = "" Then prevDate = rawDate
                prevMargin = wsLog.Cells(lj, 3).Value
                prevTime = Trim(wsLog.Cells(lj, 4).Value)
                logRow = lj
                foundLog = True
                Exit For
            End If
        Next lj

        ' ② 寫入 Previous Margin（C/H/M/R 14行）& 時間（D/I/N/S 14行）
        Dim timeCol As Long
        timeCol = valCol + 1   ' D/I/N/S 列

        If foundLog And prevDate = today Then
            ' 同天有記錄：顯示上次值及時間
            wsDst.Cells(14, valCol).Value = prevMargin
            wsDst.Cells(14, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(14, timeCol).Value = prevTime
        ElseIf foundLog And prevDate <> today Then
            ' 跨天：顯示上次值，時間加上日期備注
            wsDst.Cells(14, valCol).Value = prevMargin
            wsDst.Cells(14, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(14, timeCol).Value = prevTime & " (" & prevDate & ")"
        Else
            ' 首次運行：無歷史記錄
            wsDst.Cells(14, valCol).Value = "N/A"
            wsDst.Cells(14, timeCol).Value = ""
        End If

        ' ③ 寫入新 Margin 值
        wsDst.Cells(5, valCol).Value = valMarginBase
        wsDst.Cells(5, valCol).NumberFormat = "#,##0.00"

        ' ④ 比較標色（同天才比較）
        If foundLog And prevDate = today Then
            If valMarginBase < prevMargin Then
                ' 本次 < 上次 → 標紅
                wsDst.Cells(5, valCol).Interior.Color = RGB(255, 199, 199)
                wsDst.Cells(5, valCol).Font.Color = RGB(180, 0, 0)
            Else
                ' 本次 >= 上次 → 清除標色
                wsDst.Cells(5, valCol).Interior.ColorIndex = xlNone
                wsDst.Cells(5, valCol).Font.ColorIndex = xlAutomatic
            End If
        Else
            ' 首次或跨天 → 清除標色
            wsDst.Cells(5, valCol).Interior.ColorIndex = xlNone
            wsDst.Cells(5, valCol).Font.ColorIndex = xlAutomatic
        End If

        ' ⑤ 更新 _MarginLog（最後才寫，避免跟自己比）
        If foundLog Then
            wsLog.Cells(logRow, 1).NumberFormat = "@"
            wsLog.Cells(logRow, 1).Value = today
            wsLog.Cells(logRow, 3).Value = valMarginBase
            wsLog.Cells(logRow, 4).NumberFormat = "@"
            wsLog.Cells(logRow, 4).Value = nowTime
        Else
            Dim newRow As Long
            newRow = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row + 1
            wsLog.Cells(newRow, 1).NumberFormat = "@"
            wsLog.Cells(newRow, 1).Value = today
            wsLog.Cells(newRow, 2).Value = acctNum
            wsLog.Cells(newRow, 3).Value = valMarginBase
            wsLog.Cells(newRow, 4).NumberFormat = "@"
            wsLog.Cells(newRow, 4).Value = nowTime
        End If

        ' --------------------------------------------------
        ' 3. Excess（行6）= 公式 + CNH/USD 備注
        ' --------------------------------------------------
        Dim equityAddr As String
        Dim marginAddr As String
        Dim noteCol    As Long
        equityAddr = wsDst.Cells(4, valCol).Address(False, False)
        marginAddr = wsDst.Cells(5, valCol).Address(False, False)
        noteCol = valCol + 2      ' E/J/O/T（跳過 D 的 deposite 欄）

        If hasCNH And colAvail > 0 Then
            Dim cnh_m As String
            Dim usd_m As String
            cnh_m = Format(Int(valAvailCNH / 10000) / 100, "0.00") & "M CNH"
            usd_m = Format(Int(valAvailUSD / 10000) / 100, "0.00") & "M USD"

            wsDst.Cells(6, valCol).Formula = "=" & equityAddr & "-" & marginAddr
            wsDst.Cells(6, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(6, noteCol).Value = "(" & cnh_m & " and " & usd_m & ")"
        Else
            wsDst.Cells(6, valCol).Formula = "=" & equityAddr & "-" & marginAddr
            wsDst.Cells(6, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(6, noteCol).Value = ""
        End If

        ' --------------------------------------------------
        ' 4. 僅 101758000：計算匯率寫入 B1
        ' --------------------------------------------------
        If acctNum = "101758000" Then
            If valMarginBase <> 0 And valMarginUSD <> 0 Then
                wsDst.Cells(1, 2).Value = valMarginBase / valMarginUSD
                wsDst.Cells(1, 2).NumberFormat = "0.0000"
            End If
        End If

NextBlock:
    Next b

    MsgBox "完成！數據已成功對應至 OTC Sheet。", vbInformation, "MapOTCClientFunds"

End Sub


' ==================================================
' 修復工具：清空並修正 _MarginLog 格式
' ==================================================
Sub FixAndClearMarginLog()
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("_MarginLog")
    On Error GoTo 0

    If Not wsLog Is Nothing Then
        wsLog.Cells.Clear
        wsLog.Columns(1).NumberFormat = "@"
        wsLog.Columns(4).NumberFormat = "@"
        wsLog.Cells(1, 1).Value = "Date"
        wsLog.Cells(1, 2).Value = "AcctNum"
        wsLog.Cells(1, 3).Value = "MarginBase"
        wsLog.Cells(1, 4).Value = "Time"
        MsgBox "_MarginLog 已清空並修正格式。" & vbCrLf & _
               "請重新運行 MapOTCClientFunds。", vbInformation
    Else
        MsgBox "_MarginLog 不存在，運行 MapOTCClientFunds 會自動建立。", vbInformation
    End If
End Sub


' ==================================================
' 診斷工具：確認表頭及數據樣本
' ==================================================
Sub DebugOTCHeaders()

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Sheets("OTC客戶資金")

    Dim srcTotalCols As Long
    srcTotalCols = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    Dim headerList As String
    Dim h As Long
    For h = 1 To srcTotalCols
        Dim hVal As String
        hVal = wsSrc.Cells(1, h).Value
        If hVal <> "" Then
            headerList = headerList & "Col " & h & ": [" & hVal & "]" & vbCrLf
        End If
    Next h
    MsgBox "=== 所有表頭 ===" & vbCrLf & headerList, , "表頭檢查"

    Dim colAcct      As Long
    Dim colCurrGroup As Long
    Dim colCurrCode  As Long

    ' 簡體表頭經 ChrW 組出
    Dim hKeHuZhangHao As String, hZiJinZhangHao As String, hKeHuZiJinZhangHao As String
    Dim hBiZhongZuBie As String, hBiZhongZu As String, hBiZhongZuHao As String
    Dim hBiZhongHao As String, hBiZhong As String

    hKeHuZhangHao = SC_KeHuZhangHao()
    hZiJinZhangHao = SC_ZiJinZhangHao()
    hKeHuZiJinZhangHao = SC_KeHuZiJinZhangHao()
    hBiZhongZuBie = SC_BiZhongZuBie()
    hBiZhongZu = SC_BiZhongZu()
    hBiZhongZuHao = SC_BiZhongZuHao()
    hBiZhongHao = SC_BiZhongHao()
    hBiZhong = SC_BiZhong()

    For h = 1 To srcTotalCols
        Select Case Trim(wsSrc.Cells(1, h).Value)
            Case hKeHuZhangHao, "客戶賬號", hZiJinZhangHao, "資金賬號", _
                 hKeHuZiJinZhangHao, "客戶資金賬號"
                colAcct = h
            Case hBiZhongZuBie, "幣種組別", hBiZhongZuHao, "幣種組號", _
                 hBiZhongZu, "幣種組"
                colCurrGroup = h
            Case hBiZhongHao, "幣種號", hBiZhong, "幣種"
                colCurrCode = h
        End Select
    Next h

    If colAcct = 0 Then
        MsgBox "? 找不到賬號欄位！請查看表頭列表確認名稱。", vbCritical
        Exit Sub
    End If

    Dim sampleData As String
    sampleData = "賬號列=" & colAcct & "  組別列=" & colCurrGroup & _
                 "  幣種列=" & colCurrCode & vbCrLf & vbCrLf

    Dim r As Long
    For r = 2 To Application.Min(6, wsSrc.Cells(wsSrc.Rows.Count, colAcct).End(xlUp).Row)
        Dim acctVal As String
        Dim grpVal  As String
        Dim codeVal As String
        acctVal = wsSrc.Cells(r, colAcct).Value
        grpVal = IIf(colCurrGroup > 0, wsSrc.Cells(r, colCurrGroup).Value, "N/A")
        codeVal = IIf(colCurrCode > 0, wsSrc.Cells(r, colCurrCode).Value, "N/A")
        sampleData = sampleData _
            & "Row" & r & ": 賬號=[" & acctVal & "](len=" & Len(acctVal) & ")" & vbCrLf _
            & "       組別=[" & grpVal & "]  幣種=[" & codeVal & "]" & vbCrLf
    Next r
    MsgBox "=== 前5筆數據樣本 ===" & vbCrLf & sampleData, , "數據檢查"

    Dim wsDst As Worksheet
    Set wsDst = ThisWorkbook.Sheets("OTC")
    Dim blockCols As Variant
    blockCols = Array(1, 6, 11, 16)
    Dim dstInfo As String
    Dim bi As Integer
    For bi = 0 To UBound(blockCols)
        Dim ac As Long
        ac = blockCols(bi)
        dstInfo = dstInfo & "Col " & ac & " (Row3): [" & wsDst.Cells(3, ac).Value & "]" & vbCrLf
    Next bi
    MsgBox "=== OTC Sheet 賬號行 ===" & vbCrLf & dstInfo, , "目標賬號"

End Sub


' ==================================================
' 診斷工具：確認 _MarginLog 內容及日期比對
' ==================================================
Sub DebugMarginLog()

    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("_MarginLog")
    On Error GoTo 0

    If wsLog Is Nothing Then
        MsgBox "? _MarginLog Sheet 不存在！", vbCritical
        Exit Sub
    End If

    Dim logLR As Long
    logLR = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row

    If logLR < 2 Then
        MsgBox "? _MarginLog 是空的。", vbCritical
        Exit Sub
    End If

    Dim logContent As String
    Dim lj As Long
    For lj = 1 To logLR
        logContent = logContent _
            & "Row" & lj & ": " _
            & "[Date=" & wsLog.Cells(lj, 1).Value & "] " _
            & "[Acct=" & wsLog.Cells(lj, 2).Value & "] " _
            & "[Margin=" & wsLog.Cells(lj, 3).Value & "] " _
            & "[Time=" & wsLog.Cells(lj, 4).Value & "]" & vbCrLf
    Next lj
    MsgBox "=== _MarginLog 內容 ===" & vbCrLf & logContent, , "Log 記錄"

    Dim today As String
    today = Format(Date, "yyyy-mm-dd")
    Dim logDate As String
    logDate = Trim(wsLog.Cells(2, 1).Value)
    On Error Resume Next
    logDate = Format(CDate(logDate), "yyyy-mm-dd")
    On Error GoTo 0

    MsgBox "今天日期：[" & today & "]" & vbCrLf & _
           "Log 日期：[" & logDate & "]" & vbCrLf & vbCrLf & _
           IIf(today = logDate, "? 日期相同，比較邏輯可正常觸發", _
           "? 日期不同！請運行 FixAndClearMarginLog"), , "日期比對"

End Sub
