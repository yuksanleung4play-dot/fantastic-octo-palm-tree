Attribute VB_Name = "MapOTCClientFunds"
' ==================================================
' MapOTCClientFunds - full version v6
' Features:
'   1. Dynamically detect OTC client funds headers
'   2. Map Other Base equity/init margin to OTC sheet by account
'   3. Account 101758000: write FX rate to B1
'   4. Accounts with CNH: show CNH/USD available note beside Excess
'   5. Excess formula = Equity - Margin
'   6. Same-day rerun: highlight Margin red if lower than previous
'   7. C14/H14/M14/R14 show previous Margin
'   8. D14/I14/N14/S14 show previous import time
'
' Encoding (v6):
'   Big5 VBA editors corrupt Chinese string literals (SC GBK headers
'   and even TC text depending on paste/codepage). ALL Chinese strings
'   are built with ChrW() Unicode codepoints via U() / SC_* / TC_* / Msg_*.
'   Runtime code has ZERO Chinese string literals.
' ==================================================

' Build a string from Unicode codepoints (bypass editor codepage)
Private Function U(ParamArray cps() As Variant) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = LBound(cps) To UBound(cps)
        s = s & ChrW(cps(i))
    Next i
    U = s
End Function

Private Function SheetOTCClientFunds() As String
    ' OTC ke-hu zi-jin (TC sheet name)
    SheetOTCClientFunds = U(&H004F, &H0054, &H0043, &H5BA2, &H6236, &H8CC7, &H91D1)
End Function

Private Function SC_KeHuZhangHao() As String
    ' SC: kehu zhanghao
    SC_KeHuZhangHao = U(&H5BA2, &H6237, &H8D26, &H53F7)
End Function

Private Function SC_ZiJinZhangHao() As String
    ' SC: zijin zhanghao
    SC_ZiJinZhangHao = U(&H8D44, &H91D1, &H8D26, &H53F7)
End Function

Private Function SC_KeHuZiJinZhangHao() As String
    ' SC: kehu zijin zhanghao
    SC_KeHuZiJinZhangHao = U(&H5BA2, &H6237, &H8D44, &H91D1, &H8D26, &H53F7)
End Function

Private Function SC_BiZhongZuBie() As String
    ' SC: bizhong zubie
    SC_BiZhongZuBie = U(&H5E01, &H79CD, &H7EC4, &H522B)
End Function

Private Function SC_BiZhongZu() As String
    ' SC: bizhong zu
    SC_BiZhongZu = U(&H5E01, &H79CD, &H7EC4)
End Function

Private Function SC_BiZhongZuHao() As String
    ' SC: bizhong zuhao
    SC_BiZhongZuHao = U(&H5E01, &H79CD, &H7EC4, &H53F7)
End Function

Private Function SC_BiZhongHao() As String
    ' SC: bizhong hao
    SC_BiZhongHao = U(&H5E01, &H79CD, &H53F7)
End Function

Private Function SC_BiZhong() As String
    ' SC: bizhong
    SC_BiZhong = U(&H5E01, &H79CD)
End Function

Private Function SC_DangRiQuanYi() As String
    ' SC: dangri quanyi
    SC_DangRiQuanYi = U(&H5F53, &H65E5, &H6743, &H76CA)
End Function

Private Function SC_JinRiQuanYi() As String
    ' SC: jinri quanyi
    SC_JinRiQuanYi = U(&H4ECA, &H65E5, &H6743, &H76CA)
End Function

Private Function SC_JinQuanYi() As String
    ' SC: jin quanyi
    SC_JinQuanYi = U(&H4ECA, &H6743, &H76CA)
End Function

Private Function SC_KeHuChuShiBaoZhengJin() As String
    ' SC: kehu chushi baozhengjin
    SC_KeHuChuShiBaoZhengJin = U(&H5BA2, &H6237, &H521D, &H59CB, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_ChuShiBaoZhengJin() As String
    ' SC: chushi baozhengjin
    SC_ChuShiBaoZhengJin = U(&H521D, &H59CB, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_BaoZhengJin() As String
    ' SC: baozhengjin
    SC_BaoZhengJin = U(&H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_KeHuBaoZhengJin() As String
    ' SC: kehu baozhengjin
    SC_KeHuBaoZhengJin = U(&H5BA2, &H6237, &H4FDD, &H8BC1, &H91D1)
End Function

Private Function SC_KeYong1() As String
    ' SC: keyong1
    SC_KeYong1 = U(&H53EF, &H7528, &H0031)
End Function

Private Function SC_KeYong() As String
    ' SC: keyong
    SC_KeYong = U(&H53EF, &H7528)
End Function

Private Function TC_KeHuZhangHao() As String
    ' TC: kehu zhanghao
    TC_KeHuZhangHao = U(&H5BA2, &H6236, &H8CEC, &H865F)
End Function

Private Function TC_ZiJinZhangHao() As String
    ' TC: zijin zhanghao
    TC_ZiJinZhangHao = U(&H8CC7, &H91D1, &H8CEC, &H865F)
End Function

Private Function TC_KeHuZiJinZhangHao() As String
    ' TC: kehu zijin zhanghao
    TC_KeHuZiJinZhangHao = U(&H5BA2, &H6236, &H8CC7, &H91D1, &H8CEC, &H865F)
End Function

Private Function TC_BiZhongZuBie() As String
    ' TC: bizhong zubie
    TC_BiZhongZuBie = U(&H5E63, &H7A2E, &H7D44, &H5225)
End Function

Private Function TC_BiZhongZu() As String
    ' TC: bizhong zu
    TC_BiZhongZu = U(&H5E63, &H7A2E, &H7D44)
End Function

Private Function TC_BiZhongZuHao() As String
    ' TC: bizhong zuhao
    TC_BiZhongZuHao = U(&H5E63, &H7A2E, &H7D44, &H865F)
End Function

Private Function TC_BiZhongHao() As String
    ' TC: bizhong hao
    TC_BiZhongHao = U(&H5E63, &H7A2E, &H865F)
End Function

Private Function TC_BiZhong() As String
    ' TC: bizhong
    TC_BiZhong = U(&H5E63, &H7A2E)
End Function

Private Function TC_DangRiQuanYi() As String
    ' TC: dangri quanyi
    TC_DangRiQuanYi = U(&H7576, &H65E5, &H6B0A, &H76CA)
End Function

Private Function TC_JinRiQuanYi() As String
    ' TC: jinri quanyi
    TC_JinRiQuanYi = U(&H4ECA, &H65E5, &H6B0A, &H76CA)
End Function

Private Function TC_JinQuanYi() As String
    ' TC: jin quanyi
    TC_JinQuanYi = U(&H4ECA, &H6B0A, &H76CA)
End Function

Private Function TC_KeHuChuShiBaoZhengJin() As String
    ' TC: kehu chushi baozhengjin
    TC_KeHuChuShiBaoZhengJin = U(&H5BA2, &H6236, &H521D, &H59CB, &H4FDD, &H8B49, &H91D1)
End Function

Private Function TC_ChuShiBaoZhengJin() As String
    ' TC: chushi baozhengjin
    TC_ChuShiBaoZhengJin = U(&H521D, &H59CB, &H4FDD, &H8B49, &H91D1)
End Function

Private Function TC_BaoZhengJin() As String
    ' TC: baozhengjin
    TC_BaoZhengJin = U(&H4FDD, &H8B49, &H91D1)
End Function

Private Function TC_KeHuBaoZhengJin() As String
    ' TC: kehu baozhengjin
    TC_KeHuBaoZhengJin = U(&H5BA2, &H6236, &H4FDD, &H8B49, &H91D1)
End Function

Private Function MsgMissingAcct() As String
    ' UI message: MsgMissingAcct
    MsgMissingAcct = U(&H0020, &H0020, &H002D, &H0020, &H5BA2, &H6236, &H8CEC, &H865F, &HFF0F, &H8CC7, &H91D1, &H8CEC, &H865F)
End Function

Private Function MsgMissingCurrGroup() As String
    ' UI message: MsgMissingCurrGroup
    MsgMissingCurrGroup = U(&H0020, &H0020, &H002D, &H0020, &H5E63, &H7A2E, &H7D44, &H5225, &HFF0F, &H5E63, &H7A2E, &H7D44)
End Function

Private Function MsgMissingCurrCode() As String
    ' UI message: MsgMissingCurrCode
    MsgMissingCurrCode = U(&H0020, &H0020, &H002D, &H0020, &H5E63, &H7A2E, &H865F, &HFF0F, &H5E63, &H7A2E)
End Function

Private Function MsgMissingDailyEq() As String
    ' UI message: MsgMissingDailyEq
    MsgMissingDailyEq = U(&H0020, &H0020, &H002D, &H0020, &H7576, &H65E5, &H6B0A, &H76CA, &HFF0F, &H4ECA, &H65E5, &H6B0A, &H76CA)
End Function

Private Function MsgMissingMargin() As String
    ' UI message: MsgMissingMargin
    MsgMissingMargin = U(&H0020, &H0020, &H002D, &H0020, &H5BA2, &H6236, &H521D, &H59CB, &H4FDD, &H8B49, &H91D1, &HFF0F, &H4FDD, &H8B49, &H91D1)
End Function

Private Function MsgMissingColsBody() As String
    ' UI message: MsgMissingColsBody
    MsgMissingColsBody = U(&H004F, &H0054, &H0043, &H5BA2, &H6236, &H8CC7, &H91D1, &H0020, &H8868, &H4E2D, &H627E, &H4E0D, &H5230, &H4EE5, &H4E0B, &H6B04, &H4F4D, &HFF1A)
End Function

Private Function MsgConfirmRetry() As String
    ' UI message: MsgConfirmRetry
    MsgConfirmRetry = U(&H8ACB, &H78BA, &H8A8D, &H8868, &H982D, &H540D, &H7A31, &H5F8C, &H91CD, &H8A66, &H3002)
End Function

Private Function MsgTitleColFail() As String
    ' UI message: MsgTitleColFail
    MsgTitleColFail = U(&H6B04, &H4F4D, &H8B58, &H5225, &H5931, &H6557)
End Function

Private Function MsgDone() As String
    ' UI message: MsgDone
    MsgDone = U(&H5B8C, &H6210, &HFF01, &H6578, &H64DA, &H5DF2, &H6210, &H529F, &H5C0D, &H61C9, &H81F3, &H0020, &H004F, &H0054, &H0043, &H0020, &H0053, &H0068, &H0065, &H0065, &H0074, &H3002)
End Function

Private Function MsgLogCleared() As String
    ' UI message: MsgLogCleared
    MsgLogCleared = U(&H005F, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067, &H0020, &H5DF2, &H6E05, &H7A7A, &H4E26, &H4FEE, &H6B63, &H683C, &H5F0F, &H3002)
End Function

Private Function MsgRerun() As String
    ' UI message: MsgRerun
    MsgRerun = U(&H8ACB, &H91CD, &H65B0, &H904B, &H884C, &H0020, &H004D, &H0061, &H0070, &H004F, &H0054, &H0043, &H0043, &H006C, &H0069, &H0065, &H006E, &H0074, &H0046, &H0075, &H006E, &H0064, &H0073, &H3002)
End Function

Private Function MsgLogMissing() As String
    ' UI message: MsgLogMissing
    MsgLogMissing = U(&H005F, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067, &H0020, &H4E0D, &H5B58, &H5728, &HFF0C, &H904B, &H884C, &H0020, &H004D, &H0061, &H0070, &H004F, &H0054, &H0043, &H0043, &H006C, &H0069, &H0065, &H006E, &H0074, &H0046, &H0075, &H006E, &H0064, &H0073, &H0020, &H6703, &H81EA, &H52D5, &H5EFA, &H7ACB, &H3002)
End Function

Private Function MsgAllHeaders() As String
    ' UI message: MsgAllHeaders
    MsgAllHeaders = U(&H003D, &H003D, &H003D, &H0020, &H6240, &H6709, &H8868, &H982D, &H0020, &H003D, &H003D, &H003D)
End Function

Private Function MsgTitleHeaderCheck() As String
    ' UI message: MsgTitleHeaderCheck
    MsgTitleHeaderCheck = U(&H8868, &H982D, &H6AA2, &H67E5)
End Function

Private Function MsgAcctColNotFound() As String
    ' UI message: MsgAcctColNotFound
    MsgAcctColNotFound = U(&H003F, &H0020, &H627E, &H4E0D, &H5230, &H8CEC, &H865F, &H6B04, &H4F4D, &HFF01, &H8ACB, &H67E5, &H770B, &H8868, &H982D, &H5217, &H8868, &H78BA, &H8A8D, &H540D, &H7A31, &H3002)
End Function

Private Function MsgAcctColEq() As String
    ' UI message: MsgAcctColEq
    MsgAcctColEq = U(&H8CEC, &H865F, &H5217, &H003D)
End Function

Private Function MsgGroupColEq() As String
    ' UI message: MsgGroupColEq
    MsgGroupColEq = U(&H0020, &H0020, &H7D44, &H5225, &H5217, &H003D)
End Function

Private Function MsgCurrColEq() As String
    ' UI message: MsgCurrColEq
    MsgCurrColEq = U(&H0020, &H0020, &H5E63, &H7A2E, &H5217, &H003D)
End Function

Private Function MsgRowAcct() As String
    ' UI message: MsgRowAcct
    MsgRowAcct = U(&H003A, &H0020, &H8CEC, &H865F, &H003D, &H005B)
End Function

Private Function MsgRowGroup() As String
    ' UI message: MsgRowGroup
    MsgRowGroup = U(&H0020, &H0020, &H0020, &H0020, &H0020, &H0020, &H0020, &H7D44, &H5225, &H003D, &H005B)
End Function

Private Function MsgRowCurr() As String
    ' UI message: MsgRowCurr
    MsgRowCurr = U(&H005D, &H0020, &H0020, &H5E63, &H7A2E, &H003D, &H005B)
End Function

Private Function MsgSampleTitle() As String
    ' UI message: MsgSampleTitle
    MsgSampleTitle = U(&H003D, &H003D, &H003D, &H0020, &H524D, &H0035, &H7B46, &H6578, &H64DA, &H6A23, &H672C, &H0020, &H003D, &H003D, &H003D)
End Function

Private Function MsgTitleDataCheck() As String
    ' UI message: MsgTitleDataCheck
    MsgTitleDataCheck = U(&H6578, &H64DA, &H6AA2, &H67E5)
End Function

Private Function MsgOtcAcctRows() As String
    ' UI message: MsgOtcAcctRows
    MsgOtcAcctRows = U(&H003D, &H003D, &H003D, &H0020, &H004F, &H0054, &H0043, &H0020, &H0053, &H0068, &H0065, &H0065, &H0074, &H0020, &H8CEC, &H865F, &H884C, &H0020, &H003D, &H003D, &H003D)
End Function

Private Function MsgTitleTargetAcct() As String
    ' UI message: MsgTitleTargetAcct
    MsgTitleTargetAcct = U(&H76EE, &H6A19, &H8CEC, &H865F)
End Function

Private Function MsgLogSheetMissing() As String
    ' UI message: MsgLogSheetMissing
    MsgLogSheetMissing = U(&H003F, &H0020, &H005F, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067, &H0020, &H0053, &H0068, &H0065, &H0065, &H0074, &H0020, &H4E0D, &H5B58, &H5728, &HFF01)
End Function

Private Function MsgLogEmpty() As String
    ' UI message: MsgLogEmpty
    MsgLogEmpty = U(&H003F, &H0020, &H005F, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067, &H0020, &H662F, &H7A7A, &H7684, &H3002)
End Function

Private Function MsgLogContent() As String
    ' UI message: MsgLogContent
    MsgLogContent = U(&H003D, &H003D, &H003D, &H0020, &H005F, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067, &H0020, &H5167, &H5BB9, &H0020, &H003D, &H003D, &H003D)
End Function

Private Function MsgTitleLogRecord() As String
    ' UI message: MsgTitleLogRecord
    MsgTitleLogRecord = U(&H004C, &H006F, &H0067, &H0020, &H8A18, &H9304)
End Function

Private Function MsgTodayDate() As String
    ' UI message: MsgTodayDate
    MsgTodayDate = U(&H4ECA, &H5929, &H65E5, &H671F, &HFF1A, &H005B)
End Function

Private Function MsgLogDate() As String
    ' UI message: MsgLogDate
    MsgLogDate = U(&H004C, &H006F, &H0067, &H0020, &H65E5, &H671F, &HFF1A, &H005B)
End Function

Private Function MsgDateSame() As String
    ' UI message: MsgDateSame
    MsgDateSame = U(&H003F, &H0020, &H65E5, &H671F, &H76F8, &H540C, &HFF0C, &H6BD4, &H8F03, &H908F, &H8F2F, &H53EF, &H6B63, &H5E38, &H89F8, &H767C)
End Function

Private Function MsgDateDiff() As String
    ' UI message: MsgDateDiff
    MsgDateDiff = U(&H003F, &H0020, &H65E5, &H671F, &H4E0D, &H540C, &HFF01, &H8ACB, &H904B, &H884C, &H0020, &H0046, &H0069, &H0078, &H0041, &H006E, &H0064, &H0043, &H006C, &H0065, &H0061, &H0072, &H004D, &H0061, &H0072, &H0067, &H0069, &H006E, &H004C, &H006F, &H0067)
End Function

Private Function MsgTitleDateCompare() As String
    ' UI message: MsgTitleDateCompare
    MsgTitleDateCompare = U(&H65E5, &H671F, &H6BD4, &H5C0D)
End Function


Sub MapOTCClientFunds()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Set wsSrc = ThisWorkbook.Sheets(SheetOTCClientFunds())
    Set wsDst = ThisWorkbook.Sheets("OTC")

    ' ==================================================
    ' Step 1: find header column indexes
    ' ==================================================
    Dim headerRow As Long
    headerRow = 1

    Dim colAcct       As Long
    Dim colCurrGroup  As Long
    Dim colCurrCode   As Long
    Dim colDailyEq    As Long
    Dim colInitMargin As Long
    Dim colAvail      As Long

    ' Pre-build SC + TC header strings via ChrW
    Dim hKeHuZhangHao As String, hZiJinZhangHao As String, hKeHuZiJinZhangHao As String
    Dim hBiZhongZuBie As String, hBiZhongZu As String, hBiZhongZuHao As String
    Dim hBiZhongHao As String, hBiZhong As String
    Dim hDangRiQuanYi As String, hJinRiQuanYi As String, hJinQuanYi As String
    Dim hKeHuChuShiBZJ As String, hChuShiBZJ As String, hBZJ As String, hKeHuBZJ As String
    Dim hKeYong1 As String, hKeYong As String

    Dim tKeHuZhangHao As String, tZiJinZhangHao As String, tKeHuZiJinZhangHao As String
    Dim tBiZhongZuBie As String, tBiZhongZu As String, tBiZhongZuHao As String
    Dim tBiZhongHao As String, tBiZhong As String
    Dim tDangRiQuanYi As String, tJinRiQuanYi As String, tJinQuanYi As String
    Dim tKeHuChuShiBZJ As String, tChuShiBZJ As String, tBZJ As String, tKeHuBZJ As String

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

    tKeHuZhangHao = TC_KeHuZhangHao()
    tZiJinZhangHao = TC_ZiJinZhangHao()
    tKeHuZiJinZhangHao = TC_KeHuZiJinZhangHao()
    tBiZhongZuBie = TC_BiZhongZuBie()
    tBiZhongZu = TC_BiZhongZu()
    tBiZhongZuHao = TC_BiZhongZuHao()
    tBiZhongHao = TC_BiZhongHao()
    tBiZhong = TC_BiZhong()
    tDangRiQuanYi = TC_DangRiQuanYi()
    tJinRiQuanYi = TC_JinRiQuanYi()
    tJinQuanYi = TC_JinQuanYi()
    tKeHuChuShiBZJ = TC_KeHuChuShiBaoZhengJin()
    tChuShiBZJ = TC_ChuShiBaoZhengJin()
    tBZJ = TC_BaoZhengJin()
    tKeHuBZJ = TC_KeHuBaoZhengJin()

    Dim srcTotalCols As Long
    srcTotalCols = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column

    Dim h As Long
    For h = 1 To srcTotalCols
        Dim hVal As String
        hVal = Trim(wsSrc.Cells(headerRow, h).Value)

        Select Case hVal
            Case hKeHuZhangHao, tKeHuZhangHao, _
                 hZiJinZhangHao, tZiJinZhangHao, _
                 hKeHuZiJinZhangHao, tKeHuZiJinZhangHao
                colAcct = h

            Case hBiZhongZuBie, tBiZhongZuBie, _
                 hBiZhongZu, tBiZhongZu, _
                 hBiZhongZuHao, tBiZhongZuHao
                colCurrGroup = h

            Case hBiZhongHao, tBiZhongHao, _
                 hBiZhong, tBiZhong
                colCurrCode = h

            Case hDangRiQuanYi, tDangRiQuanYi, _
                 hJinRiQuanYi, tJinRiQuanYi, _
                 hJinQuanYi, tJinQuanYi
                colDailyEq = h

            Case hKeHuChuShiBZJ, tKeHuChuShiBZJ, _
                 hChuShiBZJ, tChuShiBZJ, _
                 hBZJ, tBZJ, _
                 hKeHuBZJ, tKeHuBZJ
                colInitMargin = h

            Case hKeYong1, hKeYong
                colAvail = h
        End Select
    Next h

    ' Exit if required columns missing
    If colAcct = 0 Or colCurrGroup = 0 Or colCurrCode = 0 _
       Or colDailyEq = 0 Or colInitMargin = 0 Then

        Dim missingCols As String
        If colAcct = 0 Then missingCols = missingCols & MsgMissingAcct() & vbCrLf
        If colCurrGroup = 0 Then missingCols = missingCols & MsgMissingCurrGroup() & vbCrLf
        If colCurrCode = 0 Then missingCols = missingCols & MsgMissingCurrCode() & vbCrLf
        If colDailyEq = 0 Then missingCols = missingCols & MsgMissingDailyEq() & vbCrLf
        If colInitMargin = 0 Then missingCols = missingCols & MsgMissingMargin() & vbCrLf

        MsgBox MsgMissingColsBody() & vbCrLf & vbCrLf & _
               missingCols & vbCrLf & MsgConfirmRetry(), vbCritical, MsgTitleColFail()
        Exit Sub
    End If

    ' ==================================================
    ' Step 2: ensure _MarginLog sheet exists
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

    Dim today   As String
    Dim nowTime As String
    today = Format(Date, "yyyy-mm-dd")
    nowTime = Format(Now, "hh:mm:ss")

    ' ==================================================
    ' Step 3: Previous Margin label (B13)
    ' ==================================================
    If wsDst.Cells(13, 2).Value = "" Then
        wsDst.Cells(13, 2).Value = "Previous Margin"
        wsDst.Cells(13, 2).Font.Bold = True
    End If

    ' ==================================================
    ' Step 4: main loop over OTC account blocks
    ' ==================================================
    Dim srcLR As Long
    srcLR = wsSrc.Cells(wsSrc.Rows.Count, colAcct).End(xlUp).Row

    Dim blockCols As Variant
    blockCols = Array(1, 6, 11, 16, 21)   ' A/F/K/P/U

    Dim b As Integer
    For b = 0 To UBound(blockCols)

        Dim acctCol As Long
        Dim valCol  As Long
        acctCol = blockCols(b)
        valCol = acctCol + 2    ' C/H/M/R

        Dim dstAcct As String
        dstAcct = Trim(wsDst.Cells(3, acctCol).Value)
        If dstAcct = "" Then GoTo NextBlock

        Dim acctNum As String
        Dim spPos As Long
        spPos = InStr(dstAcct, " ")
        If spPos > 0 Then
            acctNum = Trim(Left(dstAcct, spPos - 1))
        Else
            acctNum = dstAcct
        End If

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

        ' 1. Equity (row 4)
        wsDst.Cells(4, valCol).Value = valEquityBase
        wsDst.Cells(4, valCol).NumberFormat = "#,##0.00"

        ' 2. Margin (row 5) + Previous Margin (row 14) + color
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

        Dim timeCol As Long
        timeCol = valCol + 1   ' D/I/N/S

        If foundLog And prevDate = today Then
            wsDst.Cells(14, valCol).Value = prevMargin
            wsDst.Cells(14, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(14, timeCol).Value = prevTime
        ElseIf foundLog And prevDate <> today Then
            wsDst.Cells(14, valCol).Value = prevMargin
            wsDst.Cells(14, valCol).NumberFormat = "#,##0.00"
            wsDst.Cells(14, timeCol).Value = prevTime & " (" & prevDate & ")"
        Else
            wsDst.Cells(14, valCol).Value = "N/A"
            wsDst.Cells(14, timeCol).Value = ""
        End If

        wsDst.Cells(5, valCol).Value = valMarginBase
        wsDst.Cells(5, valCol).NumberFormat = "#,##0.00"

        If foundLog And prevDate = today Then
            If valMarginBase < prevMargin Then
                wsDst.Cells(5, valCol).Interior.Color = RGB(255, 199, 199)
                wsDst.Cells(5, valCol).Font.Color = RGB(180, 0, 0)
            Else
                wsDst.Cells(5, valCol).Interior.ColorIndex = xlNone
                wsDst.Cells(5, valCol).Font.ColorIndex = xlAutomatic
            End If
        Else
            wsDst.Cells(5, valCol).Interior.ColorIndex = xlNone
            wsDst.Cells(5, valCol).Font.ColorIndex = xlAutomatic
        End If

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

        ' 3. Excess (row 6)
        Dim equityAddr As String
        Dim marginAddr As String
        Dim noteCol    As Long
        equityAddr = wsDst.Cells(4, valCol).Address(False, False)
        marginAddr = wsDst.Cells(5, valCol).Address(False, False)
        noteCol = valCol + 2

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

        ' 4. FX for 101758000 -> B1
        If acctNum = "101758000" Then
            If valMarginBase <> 0 And valMarginUSD <> 0 Then
                wsDst.Cells(1, 2).Value = valMarginBase / valMarginUSD
                wsDst.Cells(1, 2).NumberFormat = "0.0000"
            End If
        End If

NextBlock:
    Next b

    MsgBox MsgDone(), vbInformation, "MapOTCClientFunds"

End Sub


' ==================================================
' Utility: clear and fix _MarginLog format
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
        MsgBox MsgLogCleared() & vbCrLf & MsgRerun(), vbInformation
    Else
        MsgBox MsgLogMissing(), vbInformation
    End If
End Sub


' ==================================================
' Debug: headers and sample rows
' ==================================================
Sub DebugOTCHeaders()

    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Sheets(SheetOTCClientFunds())

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
    MsgBox MsgAllHeaders() & vbCrLf & headerList, , MsgTitleHeaderCheck()

    Dim colAcct      As Long
    Dim colCurrGroup As Long
    Dim colCurrCode  As Long

    Dim hKeHuZhangHao As String, hZiJinZhangHao As String, hKeHuZiJinZhangHao As String
    Dim hBiZhongZuBie As String, hBiZhongZu As String, hBiZhongZuHao As String
    Dim hBiZhongHao As String, hBiZhong As String
    Dim tKeHuZhangHao As String, tZiJinZhangHao As String, tKeHuZiJinZhangHao As String
    Dim tBiZhongZuBie As String, tBiZhongZu As String, tBiZhongZuHao As String
    Dim tBiZhongHao As String, tBiZhong As String

    hKeHuZhangHao = SC_KeHuZhangHao()
    hZiJinZhangHao = SC_ZiJinZhangHao()
    hKeHuZiJinZhangHao = SC_KeHuZiJinZhangHao()
    hBiZhongZuBie = SC_BiZhongZuBie()
    hBiZhongZu = SC_BiZhongZu()
    hBiZhongZuHao = SC_BiZhongZuHao()
    hBiZhongHao = SC_BiZhongHao()
    hBiZhong = SC_BiZhong()

    tKeHuZhangHao = TC_KeHuZhangHao()
    tZiJinZhangHao = TC_ZiJinZhangHao()
    tKeHuZiJinZhangHao = TC_KeHuZiJinZhangHao()
    tBiZhongZuBie = TC_BiZhongZuBie()
    tBiZhongZu = TC_BiZhongZu()
    tBiZhongZuHao = TC_BiZhongZuHao()
    tBiZhongHao = TC_BiZhongHao()
    tBiZhong = TC_BiZhong()

    For h = 1 To srcTotalCols
        Select Case Trim(wsSrc.Cells(1, h).Value)
            Case hKeHuZhangHao, tKeHuZhangHao, hZiJinZhangHao, tZiJinZhangHao, _
                 hKeHuZiJinZhangHao, tKeHuZiJinZhangHao
                colAcct = h
            Case hBiZhongZuBie, tBiZhongZuBie, hBiZhongZuHao, tBiZhongZuHao, _
                 hBiZhongZu, tBiZhongZu
                colCurrGroup = h
            Case hBiZhongHao, tBiZhongHao, hBiZhong, tBiZhong
                colCurrCode = h
        End Select
    Next h

    If colAcct = 0 Then
        MsgBox MsgAcctColNotFound(), vbCritical
        Exit Sub
    End If

    Dim sampleData As String
    sampleData = MsgAcctColEq() & colAcct & MsgGroupColEq() & colCurrGroup & _
                 MsgCurrColEq() & colCurrCode & vbCrLf & vbCrLf

    Dim r As Long
    For r = 2 To Application.Min(6, wsSrc.Cells(wsSrc.Rows.Count, colAcct).End(xlUp).Row)
        Dim acctVal As String
        Dim grpVal  As String
        Dim codeVal As String
        acctVal = wsSrc.Cells(r, colAcct).Value
        grpVal = IIf(colCurrGroup > 0, wsSrc.Cells(r, colCurrGroup).Value, "N/A")
        codeVal = IIf(colCurrCode > 0, wsSrc.Cells(r, colCurrCode).Value, "N/A")
        sampleData = sampleData _
            & "Row" & r & MsgRowAcct() & acctVal & "](len=" & Len(acctVal) & ")" & vbCrLf _
            & MsgRowGroup() & grpVal & MsgRowCurr() & codeVal & "]" & vbCrLf
    Next r
    MsgBox MsgSampleTitle() & vbCrLf & sampleData, , MsgTitleDataCheck()

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
    MsgBox MsgOtcAcctRows() & vbCrLf & dstInfo, , MsgTitleTargetAcct()

End Sub


' ==================================================
' Debug: _MarginLog contents and date compare
' ==================================================
Sub DebugMarginLog()

    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("_MarginLog")
    On Error GoTo 0

    If wsLog Is Nothing Then
        MsgBox MsgLogSheetMissing(), vbCritical
        Exit Sub
    End If

    Dim logLR As Long
    logLR = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row

    If logLR < 2 Then
        MsgBox MsgLogEmpty(), vbCritical
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
    MsgBox MsgLogContent() & vbCrLf & logContent, , MsgTitleLogRecord()

    Dim today As String
    today = Format(Date, "yyyy-mm-dd")
    Dim logDate As String
    logDate = Trim(wsLog.Cells(2, 1).Value)
    On Error Resume Next
    logDate = Format(CDate(logDate), "yyyy-mm-dd")
    On Error GoTo 0

    MsgBox MsgTodayDate() & today & "]" & vbCrLf & _
           MsgLogDate() & logDate & "]" & vbCrLf & vbCrLf & _
           IIf(today = logDate, MsgDateSame(), MsgDateDiff()), , MsgTitleDateCompare()

End Sub
