Option Explicit On
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Drawing.Imaging

Public Class frmAfterHours
    Dim startTime As DateTime
    Dim endTime As TimeSpan
    Dim balance(4, 19) As Single '20日多空淨額之暫存陣列
    Dim days(20) As Aday '20日之日期
    Dim dataCount As Short '此變數表示資料數量(此數值會隨著大額交易資訊之有無而變動)
    Dim lblTitle(2) As Label '表格標題陣列，程式開啟時動態增加
    Dim chkSS(5) As CheckBox '圖表元件顯示與否之核取方塊，程式開啟時動態增加
    '定義六組資列序列(三大法人、合計、加權指數、十大)
    Dim ss(5) As Series
    Structure WebPosition '定義網頁座標資料結構，用以表示資料所在的表格、列、欄
        Dim table As Integer
        Dim row As Integer
        Dim column As Integer
        Sub New(tableNO As Integer, rowNO As Integer, colNo As Integer)
            table = tableNO
            row = rowNO
            column = colNo
        End Sub
    End Structure
    Structure Aday '定義日期資料結構，設定建構子以便自動初始化，方便分別取得年月日
        Dim full As Date
        Dim yy As String
        Dim mm As String
        Dim dd As String
        Sub New(adate As Date) 'Aday的建構子，分別從日期中取出年月日並補0成為兩位數(月、日)
            full = adate
            yy = adate.Year
            mm = Format(adate.Month, "00")
            dd = Format(adate.Day, "00")
        End Sub
        Sub setDate(aDate As Date) '設定一個新日期
            full = aDate
            yy = aDate.Year
            mm = Format(aDate.Month, "00")
            dd = Format(aDate.Day, "00")
        End Sub
    End Structure
    Structure StockIndex '定義股價資料結構，用以表示開盤、收盤、最高、最低
        Dim open As Single
        Dim highest As Single
        Dim lowest As Single
        Dim close As Single
        Sub New(open As Single, highest As Single, lowest As Single, close As Single)
            Me.open = open
            Me.highest = highest
            Me.lowest = lowest
            Me.close = close
        End Sub
    End Structure
    Dim WTXindex(19) As StockIndex '20天的加權指數資訊
    Dim WTXflag As Boolean = False '用以判斷是否已經取得過加權指數資訊，避免重複取得浪費資源

    '搜尋HTML字串，取得第tableNO個表格，第rowNO列，第columnNO行之資訊，並轉為數值
    Private Function SearchHTML(html As String, pos As WebPosition) As Single
        '將HTML全部轉為小寫，避免誤判
        html = LCase(html)
        '利用自訂函式ExtractText()搜尋夾在<table和</table>之間的字串
        html = ExtractText(html, "<table", "</table>", pos.table)
        '利用相同原理搜尋<tr......</tr>
        html = ExtractText(html, "<tr", "</tr>", pos.row)
        '利用相同原理搜尋<td......</td>
        html = ExtractText(html, "<td", "</td>", pos.column)
        '利用Regular Expression去除HTML標籤
        'Regular Expression: ? 接在*或+等運算子後面，代表使該運算子的次數越少越好(亦即可以只重複一次就不要重複兩次)
        '                    . 代表除了/n之外的所有單字元
        html = System.Text.RegularExpressions.Regex.Replace(html, "<(.|/n)*?>", String.Empty)
        '去掉千分位符號,
        html = html.Replace(",", String.Empty)
        '回傳數值
        Return Val(html)
    End Function

    '找到第Nth個被startText和endText包圍的字串
    '例︰找到HTML之中的第三個table
    Private Function ExtractText(originText As String, startText As String, endText As String, Nth As Integer)
        Dim posStart As Integer
        Dim posEnd As Integer
        '重複Nth次Instr()函式，找到startText的位置
        posStart = 0
        For I = 1 To Nth
            posStart = InStr(posStart + 1, originText, startText)
        Next
        '找到該段文字endText的位置
        posEnd = InStr(posStart, originText, endText) + Len(endText) - 1
        '回傳該段文字，長度為posEnd - posStart + 1
        Return Mid(originText, posStart, posEnd - posStart + 1)
    End Function

    '搜尋HTML字串，取得第tableNO個表格，第rowNO列，第columnNO行之資訊，並轉為數值
    '運用Regular Expression搜尋HTML字串
    '***因為這個方法會忽略包含其他table的table(或者包含其他tr的tr)，亦即只搜尋最深層的標籤，***
    '***所以搜尋表格時要特別注意有幾個表格，通常表格順序如同網頁瀏覽器上所看到的順序，不會有***
    '***沒看到的表格***
    Private Function ExtractHTMLValue(html As String, pos As WebPosition) As String
        html = LCase(html)
        Dim blocks As MatchCollection
        '***因為找到的片段之索引(index)是從0開始，所以所有的順序都要減去1(因為一般由1開始算)
        '找出 <table...</table> 片段
        blocks = Regex.Matches(html, "<table\b[^>]*>(?:(?>[^<]+)|<(?!table\b[^>]*>))*?</table>")
        html = blocks.Item(pos.table - 1).Value
        '找出 <tr>...</tr> 片段
        blocks = Regex.Matches(html, "<tr\b[^>]*>(?:(?>[^<]+)|<(?!tr\b[^>]*>))*?</tr>")
        html = blocks.Item(pos.row - 1).Value
        '找出 <td>...</td> 片段
        blocks = Regex.Matches(html, "<td\b[^>]*>(?:(?>[^<]+)|<(?!td\b[^>]*>))*?</td>")
        html = blocks.Item(pos.column - 1).Value
        '利用Regular Expression去除HTML標籤
        'Regular Expression: ? 接在*或+等運算子後面，代表使該運算子的次數越少越好(亦即可以只重複一次就不要重複兩次)
        '                    . 代表除了/n之外的所有單字元
        html = System.Text.RegularExpressions.Regex.Replace(html, "<(.|/n)*?>", String.Empty)
        '去除括號之後字串(大額交易人資訊會有括號代表百分比，此為不需要之資訊)
        If InStr(html, "(") > 0 Then  '有找到括號再去除，避免將字串吃乾抹淨
            html = Mid(html, 1, InStr(html, "(") - 1)
        End If

        '回傳數值
        html = html.Replace(",", String.Empty) '去除千分位符號
        html = html.Replace(" ", String.Empty) '去除空白
        Return html
    End Function

    Private Function getSourceCode(url As String, encode As Encoding) As String
        '利用WebClient開啟網址，將原始碼讀入資料串流(Stream)中，再利用StreamReader設定編碼並讀出字串
        Dim cc As New WebClient
        Dim dataS As Stream = cc.OpenRead(url)
        Dim dataR As New StreamReader(dataS, encode)
        Return dataR.ReadToEnd
    End Function

    '功能類似於GenerateInfo()，用於表格和圖表中已有資料，還要加入其他資料時使用
    '例︰表格和圖表中已有大台的相關數據，此時想統計大台加上小台的合計，就直接呼叫
    'GenerateInfoAddition()填入小台相關係數即可，此函式會自動將計算出的小台數據和已存在
    '的大台數據相加
    '*****本函式不會產生新表格和圖表，且相關變數都直接沿用GenerateInfo()之數值，因此務必
    '*****先呼叫GenerateInfo()後再使用，並確認兩者格式皆相同(若格式不同，至少GenerateInfo()產生
    '*****之表格和圖表必須是最大範圍(例︰GenerateInfo()有計入大額交易，但GenerateInfoAddition()沒有
    'coeff︰本次計算出的數據相加之前要乘上的係數，例︰原本已存在資料是TX，本次計算出的數據是MTX，則coeff = 0.25
    '，代表TX + 0.25 * MTX
    Private Sub GenerateInfoAddition(urlPrefix As String, contractID As String, isOp As Boolean, _
                             isLarge As Boolean, startPos As WebPosition, _
                             columnShift As Integer, _
                             Optional urlLargePrefix As String = "", _
                             Optional contractIDLarge As String = "", _
                             Optional largeTableShift As Integer = 0, _
                             Optional largeRowShift As Integer = 0, _
                             Optional largeColumnShift As Integer = 0, _
                             Optional coeff As Single = 1.0)
        Dim I, J As Short '迴圈變數
        Dim url As String = "" '目前欲取得資料之網址
        Dim html As String = "" '目前欲取得資料之網頁原始碼
        Dim html2 As String = "" '目前欲取得資料之網頁原始碼#2
        '取得昨日和今日資訊
        For I = 0 To 1
            '設定try...catch區塊以防止異常終止
            Try
                '取得該日期之網頁原始碼
                url = urlPrefix & "?syear=" & days(I).yy & "&smonth=" & days(I).mm & "&sday=" & days(I).dd & "&COMMODITY_ID=" & contractID
                html = getSourceCode(url, Encoding.UTF8)
                '***取得非大額交易人之資訊
                '取得自營商之買賣資訊，多空淨額直接計算，降低取得網頁資訊次數
                dgvDisplay.Rows(1).Cells(4 - 3 * I).Value += CInt(ExtractHTMLValue(html, startPos)) * coeff
                dgvDisplay.Rows(1).Cells(5 - 3 * I).Value += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row, startPos.column + 2))) * coeff
                dgvDisplay.Rows(1).Cells(6 - 3 * I).Value = dgvDisplay.Rows(1).Cells(4 - 3 * I).Value - dgvDisplay.Rows(1).Cells(5 - 3 * I).Value
                '利用迴圈取得投信&外資之買賣資訊(因為有合併儲存格的問題，無法將自營商資訊取得納入迴圈)
                For J = 1 To 2
                    dgvDisplay.Rows(1 + J).Cells(4 - 3 * I).Value += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + J, startPos.column + columnShift))) * coeff
                    dgvDisplay.Rows(1 + J).Cells(5 - 3 * I).Value += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + J, startPos.column + columnShift + 2))) * coeff
                    dgvDisplay.Rows(1 + J).Cells(6 - 3 * I).Value = dgvDisplay.Rows(1 + J).Cells(4 - 3 * I).Value - dgvDisplay.Rows(1 + J).Cells(5 - 3 * I).Value
                Next
                '***取得大額交易人資訊***
                If isLarge Then
                    url = urlLargePrefix & "?choose_yy=" & days(I).yy & "&choose_mm=" & days(I).mm & "&choose_dd=" & days(I).dd & "&choose_item=" & contractIDLarge
                    html2 = getSourceCode(url, Encoding.UTF8)
                    For J = 0 To 1
                        dgvDisplay.Rows(5 + J).Cells(4 - 3 * I).Value += CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + J * 2))) * coeff
                        dgvDisplay.Rows(5 + J).Cells(5 - 3 * I).Value += CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + 4 + J * 2))) * coeff
                        dgvDisplay.Rows(5 + J).Cells(6 - 3 * I).Value = dgvDisplay.Rows(5 + J).Cells(4 - 3 * I).Value - dgvDisplay.Rows(5 + J).Cells(5 - 3 * I).Value
                    Next
                End If
            Catch ex As Exception
                If I > 0 Then
                    MsgBox("取得盤後資訊錯誤。" & vbCrLf & ex.Message)
                    Exit For
                Else
                    MsgBox("取得盤後資訊錯誤，請確認本日盤後資訊已經公佈。" & vbCrLf & ex.Message)
                End If
            End Try
            '計算三大法人總合買賣資訊
            For J = 4 To 6
                dgvDisplay.Rows(4).Cells(J - 3 * I).Value = dgvDisplay.Rows(1).Cells(J - 3 * I).Value + dgvDisplay.Rows(2).Cells(J - 3 * I).Value + dgvDisplay.Rows(3).Cells(J - 3 * I).Value
            Next
            '將近兩日之多空淨額存入陣列，以減少之後繪圖的運算時間(少讀取兩次網頁)
            For J = 0 To 3
                balance(J, I) = dgvDisplay.Rows(J + 1).Cells(6 - 3 * I).Value
            Next
            If isLarge Then
                balance(4, I) = dgvDisplay.Rows(6).Cells(6 - 3 * I).Value
            End If
        Next
        '***計算變化量***
        '利用迴圈計算(將前面兩欄資訊相減，即今日減去昨日)
        For I = 1 To dgvDisplay.RowCount - 1
            For J = 7 To 9
                dgvDisplay.Rows(I).Cells(J).Value = dgvDisplay.Rows(I).Cells(J - 3).Value - dgvDisplay.Rows(I).Cells(J - 6).Value
            Next
        Next
        '*****開始取得20日多空淨額*****
        For I = 1 To dataCount '清空資料序列，避免點數過多
            ss(I).Points.Clear()
        Next
        For I = 2 To 19 '只需要取得2日~19日前的多空淨額(因為近兩日的資訊已經存入陣列)
            Try
                '*****開始取得資料*****
                '取得原始碼
                url = urlPrefix & "?syear=" & days(I).yy & "&smonth=" & days(I).mm & "&sday=" & days(I).dd & "&COMMODITY_ID=" & contractID
                html = getSourceCode(url, Encoding.UTF8)
                '*****資料取得完畢****
                '以自訂函式ExtractHTMLValue()解析原始碼，取得多空餘額及日期，並存入陣列
                balance(0, I) += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row, startPos.column + 4))) * coeff
                balance(1, I) += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + 1, startPos.column + columnShift + 4))) * coeff
                balance(2, I) += CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + 2, startPos.column + columnShift + 4))) * coeff
                balance(3, I) = balance(0, I) + balance(1, I) + balance(2, I)
                If isLarge Then
                    url = urlLargePrefix & "?choose_yy=" & days(I).yy & "&choose_mm=" & days(I).mm & "&choose_dd=" & days(I).dd & "&choose_item=" & contractIDLarge
                    html2 = getSourceCode(url, Encoding.UTF8)
                    balance(4, I) += (CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + 2))) - CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + 1, startPos.row + largeRowShift, startPos.column + largeColumnShift + 6)))) * coeff
                End If
            Catch ex As Exception
                MsgBox("取得多空淨額失敗。" & vbCrLf & ex.Message)
                Exit For
            End Try
        Next
        For I = 0 To 19 '將陣列資料指派給資料序列
            '將陣列資料指派給資料序列，以Format()設定日期格式再轉為字串
            For J = 1 To dataCount
                ss(J).Points.AddXY(Format(days(I).full, "yyyy/MM/dd").ToString, balance(J - 1, I))
            Next
        Next
    End Sub

    '確認該日期是否休假，若開市傳回False，若休假傳回True
    '假使傳入日期為今日，但目前時間早於15:00直接傳回False(因為此時不會有當日的開盤資料，無論該日期
    '是否是休假日)
    Private Function Dayoff(inputDate As Date) As Boolean
        Dim holiday(34) As Date
        Dim extraWorkday(3) As Date
        '定義國定假日(固定日期的假日，如元旦、228、清明，將直接以月份日期判斷)
        holiday(0) = #2/7/2013#
        holiday(1) = #2/8/2013#
        holiday(2) = #2/11/2013#
        holiday(3) = #2/12/2013#
        holiday(4) = #2/13/2013#
        holiday(5) = #2/14/2013#
        holiday(6) = #2/15/2013#
        holiday(7) = #4/5/2013#
        holiday(8) = #6/12/2013#
        holiday(9) = #9/19/2013#
        holiday(10) = #9/20/2013#
        holiday(11) = #1/19/2012#
        holiday(12) = #1/20/2012#
        holiday(13) = #1/23/2012#
        holiday(14) = #1/24/2012#
        holiday(15) = #1/25/2012#
        holiday(16) = #1/26/2012#
        holiday(17) = #1/27/2012#
        holiday(18) = #2/27/2012#
        holiday(19) = #12/31/2012#
        holiday(20) = #4/4/2014#
        holiday(21) = #6/2/2014#
        holiday(22) = #9/8/2014#
        holiday(23) = #1/2/2015#
        holiday(24) = #2/16/2015#
        holiday(25) = #2/17/2015#
        holiday(26) = #2/18/2015#
        holiday(27) = #2/19/2015#
        holiday(28) = #2/20/2015#
        holiday(29) = #2/23/2015#
        holiday(30) = #2/27/2015#
        holiday(31) = #4/3/2015#
        holiday(32) = #4/6/2015#
        holiday(33) = #9/28/2015#
        holiday(34) = #10/9/2015#
        '定義補上班日
        extraWorkday(0) = #9/14/2013#
        extraWorkday(1) = #12/22/2012#
        extraWorkday(2) = #2/23/2013#
        extraWorkday(3) = #12/27/2014#
        If inputDate.Equals(Now.Date) And Now.Hour < 15 Then '傳入日期為今日，但不到15點傳回False
            Return True
        End If
        '以月份日期判斷假日(固定日期的假日)
        If inputDate.Month = 1 And inputDate.Day = 1 Then
            Return True
        End If
        If inputDate.Month = 2 And inputDate.Day = 28 Then
            Return True
        End If
        If inputDate.Month = 4 And inputDate.Day = 4 Then
            Return True
        End If
        If inputDate.Month = 5 And inputDate.Day = 1 Then
            Return True
        End If
        If inputDate.Month = 10 And inputDate.Day = 10 Then
            Return True
        End If
        '以Weekday函式判斷是否是週末
        If Weekday(inputDate) = 1 Or Weekday(inputDate) = 7 Then
            '假如是週末，確認是否是補上班日，是的話傳回False，否則傳回True
            For Each dd In extraWorkday
                If inputDate.Equals(dd) Then
                    Return False
                End If
            Next
            Return True
        Else
            '假如不是週末，確認是否是國定假日，是的話傳回True，否則傳回False
            For Each dd In holiday
                If inputDate.Equals(dd) Then
                    Return True
                End If
            Next
            Return False
        End If
    End Function

    Private Sub dgvDisplay_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvDisplay.CellFormatting
        '以ColumnIndex屬性判斷行數，將每行分別上色
        Select Case e.ColumnIndex
            Case 0
                e.CellStyle.BackColor = Color.FromArgb(253, 228, 195)
                e.CellStyle.SelectionBackColor = Color.FromArgb(253, 228, 195)
            Case 1 To 3
                e.CellStyle.BackColor = Color.FromArgb(249, 254, 180)
                e.CellStyle.SelectionBackColor = Color.FromArgb(249, 254, 180)
            Case 4 To 6
                e.CellStyle.BackColor = Color.FromArgb(218, 253, 254)
                e.CellStyle.SelectionBackColor = Color.FromArgb(218, 253, 254)
            Case 7 To 9
                e.CellStyle.BackColor = Color.FromArgb(250, 220, 221)
                e.CellStyle.SelectionBackColor = Color.FromArgb(250, 220, 221)
        End Select
    End Sub

    '產生選定期權資訊的表格及圖表
    'isOp︰代表是否為選擇權(標題字串會顯示買方賣方)
    'isLarge︰表示是否統計大額交易人
    'startPos︰代表起始座標(自營商買方/多方)
    '若起始座標(自營商買方/多方)為(X,Y)則︰(表格皆相同，亦即相同Z座標)
    '自營商賣方/空方︰(X+2,Y)
    '      多空淨額︰(X+4,Y)
    '投信買方/空方︰(X-2, Y+1)
    '    賣方/空方︰(X,Y+1)
    '    多空淨額︰(X+2,Y+1)
    '外資買方/空方︰(X-2, Y+2)
    '    賣方/空方︰(X,Y+2)
    '    多空淨額︰(X+2,Y+2)
    'columnShift︰當資料第一列包含合併儲存格時，第二列之後的偏移值，示意圖如下︰
    '==========================================
    '        |           | cell3   | cell4
    '  cell1 |  cell2    |______________________
    '        |           | cell5   | cell6
    '==========================================
    '上述之shift值為-2，因為cell1和cell2皆是兩列合併而成
    '所以會導致︰cell3是第三個td標籤，但相對位置的cell5則會是第一個，亦即第二列之後的行數都要-2(往右為正數，往左為負數)
    'title︰圖表標題
    'urlPrefix︰代表要獲取資訊的網址(不含參數)
    'contractID︰代表要獲取的合約代碼(由網址原始碼手動查詢可得)
    'urlLargePrefix︰代表要獲取大額交易人資訊的網址(不含參數)
    'contractIDLarge︰代表要獲取的大額交易人合約代碼(由網址原始碼手動查詢可得)
    'largeTableShift︰代表大額交易人表格偏移
    'largeRowShift︰代表大額交易人表格之列偏移，說明如下︰假使三大法人之表格，自營商資訊在第四列；但大額交易人
    '的五大交易人資訊在第五列，則largeRowShift值為1，代表讀取大額交易人資訊時要多往下1行(往下為正數，往下為負數)
    'largeColumnShift︰大額交易人表格之行偏移，說明如下︰假使三大法人之表格，自營商之多方(買方)資訊在第十行；但大額交易人之五大交易人
    '多方(買方)資訊在第三行，則largeColumnShift值為-5，代表讀取大額交易人資訊時要多往左移動5行(往右為正數，往左為負數)
    '本函式僅適用於台灣期貨交易所提供的三大法人資訊
    '大額交易人資訊因為表格格式以及網頁變數完全不同，因此直接設定新的參數
    Private Sub GenerateInfo(title As String, urlPrefix As String, contractID As String, isOp As Boolean, _
                             isLarge As Boolean, startPos As WebPosition, _
                             columnShift As Integer, _
                             Optional urlLargePrefix As String = "", _
                             Optional contractIDLarge As String = "", _
                            Optional largeTableShift As Integer = 0, _
                             Optional largeRowShift As Integer = 0, _
                             Optional largeColumnShift As Integer = 0)
        Dim I, J As Short '迴圈變數
        Dim url As String = "" '目前欲取得資料之網址
        Dim html As String = "" '目前欲取得資料之網頁原始碼
        Dim html2 As String = "" '目前欲取得資料之網頁原始碼#2
        Dim mmTemp As String '目前加權指數表的月份
        '初始化變數(資料數量)
        If isLarge Then
            dataCount = 5
        Else
            dataCount = 4
        End If
        '今日日期
        days(0) = New Aday(Now.Date) '以今日日期初始化最近日期(陣列中第一個元素)
        '找出近20日之開市日期，並填入days陣列
        For I = 0 To 19
            '若該日期為休市日，則往前移動一天，直到不是休市日為止
            Do While Dayoff(days(I).full)
                days(I).setDate(days(I).full.AddDays(-1))
            Loop
            '日期往前移動一日
            days(I + 1).setDate(days(I).full.AddDays(-1))
        Next
        '取得近20日之加權指數資訊，因20日之日期跨度不會超過一個月，所以直接以第一天的月份
        '和最後一天的月份取得網頁資訊，再將兩個網頁原始碼合併成單一字串以利搜尋
        '搜尋時比對首行日期，若與days陣列中日期符合，則將後續行的值填入WTXindex陣列
        '若沒有符合日期，則會跳出錯誤訊息(列數超出索引)，以try...catch區塊避免異常終止
        If WTXflag = False Then
            Try
                '以首日的月份取得加權指數網頁資訊原始碼
                url = "http://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date=" & days(0).full.ToString("yyyyMMdd")
                html = getSourceCode(url, Encoding.GetEncoding("big5"))
                mmTemp = days(0).mm
                Dim twDate As New System.Globalization.TaiwanCalendar '定義民國紀年
                For I = 0 To 19
                    J = 3
                    '若day(i)的月份與目前加權指數表的月份(mmTemp)相同，直接從html中搜尋；若不是則重新下載加權指數表，並以mmTemp
                    '紀錄新加權指數表的月份
                    If days(I).mm <> mmTemp Then
                        url = "http://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date=" & days(I).full.ToString("yyyyMMdd")
                        html = getSourceCode(url, Encoding.GetEncoding("big5"))
                        mmTemp = days(I).mm
                    End If
                    Do
                        '以網頁中表5列j行1的資料(日期)作比對，若與day(j)相同則將表5列j行2~5的資料填
                        '入WTXindex(j)，需注意day(j)中日期是西元紀年，需轉為民國紀年(台灣證券交易所以民國紀年表示)
                        '因為前兩列為標題，所以j由3開始
                        If Date.Parse(ExtractHTMLValue(html, New WebPosition(1, J, 1))).Equals(Date.Parse(twDate.GetYear(days(I).full).ToString + days(I).full.ToString("/MM/dd"))) Then
                            WTXindex(I).open = Single.Parse(ExtractHTMLValue(html, New WebPosition(1, J, 2)))
                            WTXindex(I).highest = Single.Parse(ExtractHTMLValue(html, New WebPosition(1, J, 3)))
                            WTXindex(I).lowest = Single.Parse(ExtractHTMLValue(html, New WebPosition(1, J, 4)))
                            WTXindex(I).close = Single.Parse(ExtractHTMLValue(html, New WebPosition(1, J, 5)))
                            Exit Do
                        End If
                        J = J + 1
                    Loop
                Next
                '成功取得20天的加權指數資訊(因為沒有跳出至catch區塊)，將WTXflag設為true，避免資源浪費
                WTXflag = True
            Catch ex As Exception
                MsgBox("無法取得加權指數。" & vbCrLf & ex.Message)
            End Try
        End If
        '*****利用DataGridView作為表格，以輸出買賣資訊*****
        '先清除所有表格資訊，再新增表格欄位(5列10行)並設定標題
        Dim tagTitles(3) As String '標題字串，因為期貨是多方空方；但選擇權是買方賣方
        '決定標題字串，若為選擇權資訊，顯示買方賣方；期貨資訊則顯示多方空方
        If isOp Then
            tagTitles(0) = "買方"
            tagTitles(1) = "賣方"
            tagTitles(2) = "買賣差額"
        Else
            tagTitles(0) = "多方"
            tagTitles(1) = "空方"
            tagTitles(2) = "多空淨額"
        End If
        dgvDisplay.Columns.Clear()
        dgvDisplay.Columns.Add("", "")
        dgvDisplay.Rows.Add("")
        dgvDisplay.Rows.Add("自營商")
        dgvDisplay.Rows.Add("投信")
        dgvDisplay.Rows.Add("外資")
        dgvDisplay.Rows.Add("三大法人合計")
        For I = 0 To 2
            dgvDisplay.Columns.Add("bull" & I, tagTitles(0))
            dgvDisplay.Columns.Add("bear" & I, tagTitles(1))
            dgvDisplay.Columns.Add("balance" & I, tagTitles(2))
        Next
        For I = 0 To 2
            dgvDisplay.Rows(0).Cells(1 + 3 * I).Value = tagTitles(0)
            dgvDisplay.Rows(0).Cells(2 + 3 * I).Value = tagTitles(1)
            dgvDisplay.Rows(0).Cells(3 + 3 * I).Value = tagTitles(2)
        Next
        '設定表格格式
        For Each col In dgvDisplay.Columns
            col.Resizable = DataGridViewTriState.False
            col.ReadOnly = True
            col.SortMode = DataGridViewColumnSortMode.NotSortable
            col.Width = 80
        Next
        '設定表格位置及長寬
        dgvDisplay.Left = 10
        dgvDisplay.Top = 130
        dgvDisplay.Width = 80 * 10 + 1
        dgvDisplay.Height = 24 * 5 + 1
        '假如該表格含有大額交易人資訊，再插入新欄位並調整高度
        If isLarge Then
            '添加兩列(五大、十大)，並重新設定表格高度
            dgvDisplay.Rows.Add("五大交易人")
            dgvDisplay.Rows.Add("十大交易人")
            dgvDisplay.Height = 24 * 7 + 1
        End If
        '*****表格初始化完畢*****
        '取得昨日和今日資訊
        For I = 0 To 1
            '設定該日期之標題，並設定該標題出現的位置(對應的表格上方)
            lblTitle(I).Left = dgvDisplay.Left + 320 - 240 * I - 1
            lblTitle(I).Top = dgvDisplay.Top - 25
            lblTitle(I).Width = 240
            If I = 0 Then
                lblTitle(I).Text = "今日(" & Format(days(I).full, "yyyy年MM月dd日") & ")"
            Else
                lblTitle(I).Text = "昨日(" & Format(days(I).full, "yyyy年MM月dd日") & ")"
            End If
            '設定try...catch區塊以防止異常終止
            Try
                '取得該日期之網頁原始碼
                url = urlPrefix & "?syear=" & days(I).yy & "&smonth=" & days(I).mm & "&sday=" & days(I).dd & "&COMMODITY_ID=" & contractID
                html = getSourceCode(url, Encoding.UTF8)
                '***取得非大額交易人之資訊
                '取得自營商之買賣資訊，多空淨額直接計算，降低取得網頁資訊次數
                dgvDisplay.Rows(1).Cells(4 - 3 * I).Value = CInt(ExtractHTMLValue(html, startPos))
                dgvDisplay.Rows(1).Cells(5 - 3 * I).Value = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row, startPos.column + 2)))
                dgvDisplay.Rows(1).Cells(6 - 3 * I).Value = dgvDisplay.Rows(1).Cells(4 - 3 * I).Value - dgvDisplay.Rows(1).Cells(5 - 3 * I).Value
                '利用迴圈取得投信&外資之買賣資訊(因為有合併儲存格的問題，無法將自營商資訊取得納入迴圈)
                For J = 1 To 2
                    dgvDisplay.Rows(1 + J).Cells(4 - 3 * I).Value = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + J, startPos.column + columnShift)))
                    dgvDisplay.Rows(1 + J).Cells(5 - 3 * I).Value = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + J, startPos.column + columnShift + 2)))
                    dgvDisplay.Rows(1 + J).Cells(6 - 3 * I).Value = dgvDisplay.Rows(1 + J).Cells(4 - 3 * I).Value - dgvDisplay.Rows(1 + J).Cells(5 - 3 * I).Value
                Next
                '***取得大額交易人資訊***
                If isLarge Then
                    url = urlLargePrefix & "?choose_yy=" & days(I).yy & "&choose_mm=" & days(I).mm & "&choose_dd=" & days(I).dd & "&choose_item=" & contractIDLarge
                    html2 = getSourceCode(url, Encoding.UTF8)
                    For J = 0 To 1
                        dgvDisplay.Rows(5 + J).Cells(4 - 3 * I).Value = CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + J * 2)))
                        dgvDisplay.Rows(5 + J).Cells(5 - 3 * I).Value = CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + 4 + J * 2)))
                        dgvDisplay.Rows(5 + J).Cells(6 - 3 * I).Value = dgvDisplay.Rows(5 + J).Cells(4 - 3 * I).Value - dgvDisplay.Rows(5 + J).Cells(5 - 3 * I).Value
                    Next
                End If
            Catch ex As Exception
                If I > 0 Then
                    MsgBox("取得盤後資訊錯誤。" & vbCrLf & ex.Message)
                    Exit For
                Else
                    MsgBox("取得盤後資訊錯誤，請確認本日盤後資訊已經公佈。" & vbCrLf & ex.Message)
                End If
            End Try
            '計算三大法人總合買賣資訊
            For J = 4 To 6
                dgvDisplay.Rows(4).Cells(J - 3 * I).Value = dgvDisplay.Rows(1).Cells(J - 3 * I).Value + dgvDisplay.Rows(2).Cells(J - 3 * I).Value + dgvDisplay.Rows(3).Cells(J - 3 * I).Value
            Next
            '將近兩日之多空淨額存入陣列，以減少之後繪圖的運算時間(少讀取兩次網頁)
            For J = 0 To 3
                balance(J, I) = dgvDisplay.Rows(J + 1).Cells(6 - 3 * I).Value
            Next
            If isLarge Then
                balance(4, I) = dgvDisplay.Rows(6).Cells(6 - 3 * I).Value
            End If
        Next
        '***計算變化量***
        '設定變化量之標題，並使之出現於表格上方
        With lblTitle(2)
            .Left = dgvDisplay.Left + 560 - 1
            .Top = dgvDisplay.Top - 25
            .Width = 240
            .Text = "變化量"
        End With
        '利用迴圈計算(將前面兩欄資訊相減，即今日減去昨日)
        For I = 1 To dgvDisplay.RowCount - 1
            For J = 7 To 9
                dgvDisplay.Rows(I).Cells(J).Value = dgvDisplay.Rows(I).Cells(J - 3).Value - dgvDisplay.Rows(I).Cells(J - 6).Value
            Next
        Next
        '*****產生20天的多空淨額圖表*****
        '利用MSChart繪製圖表，設定Chart位置及大小，使之緊貼於表格(DataGridView)下緣且寬度相同
        chtDisplay.Left = dgvDisplay.Left
        chtDisplay.Top = dgvDisplay.Top + dgvDisplay.Height
        chtDisplay.Width = 800
        chtDisplay.Height = 300
        '***初始化MSChart***
        '清空內容，避免內容重複
        chtDisplay.Series.Clear()
        chtDisplay.ChartAreas.Clear()
        chtDisplay.Legends.Clear()
        chtDisplay.Titles.Clear()
        '重置所有顯示項目核取方塊
        For I = 0 To 5
            chkSS(I).Checked = False
            chkSS(I).Enabled = False
        Next
        '***設定標題(Title)、繪圖區域(ChartArea)、資列序列(Series)、圖例(Legend)***
        '添加標題，使之出現在繪圖區域上方，利用Font Class設定字型
        chtDisplay.Titles.Add(New Title(title, Docking.Top, New Font("微軟正黑體", 16), Color.Black))
        '添加繪圖區域main並初始化
        chtDisplay.ChartAreas.Add("main")
        With chtDisplay.ChartAreas("main")
            'X軸座標間距1
            .AxisX.Interval = 1
            'X軸座標傾斜角度-70度
            .AxisX.LabelStyle.Angle = -70
            'X軸資訊反轉(因為要將最靠近的日期列於右邊)
            .AxisX.IsReversed = True
            '不使用3D介面
            .Area3DStyle.Enable3D = False
            '取消所有格線
            .AxisX.MajorGrid.Enabled = False
            .AxisY.MajorGrid.Enabled = False
            .AxisY2.MajorGrid.Enabled = False
            .AxisX.MajorTickMark.Enabled = False
            .AxisY.MajorTickMark.Enabled = False
            .AxisY2.MajorTickMark.Enabled = False
            '設定座標軸標籤
            .AxisY.Title = "未平倉淨額"
            .AxisY.TextOrientation = TextOrientation.Stacked
            .AxisY.TitleFont = New Font("微軟正黑體", 12)
            .AxisY2.Title = "加權指數"
            .AxisY2.TextOrientation = TextOrientation.Stacked
            .AxisY2.TitleFont = New Font("微軟正黑體", 12)
            '加權指數軸區間定為50
            .AxisY2.Interval = 50
        End With
        '以迴圈設定資列序列，並以New產生新的Series物件
        For I = 0 To dataCount
            ss(I) = New Series
            'X軸座標格式為字串
            ss(I).XValueType = ChartValueType.String
            'Y軸座標格式為32位元整數
            ss(I).YValueType = ChartValueType.Int32
            'ss(1).CustomProperties = "PointWidth=0.5"
            '利用格式化字串設定滑鼠停駐時顯示的資訊(日期 序列名稱: 數值)
            ss(I).ToolTip = "#VALX #SERIESNAME:  #VALY"
        Next
        '分別設定各組資料序列之名稱及圖表類型
        '0: 加權指數, 1: 自營商, 2:投信, 3: 外資, 4: 三大法人合計, 5:十大交易人(選配)
        ss(0).Name = "加權指數"
        ss(0).ChartType = SeriesChartType.Line
        ss(0).Color = Color.FromArgb(200, 22, 175)
        ss(0).BorderWidth = 2
        ss(0).MarkerStyle = MarkerStyle.Circle
        ss(0).MarkerSize = 7
        ss(1).Name = "自營商"
        ss(1).ChartType = SeriesChartType.Column
        ss(2).Name = "投信"
        ss(2).ChartType = SeriesChartType.Column
        ss(3).Name = "外資"
        ss(3).ChartType = SeriesChartType.Column
        ss(4).Name = "三大法人合計"
        ss(4).ChartType = SeriesChartType.Line
        ss(4).Color = Color.FromArgb(3, 219, 29)
        ss(4).BorderWidth = 2
        ss(4).MarkerStyle = MarkerStyle.Circle
        ss(4).MarkerSize = 7
        If isLarge Then '若含有大額交易人才需設定第六組資列序列
            ss(5).Name = "十大交易人"
            ss(5).ChartType = SeriesChartType.Line
            ss(5).Color = Color.FromArgb(250, 111, 56)
            ss(5).BorderWidth = 2
            ss(5).MarkerStyle = MarkerStyle.Circle
            ss(5).MarkerSize = 7
        End If
        '加權指數之Y軸以副座標軸顯示，但X軸依然使用同樣的主座標軸(日期)
        '並重新設定加權指數之Y軸single
        ss(0).XAxisType = AxisType.Primary
        ss(0).YAxisType = AxisType.Secondary
        ss(0).YValueType = ChartValueType.Single
        '添加圖例，並設定顯示於繪圖區下方
        chtDisplay.Legends.Add("")
        chtDisplay.Legends(0).Docking = Docking.Bottom
        '設定加權指數軸之預設範圍
        chtDisplay.ChartAreas(0).AxisY2.Maximum = CInt(Format(WTXindex(0).close * 1.01 / 1000, ".0") * 1000)
        chtDisplay.ChartAreas(0).AxisY2.Minimum = CInt(Format(WTXindex(0).close * 1.01 / 1000, ".0") * 1000)
        For I = 2 To 19 '只需要取得2日~19日前的多空淨額(因為近兩日的資訊已經存入陣列)
            Try
                '*****開始取得資料*****
                '取得原始碼
                url = urlPrefix & "?syear=" & days(I).yy & "&smonth=" & days(I).mm & "&sday=" & days(I).dd & "&COMMODITY_ID=" & contractID
                html = getSourceCode(url, Encoding.UTF8)
                '*****資料取得完畢****
                '以自訂函式ExtractHTMLValue()解析原始碼，取得多空餘額及日期，並存入陣列
                balance(0, I) = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row, startPos.column + 4)))
                balance(1, I) = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + 1, startPos.column + columnShift + 4)))
                balance(2, I) = CInt(ExtractHTMLValue(html, New WebPosition(startPos.table, startPos.row + 2, startPos.column + columnShift + 4)))
                balance(3, I) = balance(0, I) + balance(1, I) + balance(2, I)
                If isLarge Then
                    url = urlLargePrefix & "?choose_yy=" & days(I).yy & "&choose_mm=" & days(I).mm & "&choose_dd=" & days(I).dd & "&choose_item=" & contractIDLarge
                    html2 = getSourceCode(url, Encoding.UTF8)
                    balance(4, I) = CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + largeTableShift, startPos.row + largeRowShift, startPos.column + largeColumnShift + 2))) - CInt(ExtractHTMLValue(html2, New WebPosition(startPos.table + 1, startPos.row + largeRowShift, startPos.column + largeColumnShift + 6)))
                End If
            Catch ex As Exception
                MsgBox("取得多空淨額失敗。" & vbCrLf & ex.Message)
                Exit For
            End Try
        Next
        For I = 0 To 19 '將陣列資料指派給資料序列
            '將加權指數資訊指派給資列序列，使用同樣的X軸(日期)
            ss(0).Points.AddXY(Format(days(I).full, "yyyy/MM/dd").ToString, Format(WTXindex(I).close, ".00"))
            '將陣列資料指派給資料序列，以Format()設定日期格式再轉為字串
            For J = 1 To dataCount
                ss(J).Points.AddXY(Format(days(I).full, "yyyy/MM/dd").ToString, balance(J - 1, I))
            Next
            '重新依照目前數據調整副座標軸
            If chtDisplay.ChartAreas(0).AxisY2.Maximum < WTXindex(I).close Then
                chtDisplay.ChartAreas(0).AxisY2.Maximum = CInt(Format(WTXindex(I).close * 1.01 / 1000, ".0") * 1000)
            End If
            If chtDisplay.ChartAreas(0).AxisY2.Minimum > WTXindex(I).close Then
                chtDisplay.ChartAreas(0).AxisY2.Minimum = CInt(Format(WTXindex(I).close * 0.99 / 1000, ".0") * 1000)
            End If
        Next
    End Sub

    Private Sub frmAfterHours_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        '初始化標題Label陣列
        For I = 0 To 2
            lblTitle(I) = New Label
            With lblTitle(I)
                .Font = New Font("微軟正黑體", 12)
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleCenter
                .Visible = True
            End With
            '加入表單內
            Me.Controls.Add(lblTitle(I))
        Next
        '初始化核取方塊Checkbox陣列
        For I = 0 To 5
            chkSS(I) = New CheckBox
            With chkSS(I)
                .Font = New Font("微軟正黑體", 11)
                .Top = 26
                .AutoSize = True
                .Enabled = False
                Select Case I
                    Case 0
                        .Left = 6
                        .Text = "加權指數"
                    Case 1
                        .Left = 99
                        .Text = "自營商"
                    Case 2
                        .Left = 177
                        .Text = "投信"
                    Case 3
                        .Left = 240
                        .Text = "外資"
                    Case 4
                        .Left = 303
                        .Text = "三大法人合計"
                    Case 5
                        .Left = 426
                        .Text = "十大交易人"
                End Select
            End With
            '加入群組方塊內
            Me.gupChartItems.Controls.Add(chkSS(I))
            '定義核取方塊事件
            AddHandler chkSS(I).CheckedChanged, AddressOf ChangeChartItems
        Next
        '初始化下拉式清單
        cboType.SelectedIndex = 0
    End Sub

    Private Sub ChangeChartItems(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If chtDisplay.Visible = True Then '圖表顯示時才處理其中元件的顯示與否
            '以核取方塊陣列的TabIndex判斷選中哪個資料序列
            '再以Contain()判斷目前是否顯示該資料序列，若有顯示將之隱藏，否則顯示
            If chtDisplay.Series.Contains(ss(sender.TabIndex)) Then
                chtDisplay.Series.RemoveAt(chtDisplay.Series.IndexOf(ss(sender.TabIndex)))
            Else
                chtDisplay.Series.Add(ss(sender.TabIndex))
            End If
        End If

    End Sub

    Private Sub btnSavePic_Click(sender As System.Object, e As System.EventArgs) Handles btnSavePic.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        Try
            saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
            saveFileDialog1.FilterIndex = 2
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then '使用者按下確認之後紀錄檔名
                GetWindowInfo(Me.Handle)
                '建立一個Bitmap作為存檔目標
                Dim Screenshot As Bitmap = New Bitmap(ClientX, ClientY - 50, PixelFormat.Format32bppArgb)
                Dim picOutput As Graphics = Graphics.FromImage(Screenshot) '建立儲存影像的Graphic
                Dim picSource As Graphics = Graphics.FromHdc(GetDC(Me.Handle)) '建立獲取來源影像的Graphic
                '以Bitbit將來源影像轉存到目標影像
                BitBlt(picOutput.GetHdc(), 0, 0, ClientX, ClientY - 50, picSource.GetHdc(), 0, 50, CopyPixelOperation.SourceCopy) '把來源影像複製到儲存影像中
                '釋放hdc
                picSource.ReleaseHdc()
                picOutput.ReleaseHdc()
                '存為圖片
                Screenshot.Save(saveFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnShow_Click(sender As System.Object, e As System.EventArgs) Handles btnShow.Click
        Dim urlPrefix As String
        Dim urlLargePrefix As String
        MsgBox("開始取得盤後資訊，這可能會耗費數分鐘，請耐心等候。")
        '將所有元件隱藏
        lblNote.Visible = False
        dgvDisplay.Visible = False
        chtDisplay.Visible = False
        For I = 0 To 2
            lblTitle(I).Visible = False
        Next
        startTime = Now '設定日期
        lblTimes.Text = "開始取得" & cboType.SelectedText & "盤後資訊並建立報表..."
        Application.DoEvents()
        Select Case cboType.SelectedIndex
            Case 0 '大台
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_3_tbl.asp"
                GenerateInfo("大台期貨（單位︰口）", urlPrefix, "TXF", False, False, New WebPosition(1, 4, 10), -2)
                lblNote.Text = ""
            Case 1 '小台
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_3_tbl.asp"
                GenerateInfo("小台期貨（單位︰口）", urlPrefix, "MXF", False, False, New WebPosition(1, 4, 10), -2)
                lblNote.Text = ""
            Case 2 '大小台合計(五大十大全)
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_3_tbl.asp"
                urlLargePrefix = "http://www.taifex.com.tw/chinese/3/7_8.asp"
                GenerateInfo("大小台期貨合計（單位︰口）", urlPrefix, "TXF", False, True, New WebPosition(1, 4, 10), -2, urlLargePrefix, "TX", 1, 1, -8)
                GenerateInfoAddition(urlPrefix, "MXF", False, False, New WebPosition(1, 4, 10), -2, , , , , 0.25)
                lblNote.Text = "※五大、十大交易人列出全部月份契約" & vbCrLf & "※已將小型台指期貨依契約規模折算後併計"
            Case 3 'OP合計(口)
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_4_tbl.asp"
                GenerateInfo("選擇權口數合計（單位︰口）", urlPrefix, "", False, False, New WebPosition(1, 22, 9), -1)
                lblNote.Text = ""
            Case 4 'OP合計(契約金額)
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_4_tbl.asp"
                GenerateInfo("選擇權契約金額合計（單位︰千元）", urlPrefix, "", False, False, New WebPosition(1, 22, 10), -1)
                lblNote.Text = ""
            Case 5 'CALL合計(五大十大全)
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_5_tbl.asp"
                urlLargePrefix = "http://www.taifex.com.tw/chinese/3/7_9.asp"
                GenerateInfo("選擇權買權合計（單位︰口）", urlPrefix, "TXO", True, True, New WebPosition(1, 4, 11), -3, urlLargePrefix, "TXO", 1, 2, -9)
                lblNote.Text = "※五大、十大交易人列出全部月份契約" & vbCrLf & "※僅列出台指選擇權"
            Case 6 'PUT合計(五大十大全)
                urlPrefix = "http://www.taifex.com.tw/chinese/3/7_12_5_tbl.asp"
                urlLargePrefix = "http://www.taifex.com.tw/chinese/3/7_9.asp"
                GenerateInfo("選擇權賣權合計（單位︰口）", urlPrefix, "TXO", True, True, New WebPosition(1, 7, 9), -1, urlLargePrefix, "TXO", 1, 2, -7)
                lblNote.Text = "※五大、十大交易人列出全部月份契約" & vbCrLf & "※僅列出台指選擇權"

        End Select
        '調整foot note標籤位置
        lblNote.Top = chtDisplay.Top + chtDisplay.Height + 10
        '將資列序列加入Chart並將所有顯示項目核取方塊都打勾
        For I = 0 To dataCount
            chtDisplay.Series.Add(ss(I))
            chkSS(I).Checked = True
            chkSS(I).Enabled = True
        Next
        '顯示標題標籤、表格及圖表(計算期間不顯示)
        lblNote.Visible = True
        dgvDisplay.Visible = True
        chtDisplay.Visible = True
        For I = 0 To 2
            lblTitle(I).Visible = True
        Next
        '紀錄時間
        endTime = Now.Subtract(startTime)
        lblTimes.Text = "計算時間︰" & endTime.Seconds.ToString & "." & endTime.Milliseconds.ToString & "秒"
    End Sub

End Class
