Module modWindowInfo
    Public x0, y0, x1, y1 As Integer '視窗四角座標
    Public Client_x0, Client_y0 As Integer '視窗左上座標(不含標題邊框)
    Public ClientX, ClientY As Integer '視窗長寬(不含標題邊框)
    Public Border As Integer '視窗邊框寬度
    Public Title As Integer '視窗標題高度
    Public hwnd As Integer '視窗hwnd

    '截圖用API
    Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
    Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
    '取得視窗大小(含邊框，功能表)
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef rectangle As RECT) As Integer
    '取得視窗大小(不含邊框，功能表，左上角為0,0)
    Public Declare Function GetClientRect Lib "user32 " (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
    '取得目前視窗hwnd
    Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Integer

    Structure RECT
        Dim x1 As Integer
        Dim y1 As Integer
        Dim x2 As Integer
        Dim y2 As Integer
    End Structure

    '取得視窗資訊
    Sub GetWindowInfo()
        Dim R As RECT
        Dim RetVal As Integer
        RetVal = GetWindowRect(GetForegroundWindow(), R)
        x0 = R.x1
        x1 = R.x2
        y0 = R.y1
        y1 = R.y2
        RetVal = GetClientRect(GetForegroundWindow(), R)
        ClientX = R.x2
        ClientY = R.y2
        Border = ((x1 - x0) - ClientX) / 2
        Title = (y1 - y0) - ClientY - Border
        Client_x0 = x0 + Border
        Client_y0 = y0 + Title
    End Sub

    '取得特定視窗資訊
    Sub GetWindowInfo(hwnd As Integer)
        Dim R As RECT
        Dim RetVal As Integer
        RetVal = GetWindowRect(hwnd, R)
        x0 = R.x1
        x1 = R.x2
        y0 = R.y1
        y1 = R.y2
        RetVal = GetClientRect(hwnd, R)
        ClientX = R.x2
        ClientY = R.y2
        Border = ((x1 - x0) - ClientX) / 2
        Title = (y1 - y0) - ClientY - Border
        Client_x0 = x0 + Border
        Client_y0 = y0 + Title
    End Sub
End Module
