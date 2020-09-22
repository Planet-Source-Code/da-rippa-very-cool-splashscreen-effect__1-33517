Attribute VB_Name = "mdlMain"
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    'SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub



Public Function GetDesktop(FRM As Form)
    Dim HW As Long
    Dim HA As Long
    Dim iLeft As Integer
    Dim iTop As Integer
    Dim iWidth As Integer
    Dim iHeight As Integer
    FRM.AutoRedraw = True
    FRM.Show
    FRM.Hide
    
    DoEvents
    HA = GetDC(GetDesktopWindow())
    iLeft = FRM.Left / Screen.TwipsPerPixelX
    iTop = FRM.Top / Screen.TwipsPerPixelY
    iWidth = FRM.ScaleWidth
    iHeight = FRM.ScaleHeight
    Call BitBlt(FRM.hdc, 0, 0, iWidth, iHeight, HA, iLeft, iTop, vbSrcCopy)
    FRM.Picture = FRM.Image
    
    FRM.Show

End Function

Public Function Gerade(Number) As Boolean 'Function to see if number is dividable by 2
If Round(Number / 2, 0) = Number / 2 Then
    Gerade = True
Else
    Gerade = False
End If
End Function

Public Sub Pause(Delay)
Dim StartTime
    StartTime = GetTickCount
    Do
    Loop Until StartTime + Delay < GetTickCount
End Sub
