Attribute VB_Name = "mMain"

Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Enum ApiConstants
    SPI_SCREENSAVERRUNNING = 97
    WS_CHILD = &H40000000
    GWL_STYLE = -16
    GWL_HWNDPARENT = -8
    HWND_TOP = 0
    HWND_TOPMOST = -1
    SWP_SHOWWINDOW = &H40
End Enum
#If False Then
Private SPI_SCREENSAVERRUNNING, WS_CHILD, GWL_STYLE, GWL_HWNDPARENT, HWND_TOP, HWND_TOPMOST, SWP_NOSIZE, SWP_NOMOVE, SWP_SHOWWINDOW
#End If

Public hThumb   As Long

Public Sub Main()

  Dim rctThumb As RECT
  Dim PrevState As Long

    With fCanvas
        Select Case LCase$(Left$(Command, 2))
        
          Case "/s"   'Screen Saver run or Preview reuqeste
            SetWindowPos .hWnd, HWND_TOPMOST, 0&, 0&, Screen.Width, Screen.Height, 0&
            hThumb = 0
            SystemParametersInfo SPI_SCREENSAVERRUNNING, True, PrevState, 0
            ShowCursor 0
            .Show vbModal
            ShowCursor 1
            SystemParametersInfo SPI_SCREENSAVERRUNNING, False, PrevState, 0
          
          Case "/p"   'Thumbnail Preview requested
            hThumb = Val(Mid$(Command$, 4))
            GetClientRect hThumb, rctThumb
            SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) Or WS_CHILD
            SetParent .hWnd, hThumb
            SetWindowLong .hWnd, GWL_HWNDPARENT, hThumb
            SetWindowPos .hWnd, HWND_TOP, 0, 0, rctThumb.Right, rctThumb.Bottom, SWP_SHOWWINDOW

          Case "/a"   'Change Password requested
            'do nothing

          Case "/c"   'Configure requested
            MsgBox "This Sreensaver has no options.", vbInformation, "Ulli's Screensaver"

        End Select
    End With 'FCANVAS

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Sep-09 14:39)  Decl: 34  Code: 41  Total: 75 Lines
':) CommentOnly: 1 (1,3%)  Commented: 8 (10,7%)  Empty: 16 (21,3%)  Max Logic Depth: 3
