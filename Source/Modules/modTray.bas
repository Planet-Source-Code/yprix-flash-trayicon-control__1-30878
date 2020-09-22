Attribute VB_Name = "modTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
    NOTIFYICONDATA) As Long
    
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
        ByVal hWnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function TrackPopupMenu Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nReserved As Long, _
        ByVal hWnd As Long, _
        ByVal lprc As Any) As Long
        
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As Any) As Long
        
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpString As Any) As Long
        
Public Declare Function DestroyMenu Lib "user32" _
        (ByVal hMenu As Long) As Long
        
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Public Declare Function DrawCaption Lib "user32" _
        (ByVal hWnd As Long, _
        ByVal hdc As Long, _
        pcRect As RECT, _
        ByVal un As Long) As Long
        
Public Declare Function GetMenuItemRect Lib "user32" _
        (ByVal hWnd As Long, ByVal hMenu As Long, _
        ByVal uItem As Long, _
        lprcItem As RECT) As Long

Public Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long

Public Declare Function GetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long) As Long
        
Public Declare Function SetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long) As Long
        
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204

Public Const GWL_WNDPROC = -4
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117

Public Const MF_APPEND = &H100&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_DEFAULT = &H1000&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_REMOVE = &H1000&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MF_CHECKED = &H8&
Public Const MFT_RADIOCHECK = &H200&

Public Const TPM_RETURNCMD = &H100&

Public Const DC_GRADIENT = &H20
Public Const DC_ACTIVE = &H1
Public Const DC_ICON = &H4
Public Const DC_SMALLCAP = &H2
Public Const DC_TEXT = &H8


Private hMenu As Long
Private hSubmenu As Long
Private chkMnuFlags(2) As Long
Private MP As PointAPI
Private sMenu As Long
Private mnuHeight As Single
Private RectPopup As RECT
Private lpPrevWndProc As Long
Private ghw As Long
Private AppForm As Form
Private TrayData As TrayIcon
Private MenuHandle As Long


Public Sub Hook(frm As Form, tr As TrayIcon)
    On Error Resume Next
        Set TrayData = tr
        Set AppForm = frm
        ghw = frm.hWnd
        lpPrevWndProc = SetWindowLong(ghw, GWL_WNDPROC, AddressOf WindowProc)
        chkMnuFlags(0) = MFT_RADIOCHECK Or MF_CHECKED
        chkMnuFlags(2) = MF_CHECKED
        MenuPopUp
End Sub

Public Sub HookAgain()
    On Error Resume Next
    If Not AppForm Is Nothing Then
        ghw = AppForm.hWnd
        lpPrevWndProc = SetWindowLong(ghw, GWL_WNDPROC, AddressOf WindowProc)
        chkMnuFlags(0) = MFT_RADIOCHECK Or MF_CHECKED
        chkMnuFlags(2) = MF_CHECKED
        MenuPopUp
    End If
End Sub

Public Sub UnHook()
    On Error Resume Next
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(ghw, GWL_WNDPROC, lpPrevWndProc)
    DestroyMenu hMenu
End Sub

Private Function WindowProc(ByVal hWnd As Long, _
                    ByVal uMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long

    On Error Resume Next

    Select Case uMsg
        
        Case WM_MEASUREITEM
             MeasureMenu lParam
             WindowProc = 0
        
        Case WM_DRAWITEM
             DrawMenu lParam
             WindowProc = 0
        
        Case Else
             WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
    
    End Select

End Function
Public Sub MeasureMenu(ByRef lP As Long)
    
    Dim MIS As MEASUREITEMSTRUCT
    
    CopyMemory MIS, ByVal lP, Len(MIS)
    MIS.itemWidth = 5
    CopyMemory ByVal lP, MIS, Len(MIS)
    
End Sub

Public Sub DrawMenu(ByRef lP As Long)
    'On Error Resume Next
    Dim DIS As DRAWITEMSTRUCT, rct As RECT, lRslt As Long
    Dim mnuItms, mnuSeps, mnuWidth, I As Long
    
    CopyMemory DIS, ByVal lP, Len(DIS)
    
    With AppForm
        mnuItms = 0
        mnuSeps = 0
        For I = 1 To TrayData.MenuItems.Count
            With TrayData.MenuItems.Item(I)
                If .MenuType = mtSeparator Then
                    mnuSeps = mnuSeps + 1
                Else
                    mnuItms = mnuItms + 1
                End If
            End With
        Next
        mnuHeight = 0
       
       'String Menus
        GetMenuItemRect .hWnd, hMenu, 1, rct
        mnuHeight = (rct.Bottom - rct.Top) * mnuItms
        RectPopup.Bottom = mnuHeight
        
        'Separators
        GetMenuItemRect .hWnd, hMenu, 3, rct
        mnuHeight = (mnuHeight + (rct.Bottom - rct.Top) * mnuSeps)
        RectPopup.Bottom = mnuHeight
        
        'set the size of our sidebar
        mnuWidth = 18
        SetRect rct, 0, 0, mnuHeight, mnuWidth
        If TrayData.CaptionPicture Is Nothing Then
           DrawCaption .hWnd, .hdc, rct, DC_SMALLCAP Or DC_ACTIVE Or DC_TEXT Or DC_GRADIENT
        End If
        
        AppForm.Print "   " & AppForm.Caption
        
        Dim X As Single, Y As Single
        Dim nColor As Long
        For X = 0 To mnuHeight
            For Y = 0 To mnuWidth - 1
                nColor = GetPixel(.hdc, X, Y)
                SetPixel DIS.hdc, Y, mnuHeight - X, nColor
            Next Y
        Next X
        .Cls
     End With
    RectPopup = rct
    
End Sub

Public Sub MenuPopUp()
    
    On Error Resume Next
    Dim X, I As Integer
    Dim Displ As Byte
    Dim cnt As Integer
    Dim style As Integer
    Dim mnu As Long
    hMenu = CreatePopupMenu()
    AppendMenu hMenu, MF_OWNERDRAW Or MF_DISABLED, 1000, 0& 'SideBar
    
    For I = 1 To TrayData.MenuItems.Count
        With TrayData.MenuItems.Item(I)
            If .MenuType = mtSeparator Then
                 AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
            Else
                If I <> 1 Then
                    style = 0&
                Else
                    style = MF_MENUBARBREAK
                End If
                
                If .Checked = True Then
                    style = style + MF_CHECKED
                End If
                
                If .Grayed = True Then
                    style = style + MF_GRAYED
                End If
                
                If .IsSubMenu = True Then
                    mnu = AppendMenu(hMenu, style + MF_POPUP, .MenuItems.GetMenu, .Caption)
    
                Else
                    mnu = AppendMenu(hMenu, style, .ID, .Caption)
                End If
                
                SetMenuItemBitmaps hMenu, I, MF_BYPOSITION, .Icon.Handle, .Icon.Handle
                
            End If
        End With
    Next
 End Sub

Public Sub MenuTrack(frm As Form)
    On Error Resume Next
    
    Dim strNr  As String
    Dim locVal As Long
    GetCursorPos MP
    
    DoEvents
    sMenu = TrackPopupMenu(hMenu, TPM_RETURNCMD, MP.X, MP.Y, 0, frm.hWnd, 0&)
    locVal = sMenu
    TrayData.NotifyMenuItemSelect (locVal)
End Sub

