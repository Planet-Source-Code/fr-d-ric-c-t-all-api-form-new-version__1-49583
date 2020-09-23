Attribute VB_Name = "APIControls"
Option Explicit

'Control specific constants
Private Const ES_AUTOHSCROLL              As Long = &H80&
Private Const BS_AUTORADIOBUTTON          As Long = &H9&
Private Const BS_AUTO3STATE               As Long = &H6&
Private Const BS_AUTOCHECKBOX             As Long = &H3&
Private Const BS_LEFTTEXT                 As Long = &H20&
Private Const BS_MULTILINE                As Long = &H2000&
Public Const BST_CHECKED                  As Long = &H1
Public Const BST_INDETERMINATE            As Long = &H2 'Grayed checkbox
Public Const BST_UNCHECKED                As Long = &H0
Public Const EN_CHANGE                    As Long = &H300
Public Const BN_CLICKED                   As Long = 0
Public Const BM_CLICK                     As Long = &HF5
Public Const BM_SETCHECK                  As Long = &HF1
Public Const BM_GETCHECK                  As Long = &HF0
'Window styles
Private Const WS_CHILD                    As Long = &H40000000
Private Const WS_VISIBLE                  As Long = &H10000000
Private Const WS_BORDER                   As Long = &H800000
Private Const WS_SYSMENU                  As Long = &H80000
Private Const WS_CAPTION                  As Long = &HC00000
Private Const WS_EX_CLIENTEDGE            As Long = &H200&
Private Const WS_TABSTOP                  As Long = &H10000
Private Const WS_MAXIMIZEBOX              As Long = &H10000
Private Const WS_MINIMIZEBOX              As Long = &H20000
Private Const WS_THICKFRAME               As Long = &H40000
Private Const WS_GROUP                    As Long = &H20000
'Window show state
Private Const SW_NORMAL                   As Long = 1  'Use when you show for the first time
Private Const SW_HIDE                     As Long = 0
Private Const SW_SHOW                     As Long = 5
'Window messages
Private Const WM_USER                     As Long = &H400
Public Const WM_DESTROY                   As Long = &H2
Public Const WM_SETFONT                   As Long = &H30
Public Const WM_SETTEXT                   As Long = &HC
Public Const WM_COMMAND                   As Long = &H111
Public Const WM_CTLCOLOREDIT              As Long = &H133
Public Const WM_MENUCOMMAND               As Long = &H126
Public Const WM_QUIT                      As Long = &H12
'Menu constants
Public Const MIIM_BITMAP                  As Long = &H80
Public Const MIIM_CHECKMARKS              As Long = &H8
Public Const MIIM_DATA                    As Long = &H20
Public Const MIIM_FTYPE                   As Long = &H100
Public Const MIIM_ID                      As Long = &H2
Public Const MIIM_STATE                   As Long = &H1
Public Const MIIM_STRING                  As Long = &H40
Public Const MIIM_SUBMENU                 As Long = &H4
Public Const MIIM_TYPE                    As Long = &H10
Public Const MFT_MENUBARBREAK             As Long = &H20&
Public Const MFT_MENUBREAK                As Long = &H40&
Public Const MFT_OWNERDRAW                As Long = &H100&
Public Const MFT_RADIOCHECK               As Long = &H200&
Public Const MFT_RIGHTJUSTIFY             As Long = &H4000&
Public Const MFT_RIGHTORDER               As Long = &H2000&
Public Const MFT_SEPARATOR                As Long = &H800&
Public Const MFS_CHECKED                  As Long = &H8&
Public Const MFS_DEFAULT                  As Long = &H1000&
Public Const MFS_DISABLED                 As Long = &H3&
Public Const MFS_ENABLED                  As Long = &H0&
Public Const MFS_HILITE                   As Long = &H80&
Public Const MFS_UNCHECKED                As Long = &H0&
Public Const MFS_UNHILITE                 As Long = &H0&
'System colors
Public Const COLOR_ACTIVEBORDER           As Long = 10
Public Const COLOR_ACTIVECAPTION          As Long = 2
Public Const COLOR_APPWORKSPACE           As Long = 12
Public Const COLOR_BACKGROUND             As Long = 1
Public Const COLOR_BTNFACE                As Long = 15
Public Const COLOR_BTNSHADOW              As Long = 16
Public Const COLOR_BTNTEXT                As Long = 18
Public Const COLOR_CAPTIONTEXT            As Long = 9
Public Const COLOR_GRAYTEXT               As Long = 17
Public Const COLOR_HIGHLIGHT              As Long = 13
Public Const COLOR_HIGHLIGHTTEXT          As Long = 14
Public Const COLOR_INACTIVEBORDER         As Long = 11
Public Const COLOR_INACTIVECAPTION        As Long = 3
Public Const COLOR_MENU                   As Long = 4
Public Const COLOR_MENUTEXT               As Long = 7
Public Const COLOR_SCROLLBAR              As Long = 0
Public Const COLOR_WINDOW                 As Long = 5
Public Const COLOR_WINDOWFRAME            As Long = 6
Public Const COLOR_WINDOWTEXT             As Long = 8
'Others
Private Const DEFAULT_GUI_FONT            As Long = 17

Public Enum WNDSTYLE
    [FixedNoSysMenu] = 0
    [FixedSysMenu] = WS_SYSMENU + WS_CAPTION
    [FixedMin] = WS_SYSMENU + WS_CAPTION + WS_MINIMIZEBOX
    [FixedMax] = WS_SYSMENU + WS_CAPTION + WS_MAXIMIZEBOX
    [FixedMinMax] = WS_SYSMENU + WS_CAPTION + WS_MINIMIZEBOX + WS_MAXIMIZEBOX
    [Resizable] = WS_SYSMENU + WS_CAPTION + WS_THICKFRAME
    [ResizableMin] = WS_SYSMENU + WS_CAPTION + WS_THICKFRAME + WS_MINIMIZEBOX
    [ResizableMax] = WS_SYSMENU + WS_CAPTION + WS_THICKFRAME + WS_MAXIMIZEBOX
    [ResizableMinMax] = WS_SYSMENU + WS_CAPTION + WS_THICKFRAME + WS_MINIMIZEBOX + WS_MAXIMIZEBOX
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private FixedNoSysMenu, FixedSysMenu, FixedMin, FixedMax, FixedMinMax, Resizable, ResizableMin
Private ResizableMax, ResizableMinMax
#End If

'Window Structure
Public Type WNDCLASS
    Style                                   As Long
    lpfnwndproc                             As Long
    cbClsextra                              As Long
    cbWndExtra2                             As Long
    hInstance                               As Long
    hIcon                                   As Long
    hCursor                                 As Long
    hbrBackground                           As Long
    lpszMenuName                            As String
    lpszClassName                           As String
End Type

'Mouse location structure
Public Type POINTAPI
    X                                       As Long
    Y                                       As Long
End Type

'Window Message structure
Public Type MSG
    hwnd                                    As Long
    message                                 As Long
    wParam                                  As Long
    lParam                                  As Long
    time                                    As Long
    pt                                      As POINTAPI
End Type

'Menu item structure
Public Type MENUITEMINFO
    cbSize                                  As Long
    fMask                                   As Long
    fType                                   As Long
    fState                                  As Long
    wID                                     As Long
    hSubMenu                                As Long
    hbmpChecked                             As Long
    hbmpUnchecked                           As Long
    dwItemData                              As Long
    dwTypeData                              As String
    cch                                     As Long
End Type

Private hFont                               As Long

'APIs used to Create the form, controls and perform message manipulation
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function GetNextDlgTabItem Lib "user32.dll" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As MSG) As Long
Public Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsDialogMessage Lib "user32.dll" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As MSG) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function InvalidateRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function CreateMenu Lib "user32.dll" () As Long
Public Declare Function SetMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function CheckMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Function CreateButton(ByVal hParent As Long, ByVal strCaption As String, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, ByVal Style As Long) As Long

Dim hTemp As Long

    'Creates a button and returns its handle
    hTemp = CreateWindowEx(0&, "BUTTON", strCaption, Style, X, Y, width, height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hFont = 0 Then
        hFont = GetStockObject(DEFAULT_GUI_FONT)
    End If
    SendMessage hTemp, WM_SETFONT, hFont, 1
    CreateButton = hTemp

End Function

Public Function CreateCheckbox(ByVal hParent As Long, ByVal strCaption As String, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                               Optional ByVal bln3States As Boolean = False, Optional ByVal blnBeginGroup As Boolean = False, Optional ByVal blnTabStop As Boolean = True, _
                               Optional ByVal blnMultiLineCaption As Boolean = True, Optional ByVal blnLeftText As Boolean = False) As Long

Dim lngStyle As Long

    'Create a checkbox
    If hParent = 0 Then Exit Function
    lngStyle = WS_CHILD Or WS_VISIBLE
    If blnTabStop Then
        lngStyle = lngStyle Or WS_TABSTOP
    End If
    If blnBeginGroup Then
        lngStyle = lngStyle Or WS_GROUP
    End If
    If blnMultiLineCaption Then
        lngStyle = lngStyle Or BS_MULTILINE
    End If
    If bln3States Then
        lngStyle = lngStyle Or BS_AUTO3STATE
    Else
        lngStyle = lngStyle Or BS_AUTOCHECKBOX
    End If
    If blnLeftText Then
        lngStyle = lngStyle Or BS_LEFTTEXT
    End If
    CreateCheckbox = CreateButton(hParent, strCaption, X, Y, width, height, lngStyle)

End Function

Public Function CreateCmdButton(ByVal hParent As Long, ByVal strCaption As String, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                                Optional ByVal blnBeginGroup As Boolean = False, Optional ByVal blnTabStop As Boolean = True, _
                                Optional ByVal blnMultiLineCaption As Boolean = True) As Long

Dim lngStyle As Long

    'Create a command button
    If hParent = 0 Then Exit Function
    lngStyle = WS_CHILD Or WS_VISIBLE
    If blnTabStop Then
        lngStyle = lngStyle Or WS_TABSTOP
    End If
    If blnBeginGroup Then
        lngStyle = lngStyle Or WS_GROUP
    End If
    If blnMultiLineCaption Then
        lngStyle = lngStyle Or BS_MULTILINE
    End If
    CreateCmdButton = CreateButton(hParent, strCaption, X, Y, width, height, lngStyle)

End Function

Public Function CreateRadioButton(ByVal hParent As Long, ByVal strCaption As String, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, _
                                  Optional ByVal blnBeginGroup As Boolean = False, Optional ByVal blnTabStop As Boolean = True, _
                                  Optional ByVal blnMultiLineCaption As Boolean = True, Optional ByVal blnLeftText As Boolean = False) As Long

Dim lngStyle As Long

    'Create a radio button
    If hParent = 0 Then Exit Function
    lngStyle = WS_CHILD Or WS_VISIBLE Or BS_AUTORADIOBUTTON
    If blnTabStop Then
        lngStyle = lngStyle Or WS_TABSTOP
    End If
    If blnBeginGroup Then
        lngStyle = lngStyle Or WS_GROUP
    End If
    If blnLeftText Then
        lngStyle = lngStyle Or BS_LEFTTEXT
    End If
    If blnMultiLineCaption Then
        lngStyle = lngStyle Or BS_MULTILINE
    End If
    CreateRadioButton = CreateButton(hParent, strCaption, X, Y, width, height, lngStyle)

End Function

Public Function CreateForm(ByVal strTitle As String, ByVal winStyle As WNDSTYLE, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, Optional lngColor As Long = COLOR_BTNFACE) As Long

Dim wc As WNDCLASS, hTemp As Long

    'Creates a form and returns its handle
    With wc
        .lpfnwndproc = GetAdd(AddressOf WndProc)
        .hbrBackground = lngColor + 1
        .lpszClassName = "CustomClass"
    End With
    RegisterClass wc
    hTemp = CreateWindowEx(0&, "CustomClass", strTitle, winStyle, X, Y, width, height, 0, 0, App.hInstance, ByVal 0&)
    ShowWindow hTemp, SW_NORMAL
    UpdateWindow hTemp
    SetFocus hTemp
    CreateForm = hTemp

End Function

Public Function CreateLabel(ByVal hParent As Long, ByVal strCaption As String, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long) As Long

Dim hTemp As Long

    'Creates a label and returns its handle
    If hParent = 0 Then Exit Function
    hTemp = CreateWindowEx(0&, "STATIC", strCaption, WS_CHILD Or WS_VISIBLE, X, Y, width, height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hFont = 0 Then
        hFont = GetStockObject(DEFAULT_GUI_FONT)
    End If
    SendMessage hTemp, WM_SETFONT, hFont, 1
    CreateLabel = hTemp

End Function

Public Function CreateMenuItem(ByVal strCaption As String, ByVal ID As Integer, ByVal hSubMenu As Long, _
                               Optional ByVal blnSeparator As Boolean = False, Optional ByVal blnDisabled As Boolean = False, _
                               Optional ByVal blnChecked As Boolean = False) As MENUITEMINFO

Dim mnuInfo As MENUITEMINFO

    With mnuInfo
        .cbSize = Len(mnuInfo)
        .fMask = MIIM_ID Or MIIM_STRING
        If hSubMenu Then
            .fMask = .fMask Or MIIM_SUBMENU
            .hSubMenu = hSubMenu
        End If
        If blnSeparator Then
            .fMask = .fMask Or MIIM_FTYPE
            .fType = MFT_SEPARATOR
        End If
        If blnDisabled Or blnChecked Then
            .fMask = .fMask Or MIIM_STATE
            If blnDisabled Then
                .fState = MFS_DISABLED
            Else
                .fState = MFS_CHECKED
            End If
        End If
        .wID = ID
        .dwTypeData = strCaption
        .cch = Len(.dwTypeData)
    End With
    CreateMenuItem = mnuInfo

End Function

Public Function CreateTextbox(ByVal hParent As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long) As Long

Dim hTemp As Long

    'Creates a textbox and returns its handle
    If hParent = 0 Then Exit Function
    hTemp = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", vbNullString, WS_CHILD Or WS_VISIBLE Or WS_BORDER Or ES_AUTOHSCROLL Or WS_TABSTOP, X, Y, width, height, hParent, 0, App.hInstance, ByVal 0&)
    If hFont = 0 Then
        hFont = GetStockObject(DEFAULT_GUI_FONT)
    End If
    SendMessage hTemp, WM_SETFONT, hFont, 1
    CreateTextbox = hTemp

End Function

Public Function GetAdd(Address As Long) As Long

    'Workaround Function since AddressOf can only be used part of a parameter
    GetAdd = Address

End Function

Public Sub GetAndDispatch(ByVal hParent As Long)

Dim udtMsg As MSG

    If hParent = 0 Then Exit Sub
    Do While GetMessage(udtMsg, 0, 0, 0)
        If IsDialogMessage(hParent, udtMsg) = 0 Then
            TranslateMessage udtMsg
            DispatchMessage udtMsg
        End If
    Loop
    UnregisterClass "CustomClass", App.hInstance

End Sub
