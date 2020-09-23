Attribute VB_Name = "CreateMainWindow"
'Module name: Input Box & Combo Box
'Version 1.0
'Date: march 2002
'Programming: M.C

'Very helpful informations found at:
'KPD-Team
'URL: http://www.allapi.net/

'? cant figure out how to convince text portion of combobox to accept
'keyboard input - have no idea why this doesn't work automaticaly

Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal WindowHwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal WindowHwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Declare Function ShowWindow Lib "user32" (ByVal WindowHwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal WindowHwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'  Define information of the window (pointed to by WindowHwnd)
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal WindowHwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal WindowHwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal WindowHwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Type POINTAPI
    x As Long
    y As Long
End Type
Type Msg
    WindowHwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' Class styles
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000
' Window styles
Public Const WS_ACTIVECAPTION = &H1
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000         ' WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_LAYOUTRTL = &H400000
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_NOACTIVATE = &H8000000
Public Const WS_EX_NOINHERITLAYOUT = &H100000
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_STATICEDGE = &H20000

Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_TABSTOP = &H10000
Public Const WS_GROUP = &H20000
Public Const WS_GT = (WS_GROUP Or WS_TABSTOP)
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME



Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000


' Color constants
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
' Window messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
' ShowWindow commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
' Standard ID's of cursors
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
Public Const GWL_WNDPROC = -4

'ComboBox styles
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_DISABLENOSCROLL = &H800&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_LOWERCASE = &H4000&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_SIMPLE = &H1&
Public Const CBS_SORT = &H100&
Public Const CBS_UPPERCASE = &H2000&

Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_GETEDITSEL = &H140
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETHORIZONTALEXTENT = &H15D
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLOCALE = &H15A
Public Const CB_GETTOPINDEX = &H15B
Public Const CB_INITSTORAGE = &H161
Public Const CB_INSERTSTRING = &H14A
Public Const CB_LIMITTEXT = &H141
Public Const CB_MSGMAX = &H15B
Public Const CB_MULTIPLEADDSTRING = &H163
Public Const CB_OKAY = 0
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_SETEDITSEL = &H142
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_SETHORIZONTALEXTENT = &H15E
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_SETLOCALE = &H159
Public Const CB_SETTOPINDEX = &H15C

Public Const WM_GETTEXT = &HD



'Dim windows, that is to be created Hwnd-s
Dim WindowHwnd As Long
Dim FormWindowHwnd As Long
Dim OKButtonHwnd  As Long
Dim CancelButtonHwnd As Long
Dim TextHwnd As Long
Dim ComboBoxHwnd As Long
Dim ComboBoxListPortionHwnd As Long
Dim win As Long
'Dimension procedures needed for that windows
Dim OKButtonOldProc As Long
Dim OKButtonNewProc As Long
Dim CancelButtonOldProc As Long
Dim CancelButtonNewProc As Long
Dim ComboBoxOldProc As Long
Dim ComboBoxNewProc As Long

'and some public dims
Dim InputBoxCaption As String
Dim InputBoxText As String
Dim AddItems() As Variant  'array
Public IBCBSelectedItem As Variant 'this will hold whatever you select in combo box
Public Sub CreateComboBoxInputBox(Caption As String, IBText As String, InputAddItems() As Variant)
InputBoxCaption = Caption
InputBoxText = IBText
AddItems = InputAddItems

Main
End Sub

Public Sub Main()
    
    Dim lngTemp As Long
    ' Register class
    If MyRegisterClass Then
        ' Window created?
        If MyCreateWindow Then
        ' Change the button's procedures
        ' Point to new address
            'OK Button
            OKButtonNewProc = GetMyWndProc(AddressOf OKButtonProc)
            OKButtonOldProc = SetWindowLong(OKButtonHwnd, GWL_WNDPROC, OKButtonNewProc)
            'Cancel Button
            CancelButtonNewProc = GetMyWndProc(AddressOf CancelButtonProc)
            CancelButtonOldProc = SetWindowLong(CancelButtonHwnd, GWL_WNDPROC, CancelButtonNewProc)
            'ComboBox
            ComboBoxNewProc = GetMyWndProc(AddressOf ComboBoxProc)
            ComboBoxOldProc = SetWindowLong(ComboBoxHwnd, GWL_WNDPROC, ComboBoxNewProc)
            
            ' Message loop
            MyMessageLoop
        End If
        ' Unregister Class
        MyUnregisterClass
    End If
End Sub
Private Function MyRegisterClass() As Boolean
    ' WNDCLASS-structure
    Dim wndcls As WNDCLASS
    wndcls.style = CS_HREDRAW + CS_VREDRAW
    wndcls.lpfnwndproc = GetMyWndProc(AddressOf MyWndProc)
    wndcls.cbClsextra = 0
    wndcls.cbWndExtra2 = 0
    wndcls.hInstance = App.hInstance
    wndcls.hIcon = 0
    wndcls.hCursor = LoadCursor(0, IDC_ARROW)
    wndcls.hbrBackground = COLOR_WINDOW
    wndcls.lpszMenuName = 0
    wndcls.lpszClassName = "myWindowClass"
    ' Register class
    MyRegisterClass = (RegisterClass(wndcls) <> 0)
End Function
Private Sub MyUnregisterClass()
    UnregisterClass "myWindowClass", App.hInstance
End Sub
Private Function MyCreateWindow() As Boolean
    Dim WindowHwnd As Long
    ' Create the window
    FormWindowHwnd = CreateWindowEx(0, "myWindowClass", InputBoxCaption, WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME, (Screen.Width / Screen.TwipsPerPixelX / 2) - 150, (Screen.Height / Screen.TwipsPerPixelY / 2) - 80, 300, 160, 0, 0, App.hInstance, ByVal 0&)
    ' The Buttons and Textbox are child windows
    OKButtonHwnd = CreateWindowEx(0, "Button", "OK", WS_CHILD, 230, 10, 60, 25, FormWindowHwnd, 0, App.hInstance, ByVal 0&)
    CancelButtonHwnd = CreateWindowEx(0&, "button", "Cancel", WS_CHILD, 230, 45, 60, 25, FormWindowHwnd&, 0&, App.hInstance, 0&)
    TextHwnd = CreateWindowEx(&H0, "static", InputBoxText, WS_CHILD, 5, 10, 200, 40, FormWindowHwnd&, 0&, App.hInstance, 0&)
    ComboBoxHwnd = CreateWindowEx(0&, "combobox", "H", CBS_DROPDOWN Or CBS_HASSTRINGS Or CBS_SORT Or WS_CHILD Or WS_VSCROLL, 5, 80, 285, 100, FormWindowHwnd&, 0, App.hInstance, 0&)
    
    If FormWindowHwnd <> 0 Then ShowWindow FormWindowHwnd, SW_SHOWNORMAL
    ' Show them
    ShowWindow OKButtonHwnd, SW_SHOWNORMAL
    ShowWindow CancelButtonHwnd, SW_SHOWNORMAL
    ShowWindow TextHwnd, SW_SHOWNORMAL
    ShowWindow ComboBoxHwnd, SW_SHOWNORMAL
    'fill the combo box with our desired items
    For i = 0 To UBound(AddItems)
    a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr(AddItems(i)))
    
    Next i
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("alfa"))
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("gama"))
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("beta"))
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("alfa"))
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("gama"))
    'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr("beta"))
    
    
    ' Go back
    MyCreateWindow = (FormWindowHwnd <> 0)
End Function
Private Function MyWndProc(ByVal WindowHwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case message
        Case WM_DESTROY
            ' Destroy window
            PostQuitMessage (0)
    End Select
    ' calls the default window procedure
    MyWndProc = DefWindowProc(WindowHwnd, message, wParam, lParam)
End Function
Function GetMyWndProc(ByVal lWndProc As Long) As Long
    GetMyWndProc = lWndProc
End Function
Private Sub MyMessageLoop()
    Dim aMsg As Msg
    Do While GetMessage(aMsg, 0, 0, 0)
        DispatchMessage aMsg
    Loop
End Sub
Private Function OKButtonProc(ByVal WindowHwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        If (message = 533) Then 'You clicked OK
                 Dim MyStr As String
                 'Create a buffer
                 MyStr = String(GetWindowTextLength(ComboBoxHwnd) + 1, Chr$(0))
                 'Get the window's text
                 a = GetWindowText(ComboBoxHwnd, MyStr, Len(MyStr))
                 IBCBSelectedItem = MyStr
                 MsgBox "Selection is now waiting as Public IBCBSelectedItem for further use, you selected:" & IBCBSelectedItem
                 DestroyWindow FormWindowHwnd 'kill our window
                 
                 
        End If
    ' calls the window procedure
    OKButtonProc = CallWindowProc(OKButtonOldProc, WindowHwnd, message, wParam, lParam)
    
End Function

Private Function CancelButtonProc(ByVal WindowHwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim x As Integer
    If (message = 533) Then
       DestroyWindow FormWindowHwnd 'kill our window
    End If
    ' calls the window procedure
    CancelButtonProc = CallWindowProc(CancelButtonOldProc, WindowHwnd, message, wParam, lParam)
End Function

Private Function ComboBoxProc(ByVal WindowHwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'I Hoped to catch here whatewer user selected,
    'no success, so any ideas ? kozlicki@yahoo.com
    Select Case message

    Case Else
    
    End Select
    ' calls the window procedure
    ComboBoxProc = CallWindowProc(ComboBoxOldProc, WindowHwnd, message, wParam, lParam)
End Function

