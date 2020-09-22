Attribute VB_Name = "mMsgBox"
Option Explicit


'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*
'
' Study this module carefully... do not make any changes until you are
' certain how these procedures work.
'
'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*


Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Type WNDCLASS
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


Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const COLOR_WINDOW = 5

' A P I   D E F I N E D   C O N S T A N T S
 Const CS_VREDRAW = &H1
 Const CS_HREDRAW = &H2

 Const CW_USEDEFAULT = &H80000000

 Const ES_MULTILINE = &H4&
 Const ES_READONLY = &H800&

 Const WS_BORDER = &H800000
 Const WS_CHILD = &H40000000
 Const WS_OVERLAPPED = &H0&
 Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
 Const WS_SYSMENU = &H80000
 Const WS_THICKFRAME = &H40000
 Const WS_MINIMIZEBOX = &H20000
 Const WS_MAXIMIZEBOX = &H10000
 Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME)


 Const WS_EX_CLIENTEDGE = &H200&


 Const WM_CLOSE = &H10
 Const WM_DESTROY = &H2
 Const WM_LBUTTONUP = &H202

 Const IDC_ARROW = 32512&

 Const IDI_APPLICATION = 32512&
 Const IDI_EXCLAMATION = 32515&

 Public Const GWL_WNDPROC = (-4)

 Const SW_SHOWNORMAL = 1
 



 Const GWL_HINSTANCE = (-6)
 Const GWL_HWNDPARENT = (-8)
 Const GWL_STYLE = (-16)
 Const GWL_EXSTYLE = (-20)
 Const GWL_USERDATA = (-21)
 Const GWL_ID = (-12)





Private m_ButOldProc As Long ''Will hold address of the old window proc for the button

' Handles for the windows on the PWord Dialog Box
Private m_Hwnd               As Long        ' Main Window

Private m_ButtonHwnd_OK      As Long        ' OK Button
Private m_ButtonHwnd_CANCEL  As Long        ' Cancel Button
Private m_ButtonHwnd_YES     As Long        ' Yes Button
Private m_ButtonHwnd_NO      As Long        ' NO Button

Private m_EditHwnd           As Long        ' Edit Control

Const gClassName = "WinClass"


' \\ M E S S A G E   B O X   S I Z E S

Private Const BX_MINWIDTH = 270&
Private Const BX_MINHEIGHT = 100&

Private Const WORD_HEIGHT = 20&


Private sIcon_Path As String


Private iMsgLeft As Integer
Private iIconLeft As Integer

Private ReturnValue As Integer
Private bIcon As Boolean

Function MBox(sMsg, _
              sCaption As String, _
    Optional lType As Long, _
    Optional lIcon As Long, _
    Optional oPath As String, _
    Optional nX As Long, _
    Optional nY As Long, _
    Optional Edge As Boolean _
        ) As Integer

    Dim lWidth As Long
    Dim lHeight As Long
    Dim iWordMultiple As Integer

    Dim lMsgWidth As Long
    Dim lMsgHeight As Long
    
    Dim ImgSpace As Integer
    Dim LongestLine As Long
    Dim Multiple As Integer
    Dim hIcon As Long
    Dim Buff As Integer

     '
     ' Is an Icon being used?
     '
     If lIcon < 5 Then
        ' N O
        iMsgLeft = 16
        Buff = 0
        ImgSpace = 0
     Else
        ' Evaluate the Icon preference
        Select Case lIcon
           
           Case IC_CUSTOM
               ' Store the Path
               sIcon_Path = oPath
               ' load the Icon
               ' I know it sucks to have to add a picturebox!!
               ' Anyone know how to do it with out??
               Form1.Picture1.Picture = LoadPicture(oPath)
        End Select
        
        iMsgLeft = 64 + 16
        Buff = 25
        ImgSpace = 64
     End If
     
     If lType = 0 Then lType = 1&
     
     '
     ' Get the Length of the message to determine the height
     ' and width of the Message Box
     '
     
     iWordMultiple = CountLetters(sMsg)
     
     
       
     If iWordMultiple > 1 Then
        lWidth = 640
        lMsgWidth = 590
     Else
     
     
       If lType = BX_YESNOCANCEL Then
        lWidth = (Len(sMsg) + (16 + 32)) * 5.5
        lMsgWidth = (lWidth - (50 + ImgSpace) + Buff)
       Else
        lWidth = (Len(sMsg) + (16 + 32)) * 4
        lMsgWidth = (lWidth - (50 + ImgSpace) + Buff)
       End If
     End If
     
     '
     '  Check for Carrige Returns
     '
     If CheckCR(sMsg) > 1 Then ' adjust width by longest line
        
        iWordMultiple = CheckCR(sMsg)
        LongestLine = GetLongestLine(sMsg)
        
        Multiple = 6
        
        If LongestLine > 50 And LongestLine < 60 Then
           Multiple = 9
        ElseIf LongestLine > 60 And LongestLine < 70 Then
           Multiple = 10
        ElseIf LongestLine > 70 And LongestLine < 80 Then
           Multiple = 11
        ElseIf LongestLine > 90 And LongestLine < 100 Then
           Multiple = 12
        ElseIf LongestLine > 100 Then
           Multiple = 13
        End If
        
        
        
        lWidth = (Len(LongestLine) + (16 + 32)) * Multiple
        lMsgWidth = (lWidth - (50 + ImgSpace) + Buff)
     End If
     
     lHeight = _
     BX_MINHEIGHT + (WORD_HEIGHT * iWordMultiple)
     
    
     
     ' create the messageBox
     CreateNewMessageBox _
        sMsg, _
        sCaption, _
        lWidth, _
        lHeight, _
        lMsgWidth, _
        iWordMultiple, _
        iMsgLeft, _
        25, _
        lType, lType, oPath, _
        nX, _
        nY, _
        Edge
        
        MBox = ReturnValue
End Function


Private Sub CreateNewMessageBox(sMsg, _
            sCaption As String, _
            lWidth As Long, _
            lHeight As Long, _
            lMsgWidth As Long, _
            iLines As Integer, _
            iLabel_X As Integer, _
            iLabel_Y As Integer, _
            Optional lType As Long, _
            Optional lIcon As Long, _
            Optional oPath As String, _
            Optional nX As Long, _
            Optional nY As Long, _
            Optional Edge As Boolean _
    )

   Dim wMsg As Msg
   
   ''Call procedure to register window classname. If false, then exit.
   If RegisterWindowClass = False Then Exit Sub
    
      ''Create window
      If CreateWindows(sMsg, _
                       sCaption, _
                       lWidth, _
                       lHeight, _
                       iLines, _
                       lMsgWidth, _
                       iLabel_X, _
                       iLabel_Y, _
                       lType, lIcon, oPath, _
                       nX, _
                        nY, _
                       Edge _
                      ) Then
      
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            Call DispatchMessage(wMsg)
         Loop
      End If

    Call UnregisterClass(gClassName$, App.hInstance)


End Sub


Private Function CreateWindows(sMsg, _
            sCaption As String, _
            lWidth As Long, _
            lHeight As Long, _
            iLines As Integer, _
            lMsgWidth As Long, _
            iLabel_X As Integer, _
            iLabel_Y As Integer, _
            Optional lType As Long, _
            Optional lIcon As Long, _
            Optional oPath As String, _
            Optional nX As Long, _
            Optional nY As Long, _
            Optional Edge As Boolean _
    ) As Boolean
       
    Dim WinEdge As Long
    Dim ButtonTop As Integer
    Dim hIcon As Long
    
    Dim wDc As Long
    
    
    
    If Edge Then WinEdge& = WS_EX_CLIENTEDGE& Else WinEdge& = 0&
    
    If sCaption = "" Then sCaption = App.EXEName
    
    CleanUp Form1.Picture1
    
    ''Create Message Box window.
    m_Hwnd& = CreateWindowEx( _
         WinEdge&, _
         gClassName$, _
         sCaption, _
         1&, _
         nX, nY, _
         lWidth, lHeight, _
         0&, _
         0&, _
         App.hInstance, _
         ByVal 0& _
      )
      
    '
    ' Store an area to draw on on the box
    wDc& = GetDC(m_Hwnd&)
    '
    ' Create buttons
    '
    ButtonTop = (iLabel_X + (WORD_HEIGHT * iLines)) + 25
    If lIcon >= 5 Then
         ButtonTop = (iLabel_X + (WORD_HEIGHT * iLines)) - 45
    End If
    CreateButtons lType, lMsgWidth, ButtonTop
    
    '
    ' The Label THe Msg Resides in
    '
    'MsgBox lMsgWidth
    
    m_EditHwnd& = CreateWindowEx( _
        0&, _
        "Edit", _
        sMsg, _
        WS_CHILD Or ES_READONLY Or ES_MULTILINE, _
        iLabel_X, _
        iLabel_Y, _
        lMsgWidth, _
        (WORD_HEIGHT * iLines), _
        m_Hwnd&, _
        0&, _
        App.hInstance, _
        0& _
      )

    
    
    'show  Windows.
    Call ShowWindow(m_Hwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_EditHwnd&, SW_SHOWNORMAL)
    
        
    
    Select Case lType
       Case BX_OK
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
         
       Case BX_OKCANCEL
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
         Call SetWindowLong(m_ButtonHwnd_CANCEL&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
             
       Case BX_YESNO
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_YES&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_YES&, GWL_WNDPROC, GetAddress(AddressOf YesWndProc))
         Call SetWindowLong(m_ButtonHwnd_NO&, GWL_WNDPROC, GetAddress(AddressOf NoWndProc))

       
       Case BX_YESNOCANCEL
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_YES&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_YES&, GWL_WNDPROC, GetAddress(AddressOf YesWndProc))
         Call SetWindowLong(m_ButtonHwnd_NO&, GWL_WNDPROC, GetAddress(AddressOf NoWndProc))
         Call SetWindowLong(m_ButtonHwnd_CANCEL&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))

            
       Case IC_EXCLAME
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
          
          hIcon = LoadStandardIcon(0&, IDI_EXCLAMATION)
          Call DrawIcon(wDc&, 16&, 25&, hIcon)
          
       Case IC_CRITICAL
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
          hIcon = LoadStandardIcon(0&, IDI_HAND)
          Call DrawIcon(wDc&, 16&, 25&, hIcon)
          
       Case IC_INFO
          m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
          hIcon = LoadStandardIcon(0&, IDI_ASTERISK)
          Call DrawIcon(wDc&, 16&, 25&, hIcon)
          
       Case IC_CONFIRM
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
          hIcon = LoadStandardIcon(0&, IDI_QUESTION)
          Call DrawIcon(wDc&, 16&, 25&, hIcon)
          
       Case IC_CUSTOM
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
         
                     
       Case Else
         m_ButOldProc& = GetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC)
         Call SetWindowLong(m_ButtonHwnd_OK&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
       
    End Select

    If lIcon = 9 Then ' custom
       
       AddIcon Form1.Picture1
       bIcon = True
    End If
    
    CreateWindows = (m_Hwnd& <> 0)
    '
End Function

Private Sub CreateButtons( _
                lType As Long, _
                LabelWidth As Long, _
                ButtTop As Integer _
             )
   
    
    'If iLines > 18 Then Offset = 8
    Select Case lType
        Case BX_OK '
           ' OK Button
           m_ButtonHwnd_OK& = CreateWindowEx( _
              0&, _
              "Button", _
              "Ok", _
              WS_CHILD, _
              (LabelWidth - (50)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
           ' display the buttons\Windows
           Call ShowWindow(m_ButtonHwnd_OK&, SW_SHOWNORMAL)
           
        Case BX_OKCANCEL
           ' OK Button
           m_ButtonHwnd_OK& = CreateWindowEx( _
              0&, _
              "Button", _
              "Ok", _
              WS_CHILD, _
              (LabelWidth - (70 * 2)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
           ' Cancel Button
           m_ButtonHwnd_CANCEL& = CreateWindowEx( _
              0, _
              "Button", _
              "&Cancel", _
              WS_CHILD, _
              (LabelWidth - (70 * 2)) + 90, _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
           ' display the buttons\Windows
           Call ShowWindow(m_ButtonHwnd_OK&, SW_SHOWNORMAL)
           Call ShowWindow(m_ButtonHwnd_CANCEL&, SW_SHOWNORMAL)
           
        Case BX_YESNO
           ' YES   B U T T O N
           m_ButtonHwnd_YES& = CreateWindowEx( _
              0&, _
              "Button", _
              "Yes", _
              WS_CHILD, _
              (LabelWidth - (70 * 2)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
            
            ' No Button
           m_ButtonHwnd_NO& = CreateWindowEx( _
              0, _
              "Button", _
              "No", _
              WS_CHILD, _
              (LabelWidth - (70 * 2)) + 90, _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
           ' display the buttons\Windows
           Call ShowWindow(m_ButtonHwnd_YES&, SW_SHOWNORMAL)
           Call ShowWindow(m_ButtonHwnd_NO&, SW_SHOWNORMAL)
                   
        Case BX_YESNOCANCEL
           ' YES   B U T T O N
           m_ButtonHwnd_YES& = CreateWindowEx( _
              0&, _
              "Button", _
              "Yes", _
              WS_CHILD, _
              (LabelWidth - (115 * 2)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
            
            ' No Button
           m_ButtonHwnd_NO& = CreateWindowEx( _
              0, _
              "Button", _
              "No", _
              WS_CHILD, _
              (LabelWidth - (115 * 2)) + 90, _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
            
             ' Cancel Button
           m_ButtonHwnd_CANCEL& = CreateWindowEx( _
              0, _
              "Button", _
              "&Cancel", _
              WS_CHILD, _
              (LabelWidth - (115 * 2)) + (90 * 2), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
           ' display the buttons\Windows
           Call ShowWindow(m_ButtonHwnd_YES&, SW_SHOWNORMAL)
           Call ShowWindow(m_ButtonHwnd_NO&, SW_SHOWNORMAL)
           Call ShowWindow(m_ButtonHwnd_CANCEL&, SW_SHOWNORMAL)
          
        'Case IC_CONFIRM
          
        Case 5, 6, 7, 8
           m_ButtonHwnd_OK& = CreateWindowEx( _
              0&, _
              "Button", _
              "Ok", _
              WS_CHILD, _
              (LabelWidth - (50)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
            Call ShowWindow(m_ButtonHwnd_OK&, SW_SHOWNORMAL)
            'Debug.Print "Make Button"
            
            
        Case Else
           m_ButtonHwnd_OK& = CreateWindowEx( _
              0&, _
              "Button", _
              "Ok", _
              WS_CHILD, _
              (LabelWidth - (50)), _
              ButtTop, 85, 25, _
              m_Hwnd&, _
              0&, _
              App.hInstance, _
              0& _
            )
            Call ShowWindow(m_ButtonHwnd_OK&, SW_SHOWNORMAL)
            Debug.Print "Make Button"
    End Select
    
    
End Sub


Private Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we
    ''can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_EXCLAMATION)  ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function





Private Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim strTemp As String
    
    Select Case uMsg&
       Case WM_DESTROY:
       Call PostQuitMessage(0&)
    End Select
    

  ''Let windows call the default window procedure since we're done.
  WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function

Private Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim TextLen As Long, Length As String
    
    Select Case uMsg&
       Case WM_LBUTTONUP:
          
          ReturnValue = mbYes
          CloseDialog
    End Select
    
FinishUp:
    
  'call the old one using CallWindowProc
  OKWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg&
       Case WM_LBUTTONUP:
           ReturnValue = mbCancel
           CloseDialog
    End Select
    
  CancelWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function YesWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg&
       Case WM_LBUTTONUP:
           ReturnValue = mbYes
           CloseDialog
    End Select
    
  YesWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function NoWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg&
       Case WM_LBUTTONUP:
           ReturnValue = mbNo
           CloseDialog
    End Select
    
  NoWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function



Private Function GetAddress(ByVal lngAddr As Long) As Long
    '
    '  Used with AddressOf to return the
    '  address in memory of a procedure.
    '
    GetAddress = lngAddr&
    
End Function

Private Sub CloseDialog()
    Call SendMessage(m_Hwnd&, WM_CLOSE, 0&, 0&)
End Sub




Private Function GetLongestLine(sMsg)
    '
    Dim i As Integer, x As Integer
    Dim LineLen() As Long
    Dim tmpStr As String
    
    
    For i = 1 To Len(sMsg)
       tmpStr = tmpStr + Mid$(sMsg, i, 1)
       
       If Right$(tmpStr, 2) = (Chr$(13) + Chr$(10)) Then
           x = x + 1
           ReDim Preserve LineLen(x)
           LineLen(x) = Len(tmpStr)
           tmpStr = ""
       End If
       
    Next i
        
    GetLongestLine = GetLargestMember(LineLen)
    
End Function


Private Function GetLargestMember(LineLen As Variant)
    '
    Dim i As Integer, x As Integer
    Dim LongestLine As Integer
    Dim tmpLineLen As Integer
    Dim NumMembers As Integer
    
    NumMembers = UBound(LineLen)
    
    ' get the 1st member of the array
    tmpLineLen = LineLen(1)
    
    Do
        i = i + 1
                    
        ' compare each other member to this one.
        ' until a bigger one is found.
        For x = 1 To NumMembers
           If tmpLineLen < LineLen(x) Then
               tmpLineLen = LineLen(x)
               Exit For
           End If
        Next x
        
        If NumMembers = 1 Then Exit Do
        
    Loop Until (i + 1) = NumMembers
       
    
    GetLargestMember = tmpLineLen + 10
End Function

Private Function CheckCR(sMsg) As Integer
    Dim i As Integer
    Dim CrCnt As Integer
    
    For i% = 1 To Len(sMsg)
        If Mid(sMsg, i%, 2) = Chr(13) + Chr(10) Then
           CrCnt = CrCnt + 1
        End If
    Next i%
    
    CheckCR = CrCnt + 1
End Function




Private Function CountLetters(sMsg) As Integer

  Dim iWordMultiple As Integer
  
    If Len(sMsg) >= 0 And Len(sMsg) < 48 Then
        iWordMultiple = 1
     ElseIf Len(sMsg) >= 48 And Len(sMsg) < (48 * 2) Then  ' Max Char on 1 line
        iWordMultiple = 2
     ElseIf Len(sMsg) >= (48 * 2) And Len(sMsg) < (48 * 3) Then
        iWordMultiple = 3
     ElseIf Len(sMsg) >= (48 * 3) And Len(sMsg) < (48 * 4) Then
        iWordMultiple = 4
     ElseIf Len(sMsg) >= (48 * 4) And Len(sMsg) < (48 * 5) Then
        iWordMultiple = 5
     ElseIf Len(sMsg) >= (48 * 5) And Len(sMsg) < (48 * 6) Then
        iWordMultiple = 6
     ElseIf Len(sMsg) >= (48 * 6) And Len(sMsg) < (48 * 7) Then
        iWordMultiple = 7
     ElseIf Len(sMsg) >= (48 * 7) And Len(sMsg) < (48 * 8) Then
        iWordMultiple = 8
     ElseIf Len(sMsg) >= (48 * 8) And Len(sMsg) < (48 * 9) Then
        iWordMultiple = 9
     ElseIf Len(sMsg) >= (48 * 9) And Len(sMsg) < (48 * 10) Then
        iWordMultiple = 10
     ElseIf Len(sMsg) >= (48 * 10) And Len(sMsg) < (48 * 11) Then
        iWordMultiple = 11
     ElseIf Len(sMsg) >= (48 * 11) And Len(sMsg) < (48 * 12) Then
        iWordMultiple = 12
     ElseIf Len(sMsg) >= (48 * 12) And Len(sMsg) < (48 * 13) Then
        iWordMultiple = 13
                      
     Else
        iWordMultiple = 14
     End If
     
     CountLetters = iWordMultiple
End Function



Sub AddIcon(Pic As PictureBox)
    '
    Dim WinHdc As Long
    Dim i As Integer
    
    WinHdc& = GetDC(m_Hwnd)
    
    For i = 1 To 3
      Call BitBlt(WinHdc&, 16, (i * 32), Pic.ScaleWidth, Pic.ScaleHeight, Pic.hDC, 0, 0, &HCC0020)
    Next i ' 10-(i*32)
End Sub


Private Sub CleanUp(Pic As PictureBox)
    '
    Dim WinHdc As Long
    Dim i As Integer
    
    WinHdc& = GetDC(m_Hwnd)
    For i = 1 To 3
    Call BitBlt(WinHdc&, 16, (i * 32), Pic.ScaleWidth, Pic.ScaleHeight, Form1.hDC, 0, 0, &HCC0020)
    Next i ' 10-(i*32)
End Sub
