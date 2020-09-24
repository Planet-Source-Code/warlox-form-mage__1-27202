VERSION 5.00
Begin VB.UserControl FormMage 
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "FormMage.ctx":0000
   ScaleHeight     =   1245
   ScaleWidth      =   2430
   ToolboxBitmap   =   "FormMage.ctx":0019
   Begin VB.PictureBox picSystray 
      Height          =   495
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Timer timWatchEvents 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   480
      X2              =   480
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   480
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   0
      Picture         =   "FormMage.ctx":032B
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FormMage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Enum values for properties
Public Enum OnTopStatus: otsFormOnTop = 1: otsFormNormal = 0: End Enum
Public Enum CloseCross: ccCrossDisabled = 1: ccCrossEnabled = 0: End Enum
Public Enum FormShape: shpNormal = 0: shpCircle = 1: shpElipse = 2: shpRoundedRect = 3: End Enum
Public Enum SkipOutMode: somFlyRight = 1: somFlyLeft = 2: somFlyUp = 3: somFlyDown = 4: somImplode = 5: somExplode = 6: End Enum

'Default property values
Private Const def_AlwaysOnTop As Integer = 0
Private Const def_XDisabled As Integer = 0
Private Const def_FormShape As Integer = 0
'Constants used for the API functions below
Private Const WM_LBUTTONDBLCLICK = &H203
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const GWL_WNDPRC = -4
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'Variables to store temporary property values
Private m_AlwaysOnTop As OnTopStatus
Private m_XDisabled As CloseCross
Private m_FormShape As FormShape

'Events for the control to raise under specified circumstances
Event FormMoved(ByVal NowX As Single, ByVal NowY As Single, ByVal NowWidth As Single, ByVal NowHeight As Single)
Event SystrayRightClick()
Event SystrayDoubleLeftClick()

'Variables used to monitor events on the parent form
Private OrgWidth As Single
Private OrgHeight As Single
Private ParentTop As Single
Private ParentLeft As Single
Private ParentHeight As Single
Private ParentWidth As Single
'Variables used in the functions of this control
Private Nid As NOTIFYICONDATA

'Local type declarations for the systray icon code
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'API Functions
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Sub DragForm()
' TheForm:  The form you want to start dragging
    
    If Parent.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage Parent.hwnd, &HA1, 2, 0&
    End If
End Sub
Public Function ClearTray()
    On Error GoTo NIDERR
    Shell_NotifyIcon NIM_DELETE, Nid
    On Error GoTo 0
    
NIDERR:
End Function

Private Function DisableX() As Long
Dim hSysMenu As Long
Dim nCnt As Long
    
    Parent.Show
    hSysMenu = GetSystemMenu(Parent.hwnd, False)
    
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar Parent.hwnd
        Else
            SetXStatus = nCnt
        End If
    End If
End Function
Public Property Get AlwaysOnTop() As OnTopStatus
    AlwaysOnTop = m_AlwaysOnTop
End Property

Public Property Let AlwaysOnTop(ByVal NewValue As OnTopStatus)
    m_AlwaysOnTop = NewValue
    
    If m_AlwaysOnTop = otsFormOnTop Then
        KeepOnTop True
    Else
        KeepOnTop False
    End If
    
    PropertyChanged "AlwaysOnTop"
End Property

Public Property Get DisabledX() As CloseCross
    DisabledX = m_XDisabled
End Property

Public Property Let DisabledX(ByVal NewValue As CloseCross)
    If Ambient.UserMode Then
        Err.Raise 382, "Form Mage", "Cannot change Disabled X property at runtime!"
    Else
        m_XDisabled = NewValue
    End If
    
    PropertyChanged "DisabledX"
End Property

Public Function FlashTitlebar() As Long
Dim Succ As Long
    Succ = FlashWindow(Parent.hwnd, 1)
End Function
Private Function KeepOnTop(ByVal TVal As Boolean) As Long
Dim Succ As Long

    If TVal Then
        Succ = SetWindowPos(Parent.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        Succ = SetWindowPos(Parent.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Function
Public Property Get FormShape() As FormShape
    FormShape = m_FormShape
End Property

Public Property Let FormShape(ByVal NewValue As FormShape)
Dim Succ As Long

    m_FormShape = NewValue
    
    If Ambient.UserMode Then
        'Shape the form now
        Succ = ShapeForm
    End If
    
    PropertyChanged "FormShape"
End Property





Public Function SendToTray(ByVal Icon As Picture, ByVal Caption As String) As Long
    Nid.cbSize = Len(Nid)
    Nid.hwnd = picSystray.hwnd
    Nid.uID = 1&
    Nid.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    Nid.uCallbackMessage = WM_MOUSEMOVE
    Nid.hIcon = Icon
    Nid.szTip = Caption & Chr(0)
    
    Shell_NotifyIcon NIM_ADD, Nid
End Function

Private Function ShapeForm() As Long
Dim RetVal As Long
Dim X1 As Single: Dim Y1 As Single
Dim X2 As Single: Dim Y2 As Single
Dim X3 As Single: Dim Y3 As Single
Dim FormX_Pix As Single
Dim FormY_Pix As Single
Dim TmpRgn As Long

    'Get the pixel width and height of the parent form
    FormX_Pix = Parent.Width / Screen.TwipsPerPixelX
    FormY_Pix = Parent.Height / Screen.TwipsPerPixelY
    
    Select Case m_FormShape
        Case 1 'Makes a Circle out of a form
            If Parent.ScaleWidth < Parent.ScaleHeight Then
                X1 = 0
                Y1 = 0
                X2 = FormX_Pix
                Y2 = FormX_Pix
            Else
                X1 = 0
                Y1 = 0
                X2 = FormY_Pix
                Y2 = FormY_Pix
            End If
            TmpRgn = CreateEllipticRgn(X1, Y1, X2, Y2)
        Case 2 'Creates an Elipse from the size of the form
            X1 = 0
            Y1 = 0
            X2 = FormX_Pix
            Y2 = FormY_Pix
            TmpRgn = CreateEllipticRgn(X1, Y1, X2, Y2)
        Case 3 'Rounds the corners of the form
            X1 = 0
            Y1 = 0
            X2 = FormX_Pix
            Y2 = FormY_Pix
            X3 = 24
            Y3 = 24
            TmpRgn = CreateRoundRectRgn(X1, Y1, X2, Y2, X3, Y3)
    End Select
    
    RetVal = SetWindowRgn(Parent.hwnd, TmpRgn, True)
    
    ShapeForm = RetVal
End Function



Public Function SkipOutUnload(ByVal SkipOut As SkipOutMode) As Long

End Function

Private Sub picSystray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static lngMsg As Long
Static blnFlag As Boolean
Dim result As Long
        
    lngMsg = X / Screen.TwipsPerPixelX
    If Not blnFlag Then
        blnFlag = True
        Select Case lngMsg
            Case WM_LBUTTONDBLCLICK
                RaiseEvent SystrayDoubleLeftClick
            Case WM_RBUTTONUP
                RaiseEvent SystrayRightClick
        End Select
        blnFlag = False
    End If
End Sub


Private Sub timWatchEvents_Timer()
    If Ambient.UserMode Then
        If ParentTop = -1024 Then
            ParentTop = Parent.Top
            ParentLeft = Parent.Left
            ParentHeight = Parent.Height
            ParentWidth = Parent.Width
        Else
            If Parent.Top <> ParentTop Or Parent.Left <> ParentLeft Then
                'Form has moved
                RaiseEvent FormMoved(Parent.Left, Parent.Top, Parent.Width, Parent.Height)
                ParentTop = Parent.Top
                ParentLeft = Parent.Left
            End If
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    m_XDisabled = def_XDisabled
    m_AlwaysOnTop = def_AlwaysOnTop
    m_FormShape = def_FormShape
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_XDisabled = PropBag.ReadProperty("DisabledX", def_XDisabled)
    m_FormShape = PropBag.ReadProperty("FormShape", def_FormShape)
    m_AlwaysOnTop = PropBag.ReadProperty("AlwaysOnTop", def_AlwaysOnTop)
    
    On Error GoTo LEAVEFUNC
    If Ambient.UserMode Then
        OrgWidth = Parent.Width
        OrgHeight = Parent.Height
        timWatchEvents.Enabled = True
        If m_XDisabled = ccCrossDisabled Then DisableX
        ShapeForm
        If m_AlwaysOnTop = otsFormOnTop Then
            KeepOnTop True
        Else
            KeepOnTop False
        End If
    End If
    On Error GoTo 0
    
LEAVEFUNC:
End Sub

Private Sub UserControl_Resize()
    Width = 495
    Height = 495
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DisabledX", m_XDisabled, def_XDisabled)
    Call PropBag.WriteProperty("FormShape", m_FormShape, def_FormShape)
    Call PropBag.WriteProperty("AlwaysOnTop", m_AlwaysOnTop, def_AlwaysOnTop)
End Sub


