VERSION 5.00
Begin VB.UserControl EncartaFrm 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label top 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "mY cAPTIOn"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label bottom 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label right 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   135
   End
   Begin VB.Label left 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "EncartaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Bars_BGCOLOR As Long 'BACK COLOR OF THE BORDERS
Private bars_overcolor As Long 'BACK COLOR WHEN THE MOUSE IS OVER
Private top_forecolor As Long 'FORE COLOR FOR THE TEXT IN THE TOP
Private top_forecolorover As Long 'FORE COLOR WHEN MOUSE OVER OCCURS
Private width_bars As Integer ' THE WIDTH OF THE BORDERS

Private display As String 'THE TEXT IN THE TOP
Private overbold As Boolean 'WE WANT THE TEXT TO BE BOLD ON MOUSE OVER?
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
'TO DETECT THE MOUSE OVER....
Private Sub top_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mouse_movement
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    width_bars = 2 'THE START WIDTH OF THE CORNERS
'DEFAULT VALUES OF THE BGCOLORS AND FORECOLORS
    Bars_BGCOLOR = &HC0C0C0
    bars_overcolor = &H0&
    top_forecolor = &HFFFFFF
    top_forecolorover = &HFF0000
    
'THE POSITION OF ALL OF THE 4 CORNERS AND THE COLORS 4 THEM
    top.ForeColor = top_forecolor
    top.left = 0
    top.top = 0
    top.Width = UserControl.Width
    top.Height = 255
    
    
    left.left = 0
    left.top = 0
    left.Width = width_bars * Screen.TwipsPerPixelX
    left.Height = UserControl.Height
    
    right.top = 0
    right.Width = width_bars * Screen.TwipsPerPixelX
    right.left = UserControl.Width - right.Width
    right.Height = UserControl.Height
    right.Visible = False
    
    bottom.left = 0
    bottom.Height = width_bars * Screen.TwipsPerPixelX
    bottom.top = UserControl.Height - bottom.Height
    bottom.Width = UserControl.Width
    
    bottom.BackColor = Bars_BGCOLOR
    left.BackColor = Bars_BGCOLOR
    right.BackColor = Bars_BGCOLOR
    top.BackColor = Bars_BGCOLOR
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
'    mouse_movement
End Sub


Private Sub UserControl_Resize()
'WE NEED TO ADJUST THE BORDERS TO THE NEW SIZE, ALSO WE CHECK 4 THE CAPTION OF THE TOP
    Resize_Bars
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Bars_BGCOLOR = PropBag.ReadProperty("BarsBGCOLOR", &HC0C0C0)
    bars_overcolor = PropBag.ReadProperty("BarsBGCOLOROVER", &H0&)
    top_forecolor = PropBag.ReadProperty("BarsFORECOLOR", &HFFFFFF)
    top_forecolorover = PropBag.ReadProperty("BarsFORECOLOROVER", &HFF0000)
    display = PropBag.ReadProperty("BarsCAPTION", "mY cAPTIOn")
    overbold = PropBag.ReadProperty("BarsFONTBOLD", False)
    top.Alignment = PropBag.ReadProperty("BarsFontAlign", 2)
    Set top.Font = PropBag.ReadProperty("BarsFont", Ambient.Font)
    set_colors
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BarsBGCOLOR", Bars_BGCOLOR, &HC0C0C0)
    Call PropBag.WriteProperty("BarsBGCOLOROVER", bars_overcolor, &H0&)
    Call PropBag.WriteProperty("BarsFORECOLOR", top_forecolor, &HFFFFFF)
    Call PropBag.WriteProperty("BarsFORECOLOROVER", top_forecolorover, &HFF0000)
    Call PropBag.WriteProperty("BarsCAPTION", display, "mY cAPTIOn")
    Call PropBag.WriteProperty("BarsFONTBOLD", overbold, False)
    Call PropBag.WriteProperty("BarsCAPTION", m_BarsCAPTION, m_def_BarsCAPTION)
    Call PropBag.WriteProperty("BarsFontAlign", top.Alignment, 2)
    Call PropBag.WriteProperty("BarsFont", top.Font, Ambient.Font)
    mouse_out
End Sub
'THE PROPERTY STUFF GOES AND STARTS HERE
Public Property Get BarsBGCOLOR() As OLE_COLOR
    BarsBGCOLOR = Bars_BGCOLOR
    set_colors
End Property
Public Property Let BarsBGCOLOR(ByVal New_BackColor As OLE_COLOR)
    Bars_BGCOLOR = New_BackColor
    PropertyChanged "BarsBGCOLOR"
End Property
Public Property Get BarsBGCOLOROVER() As OLE_COLOR
    BarsBGCOLOROVER = bars_overcolor
    set_colors
End Property
Public Property Let BarsBGCOLOROVER(ByVal New_BackColor As OLE_COLOR)
    bars_overcolor = New_BackColor
    PropertyChanged "BarsBGCOLOROVER"
End Property
Public Property Get BarsFORECOLOR() As OLE_COLOR
    BarsFORECOLOR = top_forecolor
    set_colors
End Property
Public Property Let BarsFORECOLOR(ByVal New_BackColor As OLE_COLOR)
    top_forecolor = New_BackColor
    PropertyChanged "BarsFORECOLOR"
End Property
Public Property Get BarsFORECOLOROVER() As OLE_COLOR
    BarsFORECOLOROVER = top_forecolorover
    set_colors
End Property
Public Property Let BarsFORECOLOROVER(ByVal New_BackColor As OLE_COLOR)
    top_forecolorover = New_BackColor
    PropertyChanged "BarsFORECOLOROVER"
End Property
Public Property Get BarsCAPTION() As String
    BarsCAPTION = display
    set_colors
End Property
Public Property Let BarsCAPTION(ByVal Caption As String)
    display = Caption
    PropertyChanged "BarsCAPTION"
End Property
Public Property Get BarsFONTBOLD() As Boolean
    BarsFONTBOLD = overbold
    set_colors
End Property
Public Property Let BarsFONTBOLD(ByVal Caption As Boolean)
    overbold = Caption
    PropertyChanged "BarsFONTBOLD"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=top,top,-1,Font
Public Property Get BarsFont() As Font
    Set BarsFont = top.Font
End Property

Public Property Set BarsFont(ByVal New_Font As Font)
    Set top.Font = New_Font
    PropertyChanged "BarsFont"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=top,top,-1,Alignment
Public Property Get BarsFontAlign() As AlignmentConstants
    BarsFontAlign = top.Alignment
End Property

Public Property Let BarsFontAlign(ByVal New_Alignment As AlignmentConstants)
    top.Alignment() = New_Alignment
    PropertyChanged "BarsFontAlign"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property


Private Sub set_colors()
'WE CHANGE THE COLORS HERE, THE DEFAULTS WHEN THE MOUSE IS NOT OVER
    left.BackColor = Bars_BGCOLOR
    right.BackColor = Bars_BGCOLOR
    top.BackColor = Bars_BGCOLOR
    bottom.BackColor = Bars_BGCOLOR
    top.ForeColor = top_forecolor
    right.Visible = False
    top.Caption = display
    top.FontBold = False
End Sub
Private Sub set_colorsover()
'NOW WE PUT THE MOUSE OVER COLORS IN HERE
    left.BackColor = bars_overcolor
    right.BackColor = bars_overcolor
    top.BackColor = bars_overcolor
    bottom.BackColor = bars_overcolor
    top.ForeColor = top_forecolorover
    top.Caption = display
    top.FontBold = overbold
End Sub
Public Sub mouse_out()
'HERE WE SET THE COLORS AND CHECK IF IT'S ALREADY CHANGED.....
    If top.BackColor = Bars_BGCOLOR Then
        Exit Sub
    End If
    set_colors
    width_bars = 2
    top.Caption = display
    UserControl_Resize
End Sub
Sub mouse_movement()
    If top.BackColor = bars_overcolor Then 'We don't want to change the colors if it's already changed.....
        Exit Sub
    End If
    set_colorsover
    right.Visible = True
    width_bars = 3
    UserControl_Resize
End Sub
Sub Resize_Bars()
    left.Width = width_bars * Screen.TwipsPerPixelX
    right.Width = width_bars * Screen.TwipsPerPixelX
    bottom.Height = width_bars * Screen.TwipsPerPixelX
    
    left.Height = UserControl.Height
    right.Height = UserControl.Height
    right.left = UserControl.Width - right.Width
    bottom.top = UserControl.Height - bottom.Height
    bottom.Width = UserControl.Width
    top.top = 0
    top.Width = UserControl.Width
    top.Caption = display
End Sub


