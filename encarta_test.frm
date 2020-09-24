VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin Project1.EncartaFrm EncartaFrm1 
      Height          =   4335
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7646
      BarsBGCOLOR     =   8421504
      BarsFORECOLOROVER=   16777215
      BarsCAPTION     =   "Encarta Frame Style"
      BeginProperty BarsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a control to be like the one that Microsoft Encarta uses to show the description of articles etc."
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'USED TO DETECT MOUSE OUT
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub EncartaFrm1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'WITH THIS CODE WE DETECT WHEN THE MOUSE HAS LEFT THE CONTROL
    With EncartaFrm1
        If Button = 0 Then
            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
                ReleaseCapture 'THE MOUSE IS OUT
                .mouse_out
            Else
                'THE MOUSE IS OVER
                SetCapture .hWnd
                .mouse_movement
            End If
        End If
    End With

End Sub
