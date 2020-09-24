VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "vV-trygo"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDraw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Draw angle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4890
      TabIndex        =   2
      Top             =   960
      Width           =   1410
   End
   Begin VB.TextBox txtAngle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Angle ~=              °"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   168
      X2              =   256
      Y1              =   240
      Y2              =   192
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'We need some API to paint angle:
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function AngleArc Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Single, ByVal eSweepAngle As Single) As Long

Private Const AngR As Long = 20 'Radius of angle which we draw
Dim Angle As Single

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Change line coords:
    Line1.X2 = X
    Line1.Y2 = Y
    
    'Calculate angle and show it
    Angle = Rad2Deg(AngleBetween(Line1.X1, Line1.Y1, CLng(X), CLng(Y)))
    txtAngle.Text = Round(Angle, 1)
    
    'Drawing angle
    Cls
    If chkDraw.Value Then
        MoveToEx Me.hdc, Line1.X1 + AngR, Line1.Y1, ByVal 0&
        AngleArc Me.hdc, Line1.X1, Line1.Y1, AngR, 0, Angle
    End If
    
End Sub
