VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "game.frx":0000
   ScaleHeight     =   4920
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   8760
      Picture         =   "game.frx":B3C5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   8760
      Picture         =   "game.frx":BA42
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   615
   End
   Begin VB.HScrollBar HS 
      Height          =   180
      Left            =   6120
      Max             =   150
      Min             =   40
      MouseIcon       =   "game.frx":C022
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2880
      Value           =   40
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   6360
      MouseIcon       =   "game.frx":C32C
      MousePointer    =   99  'Custom
      Picture         =   "game.frx":C636
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   475
      Left            =   8040
      MouseIcon       =   "game.frx":CA78
      MousePointer    =   99  'Custom
      Picture         =   "game.frx":CD82
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Pause Accelaration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   6480
      MouseIcon       =   "game.frx":D1C4
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   360
      MouseIcon       =   "game.frx":D4CE
      MousePointer    =   99  'Custom
      ScaleHeight     =   3915
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   360
      Width           =   5415
      Begin VB.PictureBox blast 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         ScaleHeight     =   735
         ScaleWidth      =   975
         TabIndex        =   10
         Top             =   2640
         Width           =   975
         Begin VB.Image Image1 
            Height          =   735
            Left            =   0
            Picture         =   "game.frx":D620
            Stretch         =   -1  'True
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox img 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4875
         Picture         =   "game.frx":324CE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   3435
         Width           =   480
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MISS :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   390
      Left            =   6120
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HIT :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   390
      Left            =   6240
      TabIndex        =   4
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   7560
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xval As Integer
Dim yval As Integer
Dim hit As Integer, mute As Integer
Dim miss As Integer, inter As Integer
Option Explicit
Private Declare Function PlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim play As Long





Private Sub Command1_Click()
If inter < 1500 Then
inter = inter + 10
End If
HS.Value = inter / 10
Label5.Caption = inter
Timer1.Interval = inter
End Sub

Private Sub Command2_Click()
If inter > 400 Then
inter = inter - 10
End If
HS.Value = inter / 10
Label5.Caption = inter
Timer1.Interval = inter
End Sub

Private Sub Command3_Click()
mute = 0
Command4.ZOrder
End Sub

Private Sub Command4_Click()
mute = 1
Command3.ZOrder
End Sub

Private Sub Command5_Click()
Call Form_Load
End Sub

Private Sub Form_Load()
hit = 0
Command3.ZOrder
mute = 1
miss = 0
inter = 1500
HS.Value = inter / 10
Label5.Caption = inter
img.Visible = False
blast.Visible = False
Label1.Caption = hit
Label2.Caption = miss
End Sub





Private Sub HS_Change()
inter = HS.Value * 10
Timer1.Interval = inter
Label5.Caption = inter
End Sub

Private Sub HS_Scroll()
inter = HS.Value * 10
Timer1.Interval = inter
Label5.Caption = inter
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
hit = hit + 1
Label1.Caption = hit

If Check1.Value = False And inter > 400 Then
inter = inter - 10
HS.Value = inter / 10
End If
Timer1.Interval = inter
Label5.Caption = inter
If mute Then
PlaySound "hitt.wav", 1
End If
img.Visible = False
blast.Top = yval - 100
blast.Left = xval - 200
blast.Visible = True
End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
miss = miss + 1

Label2.Caption = miss
If mute Then
play = PlaySound("miss.wav", 1)
End If
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = True
    Randomize
    img.Visible = True
    
    yval = Int((3435 * ((Rnd * 9999) + 1)) / 10000)
    img.Top = yval
    xval = Int((4835 * ((Rnd * 9999) + 1)) / 10000)
    img.Left = xval
    img.ZOrder
End Sub
