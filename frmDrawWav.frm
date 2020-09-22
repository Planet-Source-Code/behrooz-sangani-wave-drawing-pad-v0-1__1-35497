VERSION 5.00
Begin VB.Form frmDrawWav 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wave Drawing Pad v0.1"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox HandPic 
      Height          =   615
      Left            =   120
      Picture         =   "frmDrawWav.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Text            =   "100"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox chkDel 
      Caption         =   "Delete temp file on exit"
      Height          =   195
      Left            =   1080
      TabIndex        =   16
      Top             =   5040
      Width           =   2415
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox ColPic 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtOutputPath 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   4680
      Width           =   5415
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Save and Play"
      Height          =   555
      Left            =   1080
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   1080
      ScaleHeight     =   2955
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   555
      Left            =   5880
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "BPP (Byte per pixel)"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(c)2002 Behrooz Sangani <bs20014@yahoo.com>"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   3600
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   3600
      X2              =   0
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label3 
      Caption         =   "Save To:"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X,Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"frmDrawWav.frx":0152
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   6135
   End
End
Attribute VB_Name = "frmDrawWav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Wave Drawing Pad
'  Paint pad that plays your drawings!
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 05/06/2002
'  WebSite: http://www.geocities.com/bs20014/
'  Legal Copyright: Behrooz Sangani © 05/06/2002
'=========================================================================================
'   This is the first version and just a test.
'   I'm planning to make this more sophisticated and use
'   standard frequencies of piano notes or.... But first I
'   have to know if anybody likes such crazy thing or not.
'
'   Please vote and leave comments. It would be greatly appreciated.
'   If you have any other way to calculate sound bytes from colors
'   and mouse positions I'll be glad to know!
'=========================================================================================

Dim ButPressed As Boolean                   'Mouse button pressed
Dim Col As Integer                          'Color that multiplies frequency
Dim DataBytes(0 To 1, 0 To 44099) As Long   'sound bytes
Dim sam As Long
Dim chan As Integer
Private Const Hertz As Single = 2 * 3.14159265 / 44100 'use sin(sam*hertz) to get a tone of 1 hertz

Private Sub cmdPlay_Click()

    MousePointer = vbHourglass
    'Create the temp output file
    WriteWave txtOutputPath.Text, CreateWaveArray(DataBytes, 44100, 16)

    MousePointer = vbDefault
    'Play the sound
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(txtOutputPath.Text, wFlags%)

End Sub


Private Sub cmdRefresh_Click()

    sam = 0
    Picture1.Cls
    Label1.Caption = "X,Y"

End Sub

Private Sub ColPic_Click(Index As Integer)

    'Set the user defined color
    For i = 0 To 7
        ColPic(i).BorderStyle = 0
    Next i
    ColPic(Index).BorderStyle = 1
    Col = Index + 5
    Picture1.ForeColor = ColPic(Index).BackColor

End Sub

Private Sub Form_Load()

    'default values
    sam = 0
    txtOutputPath.Text = App.Path & "\tempwave.wav"
    Col = 10
    Picture1.ForeColor = ColPic(5).BackColor
    ColPic(5).BorderStyle = 1
    Picture1.DrawWidth = 8
    'Somehow needed
    SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Delete temp file if desired
    On Error Resume Next
    If chkDel.Value = 1 Then Kill txtOutputPath.Text
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButPressed = True
    Picture1.MousePointer = 99
    Picture1.MouseIcon = HandPic.Picture
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If mouse button pressed then
    If ButPressed Then

        Label1.Caption = X

        For i = sam To sam + txtTime.Text
            'Don't go further than 44100
            If i < 44100 Then
                'sound byte=(Your chosen color)*(mouse x position)* Sin(mouse y position * our tone)
                'This means that the bottom left of the pad makes low bass tones and
                'the top right of the pad has strong treble tones. The darker color you
                'pick the stronger the tone you would have
                For chan = 0 To 1
                     DataBytes(0, i) = Col * X * Sin(Y * i * Hertz)
                     DataBytes(1, i) = Y * Sin(Col * X * i * Hertz)
                Next chan
                Picture1.PSet (X, Y)
            Else
                'one second of creativity is enough
                Label1.Caption = "End"
            End If

        Next i

        'Do not let the drawing last forever
        'changing txtTime will change the drawing time.
        'I named it BPP(Byte per pixel)!:)
        sam = sam + txtTime.Text
    End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButPressed = False
    Picture1.MousePointer = 0
End Sub
