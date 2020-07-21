VERSION 5.00
Begin VB.Form frmAcercade 
   BackColor       =   &H0076DEC1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de Visual ffmpeg.exe"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0076DEC1&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   240
      Picture         =   "frmAcercade.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   720
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0076DEC1&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAcercade.frx":57E2
      Top             =   120
      Width           =   8415
   End
   Begin Visualffmpeg.ChameleonBtn cmdcarpeta 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   7790273
      BCOLO           =   9037785
      FCOL            =   16384
      FCOLO           =   32768
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAcercade.frx":5C9B
      PICN            =   "frmAcercade.frx":5CB7
      PICH            =   "frmAcercade.frx":6251
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAcercade.frx":67EB
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   2137
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   975
      Left            =   120
      Top             =   2040
      Width           =   8415
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcarpeta_Click()
Unload Me
End Sub

Private Sub Form_Load()
With frmPrograma
 Me.Icon = .Icon
 Me.Caption = "Acerca de " & .Caption
End With
End Sub
