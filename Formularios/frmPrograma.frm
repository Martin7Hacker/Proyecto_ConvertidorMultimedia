VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmPrograma 
   BackColor       =   &H0076DEC1&
   Caption         =   "Visual ffmpeg v2017"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12585
   Icon            =   "frmPrograma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12585
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   5760
      Top             =   5160
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   17
      BmpCount        =   17
      CheckBorderColor=   9037785
      SelMenuBorder   =   9037785
      SelMenuBackColor=   12382177
      SelMenuForeColor=   0
      SelCheckBackColor=   8454143
      MenuBorderColor =   9037785
      SeparatorColor  =   -2147483632
      MenuBackColor   =   9037785
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   9037785
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   9037785
      MenuPopupBackColor=   9037785
      ShortCutNormalColor=   9037785
      ShortCutSelectColor=   9037785
      ArrowNormalColor=   9037785
      ArrowSelectColor=   9037785
      ShadowColor     =   9037785
      Bmp:1           =   "frmPrograma.frx":0CCA
      Key:1           =   "#Drivers"
      Bmp:2           =   "frmPrograma.frx":2A32
      Key:2           =   "#Facebook"
      Bmp:3           =   "frmPrograma.frx":379A
      Key:3           =   "#Twitter"
      Bmp:4           =   "frmPrograma.frx":4502
      Key:4           =   "#Instagram"
      Bmp:5           =   "frmPrograma.frx":526A
      Key:5           =   "#Youtube"
      Bmp:6           =   "frmPrograma.frx":5FD2
      Key:6           =   "#VisaulizarMotor"
      Bmp:7           =   "frmPrograma.frx":6D3A
      Key:7           =   "#CargarArchivos"
      Bmp:8           =   "frmPrograma.frx":7162
      Key:8           =   "#CambiarDirectorio"
      Bmp:9           =   "frmPrograma.frx":758A
      Key:9           =   "#EliminarSeleciónado"
      Bmp:10          =   "frmPrograma.frx":79B2
      Key:10          =   "#eliminarTodo"
      Bmp:11          =   "frmPrograma.frx":7DDA
      Key:11          =   "#VerSoloGraficosparaWindows"
      Bmp:12          =   "frmPrograma.frx":8202
      Key:12          =   "#VerGraficosyConsola"
      Bmp:13          =   "frmPrograma.frx":862A
      Key:13          =   "#ArchivoDecodificar"
      Bmp:14          =   "frmPrograma.frx":8A52
      Key:14          =   "#Detener"
      Bmp:15          =   "frmPrograma.frx":8E7A
      Key:15          =   "#AcualizarSoftware"
      Bmp:16          =   "frmPrograma.frx":92A2
      Key:16          =   "#AcercadeVisualffmpeg"
      Bmp:17          =   "frmPrograma.frx":96CA
      Key:17          =   "#DonacinesparasegirconlosProyectos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   7785
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2999
            MinWidth        =   2999
            Picture         =   "frmPrograma.frx":9AF2
            Text            =   ""
            TextSave        =   "08/06/2017"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Fecha Actual del Sistema"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   2999
            MinWidth        =   2999
            Picture         =   "frmPrograma.frx":A08C
            Text            =   ""
            TextSave        =   "10:23"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Hora Actual del Sistema"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Numerico"
            TextSave        =   "Numerico"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4092
            MinWidth        =   4092
            Picture         =   "frmPrograma.frx":A626
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Visualffmpeg.ChameleonBtn cmdCargarArchivos 
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Cargar Archivos"
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
      MICON           =   "frmPrograma.frx":CB38
      PICN            =   "frmPrograma.frx":CB54
      PICH            =   "frmPrograma.frx":D0EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11040
      Top             =   6960
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2040
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   94
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E4FBFC&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   10905
      TabIndex        =   11
      Top             =   960
      Width           =   10935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   7695
      End
      Begin VB.Label lblselector 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   690
      End
   End
   Begin VB.ComboBox cobLista 
      BackColor       =   &H00B1EFE6&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   70
      Width           =   6855
   End
   Begin VB.PictureBox picPrograma 
      AutoSize        =   -1  'True
      BackColor       =   &H00B1EFE6&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   11280
      Picture         =   "frmPrograma.frx":D688
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   8
      ToolTipText     =   "Motor de Decodificación ffmpeg.exe"
      Top             =   960
      Width           =   720
   End
   Begin ComctlLib.ListView lv 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2778
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   16384
      BackColor       =   11661286
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtOutPut 
      BackColor       =   &H00B1EFE6&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6120
      Width           =   9495
   End
   Begin VB.TextBox txtCL 
      BackColor       =   &H00B1EFE6&
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   6960
      Width           =   9495
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11640
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   16384
      BackColor       =   11661286
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin Visualffmpeg.ChameleonBtn eliminarselecionado 
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Eliminar Seleccionado"
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
      MICON           =   "frmPrograma.frx":12E6A
      PICN            =   "frmPrograma.frx":12E86
      PICH            =   "frmPrograma.frx":13420
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdeleimartodo 
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Eliminar Todo"
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
      MICON           =   "frmPrograma.frx":139BA
      PICN            =   "frmPrograma.frx":139D6
      PICH            =   "frmPrograma.frx":13F70
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdgraficos 
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Ver Solo Graficos Para Windows"
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
      MICON           =   "frmPrograma.frx":1450A
      PICN            =   "frmPrograma.frx":14526
      PICH            =   "frmPrograma.frx":14AC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdConsola 
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Ver Graficos y Consola de MS-DOS"
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
      MICON           =   "frmPrograma.frx":1505A
      PICN            =   "frmPrograma.frx":15076
      PICH            =   "frmPrograma.frx":15610
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdConvert 
      Height          =   495
      Left            =   11040
      TabIndex        =   22
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Convertir "
      ENAB            =   0   'False
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
      MICON           =   "frmPrograma.frx":15BAA
      PICN            =   "frmPrograma.frx":15BC6
      PICH            =   "frmPrograma.frx":16160
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdAcercade 
      Height          =   375
      Left            =   12030
      TabIndex        =   23
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmPrograma.frx":166FA
      PICN            =   "frmPrograma.frx":16716
      PICH            =   "frmPrograma.frx":16CB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdDirectorio 
      Height          =   375
      Left            =   9720
      TabIndex        =   24
      Top             =   6550
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Directorio"
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
      MICON           =   "frmPrograma.frx":1724A
      PICN            =   "frmPrograma.frx":17266
      PICH            =   "frmPrograma.frx":17800
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdDetener 
      Height          =   495
      Left            =   11040
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Detener"
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
      MICON           =   "frmPrograma.frx":17D9A
      PICN            =   "frmPrograma.frx":17DB6
      PICH            =   "frmPrograma.frx":18350
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdcarpeta 
      Height          =   375
      Left            =   9720
      TabIndex        =   31
      Top             =   6170
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Abrir Carpeta "
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
      MICON           =   "frmPrograma.frx":188EA
      PICN            =   "frmPrograma.frx":18906
      PICH            =   "frmPrograma.frx":18EA0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdAcualizar 
      Height          =   375
      Left            =   9720
      TabIndex        =   32
      Top             =   5770
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Actualizar"
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
      MICON           =   "frmPrograma.frx":1943A
      PICN            =   "frmPrograma.frx":19456
      PICH            =   "frmPrograma.frx":199F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdDonar 
      Height          =   375
      Left            =   11400
      TabIndex        =   33
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmPrograma.frx":19F8A
      PICN            =   "frmPrograma.frx":19FA6
      PICH            =   "frmPrograma.frx":1A540
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdMiRegistroDeFacebook 
      Height          =   375
      Left            =   10800
      TabIndex        =   34
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmPrograma.frx":1AADA
      PICN            =   "frmPrograma.frx":1AAF6
      PICH            =   "frmPrograma.frx":1B7D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Visualffmpeg.ChameleonBtn cmdWindows 
      Height          =   375
      Left            =   10250
      TabIndex        =   35
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmPrograma.frx":1C4AA
      PICN            =   "frmPrograma.frx":1C4C6
      PICH            =   "frmPrograma.frx":21CB8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      Height          =   7815
      Left            =   20
      Top             =   15
      Width           =   12615
   End
   Begin VB.Label lblTiempo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1920
      TabIndex        =   28
      Top             =   5160
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tiempo Trascurrido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   1710
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Formato de Salida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Formato de Entrada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Descodificar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   5640
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Visualizar Etapas de Comandos MS-DOS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   3510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Directorio de Salida de Archivos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   2820
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar archivos a convertir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar formato de codificación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu opciones 
      Caption         =   "opciones"
      Visible         =   0   'False
      Begin VB.Menu esp1 
         Caption         =   "-"
      End
      Begin VB.Menu Drivers 
         Caption         =   "&Drivers / Preset"
      End
      Begin VB.Menu esp2 
         Caption         =   "-"
      End
      Begin VB.Menu Facebook 
         Caption         =   "&Facebook"
      End
      Begin VB.Menu Twitter 
         Caption         =   "&Twitter"
      End
      Begin VB.Menu Instagram 
         Caption         =   "&Instagram"
      End
      Begin VB.Menu Youtube 
         Caption         =   "&Youtube"
      End
      Begin VB.Menu esp3 
         Caption         =   "-"
      End
      Begin VB.Menu VisaulizarMotor 
         Caption         =   "&Visaulizar Motor de Converción"
      End
      Begin VB.Menu ESP4 
         Caption         =   "-"
      End
      Begin VB.Menu CargarArchivos 
         Caption         =   "&Cargar Archivos"
      End
      Begin VB.Menu CambiarDirectorio 
         Caption         =   "&Cambiar Directorio"
      End
      Begin VB.Menu esp5 
         Caption         =   "-"
      End
      Begin VB.Menu EliminarSeleciónado 
         Caption         =   "&Eliminar Seleciónado"
      End
      Begin VB.Menu eliminarTodo 
         Caption         =   "&Eliminar Todo "
      End
      Begin VB.Menu esp6 
         Caption         =   "-"
      End
      Begin VB.Menu VerSoloGraficosparaWindows 
         Caption         =   "&Ver Solo Graficos para Windows"
      End
      Begin VB.Menu VerGraficosyConsola 
         Caption         =   "&Ver Graficos y Consola de MS - DOS"
      End
      Begin VB.Menu ArchivoDecodificar 
         Caption         =   "&Convertir "
      End
      Begin VB.Menu Detener 
         Caption         =   "&Detener"
      End
      Begin VB.Menu AcualizarSoftware 
         Caption         =   "&Acualizar Software"
      End
      Begin VB.Menu esp7 
         Caption         =   "-"
      End
      Begin VB.Menu DonacinesparasegirconlosProyectos 
         Caption         =   "&Donaciones para seguir con los Proyectos"
      End
      Begin VB.Menu AcercadeVisualffmpeg 
         Caption         =   "&Acerca de Visual ffmpeg v2017"
      End
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'* Convertidor de Archivos Digitales utilizando el programa Opensurce *
'* ffmpeg autor: Martin Grasso.                                       *
'**********************************************************************
Option Explicit

' Colección para guardar los archivos
Dim mColFiles           As New Collection
Dim formatos            As New Collection
Dim sInput              As New Collection
Dim sOutput             As New Collection
Dim datos               As New Collection
Dim sParam              As String
Dim ret                 As String
Dim mensajes             As String
Dim iconos               As Boolean
Dim Lista               As Boolean
Dim contador As Long
Dim tiempoTrascurrido As String
Private graficosWindows As Boolean
' --------------------------------------------------------------------------------
' \\ -- Declaraciones
' --------------------------------------------------------------------------------
Private Declare Function AbrirWeb Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal _
cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, _
lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" _
(ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
Private cPresetFFMPg As cFFPresetMPG
'Declaración del Api GetFileTitle
Private Declare Function GetFileTitle _
    Lib "comdlg32.dll" _
    Alias "GetFileTitleA" ( _
        ByVal lpszFile As String, _
        ByVal lpszTitle As String, _
        ByVal cbBuf As Integer) As Integer

' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String

On Local Error GoTo errFunction
    
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
    
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
    
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path

Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString

End Function

Private Function Obtener_Nombre_Archivo(p As String)
  Dim Buffer As String
    'Buffer de caracteres
    Buffer = String(255, 0)
    'Llamada a GetFileTitle, pasandole el path, el buffer y el tamaño
    GetFileTitle p, Buffer, Len(Buffer)
    'Retornamos el nombre eliminando los espacios nulos
    Obtener_Nombre_Archivo = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
End Function

Private Sub AcercadeVisualffmpeg_Click()
cmdAcercade_Click
End Sub

Private Sub AcualizarSoftware_Click()
cmdAcualizar_Click
End Sub

Private Sub ArchivoDecodificar_Click()
cmdConvert_Click
End Sub

Private Sub CambiarDirectorio_Click()
cmdDirectorio_Click
End Sub

Private Sub CargarArchivos_Click()
cmdCargarArchivos_Click
End Sub

Private Sub cmdAcercade_Click()
On Error GoTo nose:
frmAcercade.Show 1
nose:
End Sub

Private Sub cmdAcercade_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdAcualizar_Click()
On Error GoTo nose:
Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "http://visualconvertidor.blogspot.com.uy/p/actualizar.html", _
 &O0, &O0, 0)
nose:
End Sub

Private Sub cmdAcualizar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdCargarArchivos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdcarpeta_Click()
On Error GoTo nose:
 If ret = "" Then
  Call ShellExecute(Me.hwnd, "Open", App.Path, &O0, &O0, 1)
 ElseIf Not (ret = "") Then
  Call ShellExecute(Me.hwnd, "Open", ret, &O0, &O0, 1)
 End If
nose:
End Sub

Private Sub cmdcarpeta_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdConsola_Click()
On Error GoTo nose:
graficosWindows = False
GraficoConsola True, False
nose:
End Sub

Private Sub GraficoConsola(ByVal control As Boolean, ByVal control2 As Boolean)
cmdgraficos.Enabled = control
cmdConsola.Enabled = control2
End Sub

Private Sub cmdConsola_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdConvert_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdDetener_Click()
On Error GoTo nose:
If Timer1.Enabled = True Then
If MsgBox("¿Queres Detener todos los Archivos del convertidor?" _
  , vbExclamation + vbYesNo) = vbYes Then
   iconos = True
   mensajes = "Se Detubo"
   cmdDetener.Visible = False
   cmdConvert.Visible = True
   CierraProceso "ffmpeg.exe", False
 End If
End If
nose:
End Sub

Private Sub cmdDetener_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdDirectorio_Click()
 On Error GoTo nose:
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta("Seleccioné una carpeta para guardar sus Archivos" & _
    "Digitales(Los Nombres de los archivos serán leídos de los títulos de los Archivos Digitales. ")
    guardarrutaAlmacenamiento
    txtOutPut.Text = ret
nose:
End Sub

Private Sub cmdDirectorio_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdDonar_Click()
On Error GoTo nose:
Dim x As String
 x = AbrirWeb(Me.hwnd, "Open" _
 , "http://martinsoft0.blogspot.com.uy/p/donar.html", _
 &O0, &O0, 0)
nose:
End Sub

Private Sub cmdDonar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdeleimartodo_Click()
On Error GoTo nose:
If ListView1.ListItems.Count >= 1 Then
  If MsgBox("¿Queres eliminar todos los Archivos del convertidor?" _
  , vbExclamation + vbYesNo) = vbYes Then
     ListView1.ListItems.Clear
     Set mColFiles = Nothing
     cmdConvert.Enabled = False
  End If
End If
nose:
End Sub

Private Sub cmdeleimartodo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdgraficos_Click()
On Error GoTo nose:
graficosWindows = True
GraficoConsola False, True
nose:
End Sub

Private Sub cmdgraficos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdMiRegistroDeFacebook_Click()
On Error GoTo nose:
Dim x As String
 x = AbrirWeb(Me.hwnd, "Open" _
 , "https://www.facebook.com/hacker.martin0", _
 &O0, &O0, 0)
nose:
End Sub

Private Sub cmdMiRegistroDeFacebook_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cmdWindows_Click()
On Error GoTo nose:
Dim x As String
 x = AbrirWeb(Me.hwnd, "Open" _
 , "https://ffmpeg.org/download.html", _
 &O0, &O0, 0)
nose:
End Sub

Private Sub cmdWindows_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub cobLista_Change()
panelActivo
End Sub
Private Sub cobLista_Click()
 lblselector.Caption = cobLista.Text
 panelActivo
 ProgressBar1.Value = cobLista.ListIndex
 lblDesc.Caption = cobLista.Text
 
End Sub

Private Sub cobLista_GotFocus()
lv.SelectedItem.Selected = False
Lista = True
End Sub

Private Sub cmdCargarArchivos_Click()
On Error GoTo nose:
CierraProceso "ffmpeg.exe", False
Dim lvItem As ListItem
 On Local Error GoTo error_handler
      
    ' Configurar el cuadro de diálogo
    ' ---------------------------------------------------------
    With cd
        ' Limpiar la propiedad FileName
        .FileName = vbNullString
        ' Establecer Flag para poder seleccionar múltiples archivos desde el cd
        .Flags = .Flags Or cdlOFNExplorer Or cdlOFNAllowMultiselect
        ' Tamaño de Buffer para el FileName
        .MaxFileSize = 32767 ' <- máximo 32 K
        ' Establecer filtro
        .Filter = "Todos los Archivos|*.*"
        ' Abrir
        .ShowOpen
        ' Verificar que el FileName no sea una cadena vacía
        If .FileName <> vbNullString Then
           ' Array para obtener las rutas
           Dim arrPaths() As String
           arrPaths = Split(.FileName, Chr(0))
           ' Enviar array de archivos para agregar a la colección
           Call mAddFiles(arrPaths)
           'Set arrPaths(0) = Nothing
           Erase arrPaths
           ActivoBotonConvertir
           ' Actualizar listado
           Call mUpdateList(ListView1)
           ReDim arrPaths(0)
        End If
          .FileName = vbNullString
         End With
         
     ' Error
    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical
nose:
End Sub

Private Sub ActivoBotonConvertir()
  If ListView1.ListItems.Count >= 0 Then
     cmdConvert.Enabled = True
     BotonEliminarActivo True
  End If
End Sub

Private Sub Detener_Click()
cmdDetener_Click
End Sub

Private Sub DonacinesparasegirconlosProyectos_Click()
cmdDonar_Click
End Sub

Private Sub Drivers_Click()
cmdWindows_Click
End Sub

Private Sub eliminarselecionado_Click()
 On Error GoTo nose
 If ListView1.ListItems.Count > 0 Then
   If ListView1.ListItems.Count = 1 Then
       Set mColFiles = Nothing
       Else
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
        mColFiles.Remove ListView1.SelectedItem.Index
    End If
    Call mUpdateList(ListView1)
  End If
nose:
End Sub

Private Sub EliminarSeleciónado_Click()
eliminarselecionado_Click
End Sub

Private Sub eliminarselecionado_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub eliminarTodo_Click()
cmdeleimartodo_Click
End Sub

Private Sub Facebook_Click()
cmdMiRegistroDeFacebook_Click
End Sub

Private Sub Form_Initialize()
    Call SetErrorMode(2)
    Call InitCommonControls
End Sub

Private Sub Form_Load()
  CierraProceso "ffmpeg.exe", False
  cargarRutaAlmacenamiento
  If ret = "" Then
     guardarrutaAlmacenamiento
  End If
  If ret = "" Then
    cmdDirectorio_Click
    guardarrutaAlmacenamiento
  End If
 cargarPrimeraCancionConRutaEspesifica False
    ' Crear nueva colección para guardar los archivos
    Set mColFiles = New Collection
    panelActivo
    Dim lvItem As ListItem
    Dim i      As Long
    Set cPresetFFMPg = New cFFPresetMPG
    ' -- setear listview
    With lv
        .View = lvwReport
        .ColumnHeaders.Add , , "Formatos ->", 1500
        .ColumnHeaders.Add , , "Categorias ->"
        .ColumnHeaders.Add , , "Descripciónes ->", 2777
        .ColumnHeaders.Add , , "Comandos de ffmpeg.exe ->", 2777
    End With
    'para los archivos a convertir
    With ListView1
        .View = lvwReport
        .ColumnHeaders.Add , , "ID | Directorio | Entrada * ->", 12000
    End With
    ' -- Recorrer todos los presets y cargarlos en el LV
    With cPresetFFMPg
        For i = 0 To .PresetsCount
            Call .setPreset(i)
            Set lvItem = lv.ListItems.Add(, , .PresetExtension)
            lvItem.SubItems(1) = .PresetCategory
            lvItem.SubItems(2) = .PresetDescription
            lvItem.SubItems(3) = .PresetParameters '-- linea de comandos
            cobLista.AddItem .PresetExtension & "  --->  " & .PresetDescription '.PresetDescription & .PresetParameters
            formatos.Add .PresetExtension
        'i + 1 & ")        " &
        Next
        Call lv_ItemClick(lv.ListItems(1))
    End With
    cmdgraficos.Enabled = False
    graficosWindows = True
    ToolTipText
    cmdConsola_Click ' modo MS - DOS Vision.
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Timer1.Enabled = True Then
     Cancel = 1
       If MsgBox("¿Quieres Detener la Decodificación de los Archivos en la Bandeja y Cerrar la Aplicación?", _
       vbExclamation + vbYesNo) = vbYes Then
          If Not cPresetFFMPg Is Nothing Then Set cPresetFFMPg = Nothing
          CierraProceso "ffmpeg.exe", False
          Cancel = 0
          Unload Me
       End If
      Else
      Cancel = 0
    End If
End Sub

Private Sub Instagram_Click()
Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "https://www.instagram.com/hacker.martin/", _
 &O0, &O0, 0)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub lblTiempo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
PopupMenu opciones
End If
End Sub

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
On Error GoTo nose:
    lblDesc.Caption = Item.SubItems(2) & " -- Formato: " & Item.Text
    panelActivo
    Lista = False
    
    'lv.SelectedItem.Selected = True
   cobLista.ListIndex = lv.SelectedItem.Index - 1
   ProgressBar1.Value = lv.SelectedItem.Index
nose:
End Sub

Private Sub panelActivo()
If Lista = True Then
lblselector.Visible = True
picFrame.Visible = True
ElseIf Lista = False Then
lblselector.Visible = False
picFrame.Visible = False
End If
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub picFrame_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub picPrograma_Click()
On Error GoTo nose:
Shell "ffmpeg.exe", vbNormalFocus
nose:
End Sub

Private Sub mAddFiles(arrFiles() As String)
On Local Error GoTo error_handler
      With mColFiles
        ' Si el array tiene un solo elemento, es por que se seleccionó un solo fichero
        '( Es decir Contiene la ruta completa : Dir + FileName)
        If UBound(arrFiles) = 0 Then
            ' Comprobar que la colección tiene elementos ...
            If .Count > 0 Then
                Call .Add(arrFiles(0), arrFiles(0), 1) ' agregar item en el primer lugar
            ' si no hay elementos ...
            Else
                 Call .Add(arrFiles(0), arrFiles(0))
            End If
              
        ' Si no, Hay mas de un archivo ....
        Else
            ' El primer elemento del array es el directorio ( Guardar el path en la variable  )
            Dim sDir As String
            sDir = arrFiles(0)
              
            ' verificar el separador de path
            If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
            ' Los archivos ( solo el nombre sin el path )
            Dim i As Integer
            For i = 1 To UBound(arrFiles)
               ' REcorrer el array y agregarlos a la colección
               If .Count > 0 Then
                   Call .Add(sDir & arrFiles(i), sDir & arrFiles(i), 1) 'agregar primero
               Else
                   Call .Add(sDir & arrFiles(i), sDir & arrFiles(i))
               End If
             Next
        End If
    End With
Exit Sub
error_handler:
If Err.Number = 457 Then
    Resume Next ' ignorar error cuando se agrega el mismo archivo
Else
    'MsgBox Err.Description
End If
End Sub

Private Sub mUpdateList(lBox As ListView)
Dim contador As Integer
    With lBox
        ' limpiar listbox y volver a cargar
        .ListItems.Clear
        Dim xItem As Variant
        ' recorrer items de la colección
        
        For Each xItem In mColFiles
            contador = contador + 1
            .ListItems.Add , , contador & ") " & CStr(xItem)
        Next
        ' seleccionar el primero
        
    End With
   
End Sub

Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
    Const MAX_PATH As Long = 260
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    NombreDelProceso = UCase$(NombreDelProceso)
    ReDim lProcesses(1023) As Long
   If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                    If Len(sName) = Len(NombreDelProceso) Then
                        If NombreDelProceso = UCase$(sName) Then EstaCorriendo = True: Exit Function
                    End If
                End If
            End If
            CloseHandle hProcess
        Next N
    End If
End Function

Private Sub picPrograma_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub ProgressBar1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub ProgressBar2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub StatusBar1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub Timer1_Timer()
cronometro
Dim i As Integer
If EstaCorriendo("ffmpeg.exe") Then
        ProgressBar2.Value = ProgressBar2.Value + 1
        If ProgressBar2.Value = 100 Then
           ProgressBar2.Value = 0
        End If
    Else
        Timer1.Enabled = False
         Set datos = Nothing
         Set mColFiles = Nothing
         Set formatos = Nothing
         Set sInput = Nothing
         Set sOutput = Nothing
         Set datos = Nothing
        ListView1.ListItems.Clear
        If iconos = False Then
         mensajes = "La Codificación a Terminado"
         iconos = False
        End If
        mensaje mensajes, iconos
        BotonEliminarActivo False
        cmdConvert.Enabled = False
        ProgressBar2.Value = 0
        txtCL.Text = ""
        lblTiempo = ""
        tiempoTrascurrido = ""
        contador = 0
        cmdDetener.Visible = False
        cmdConvert.Visible = True
    End If
   iconos = False
End Sub

Private Sub mensaje(ByVal mensaje As String, ByVal icono As Boolean)
Select Case icono
       Case (False)
       MsgBox mensajes, vbInformation, Me.Caption
       Case (True)
       MsgBox mensajes, vbExclamation, Me.Caption
End Select
End Sub

Public Sub BotonEliminarActivo(ByVal Activo As Boolean)
 cmdeleimartodo.Enabled = Activo
 eliminarselecionado.Enabled = Activo
End Sub

' Cerrar los procesos de Windows
Function CierraProceso(StrNombreProceso As String, Optional DecirSINO As Boolean = True) As Boolean
  Dim ListaProcesos  As Object
  Dim ObjetoWMI    As Object
  Dim ProcesoConcreto    As Object
  CierraProceso = False
  Set ObjetoWMI = GetObject("winmgmts:")
  If IsNull(ObjetoWMI) = False Then
  Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
  For Each ProcesoConcreto In ListaProcesos
    If UCase(ProcesoConcreto.Name) = UCase(StrNombreProceso) Then
        If DecirSINO Then
          If MsgBox("¿Detener Decodificación ? " & _
               ProcesoConcreto.Name & vbNewLine & _
               "...¿Está seguro?", _
               vbYesNo + vbCritical) _
               = vbYes Then
           ProcesoConcreto.Terminate (0)
           CierraProceso = True
          End If
        Else
         ProcesoConcreto.Terminate (0)
         CierraProceso = True
        End If
     End If
    Next
  Else
  'pon aqui un msgbox con el error que se produzca
  End If
  Set ListaProcesos = Nothing
  Set ObjetoWMI = Nothing
End Function

Private Sub cargarPrimeraCancionConRutaEspesifica(ByVal Activo As Boolean)
Select Case Activo
       Case True
  If ret = "" Then
     txtOutPut = App.Path & "\" & Date & Obtener_Nombre_Archivo _
    (ListView1.ListItems(1).Text) & "." & lv.SelectedItem.Text
  Else
    txtOutPut = ret & "\" & Date & Obtener_Nombre_Archivo _
   (ListView1.ListItems(1).Text) & "." & lv.SelectedItem.Text
 End If
     Case False
    If ret = "" Then
    txtOutPut = App.Path & "\"
  Else
    txtOutPut = ret & "\"
 End If
End Select
End Sub

Private Sub cmdConvert_Click()
cmdConvert.Visible = False
cmdDetener.Visible = True
If ret = "" Then
     txtOutPut = App.Path & "\" & Obtener_Nombre_Archivo _
    (ListView1.ListItems(1).Text) & "." & lv.SelectedItem.Text
  Else
    txtOutPut = ret & "\" & Obtener_Nombre_Archivo _
   (ListView1.ListItems(1).Text) & "." & lv.SelectedItem.Text
End If
 
  If (ListView1.ListItems.Count) = 1 Then
    ' cuand es igual a 1
     convertirSoloUnaPista
     Timer1.Enabled = True
  ElseIf (ListView1.ListItems.Count) > 1 Then
              ' cuando es mayor a 1
     convertirMayoresA1Pista
     Timer1.Enabled = True
 End If
End Sub

Private Sub convertirSoloUnaPista()
    Dim sInput      As String
    Dim sParam      As String
    Dim sOutput     As String
    'carga solo 1 pista
    txtOutPut = App.Path & "\" & Obtener_Nombre_Archivo _
    (ListView1.ListItems(1).Text) & "." & lv.SelectedItem.Text
     ' -- setear los parámetros
    sInput = App.Path & "\ffmpeg.exe " & "-i " & Chr(34) & mColFiles.Item(1) & Chr(34)
    sParam = " " & lv.SelectedItem.SubItems(3) & " "
    GuardarEnDirectorio 1 ' solo va a existir un archivo en ese directorio
    sOutput = " " & Chr(34) & txtOutPut.Text & Chr(34)
    ' -- Ejecutar la línea de comandos con función shell
    txtCL.Text = sInput & sParam & sOutput
    graficosWindowsSINO
End Sub

Private Sub convertirMayoresA1Pista()
    Dim sInput As New Collection
    Dim sParam        As String
    Dim sOutput       As New Collection
    Dim datos         As New Collection
    Dim rec As Integer
    ' -- setear los parámetros
    For rec = 1 To ListView1.ListItems.Count
     sInput.Add App.Path & "\ffmpeg.exe " & "-i " & Chr(34) & mColFiles.Item(rec) & Chr(34)
    Next rec
    sParam = " " & lv.SelectedItem.SubItems(3) & " "
    For rec = 1 To ListView1.ListItems.Count
    txtOutPut.Text = App.Path & "\" & Obtener_Nombre_Archivo(ListView1.ListItems(rec).Text) & _
    "." & lv.SelectedItem.Text
    GuardarEnDirectorio rec ' multiples Archivos
    sOutput.Add " " & Chr(34) & txtOutPut.Text & Chr(34)
    Next rec
    For rec = 1 To ListView1.ListItems.Count
    datos.Add sInput(rec) & sParam & sOutput(rec)
     txtCL.Text = sInput.Item(rec) & sParam & sOutput.Item(rec)
    Next rec
    ' -- Ejecutar la línea de comandos con función shell
    For rec = 1 To ListView1.ListItems.Count
    graficosWindowsSINO
    txtCL.Text = datos(rec)
    Next rec
End Sub

Private Sub graficosWindowsSINO()
 If graficosWindows = True Then
    If Len(txtCL.Text) Then Shell txtCL.Text, vbHide
 ElseIf graficosWindows = False Then
    If Len(txtCL.Text) Then Shell txtCL.Text, vbNormalFocus
 End If
End Sub

Private Sub cronometro()
 contador = contador + 1
    tiempoTrascurrido = Format(Int(contador / 36000) Mod 24, "00") & ":" & _
                        Format(Int(contador / 600) Mod 60, "00") & ":" & _
                        Format(Int(contador / 10) Mod 60, "00") & ":" & _
                        Format(contador Mod 99, "00")
    lblTiempo.Caption = tiempoTrascurrido
End Sub

Private Sub GuardarEnDirectorio(ByVal rec As Integer)
 If ret = "" Then
    txtOutPut.Text = App.Path & "\" & Obtener_Nombre_Archivo _
    (ListView1.ListItems(rec).Text) & "." & lv.SelectedItem.Text
    Else
    txtOutPut.Text = ret & "\" & Obtener_Nombre_Archivo _
    (ListView1.ListItems(rec).Text) & "." & lv.SelectedItem.Text
 End If
End Sub

' Abrir Ruta Definida Anterior Mente por el Usuario
Private Sub cargarRutaAlmacenamiento()
On Error GoTo nose
Dim cargar As String
Open "Directorio.ini" For Input As 1
 Do While Not EOF(1)
  Line Input #1, ret ' ruta de almacenamiento
 Loop
 Close #1
nose:
End Sub

' Guardar el Archivo en el Directorio
Private Sub guardarrutaAlmacenamiento()
On Error GoTo nose
Open "Directorio.ini" For Output As 1
 Print #1, ret ' ruta de almacenamiento
Close #1
nose:
End Sub

Private Sub ToolTipText()
'bóton Donación
cmdDonar.ToolTipText = "Oprima Aquí para Gestionar una Donación Saludable."
'bóton Acerca de
cmdAcercade.ToolTipText = "Acerca de Virtual ffmpeg v2017."
'Listview
lv.ToolTipText = "Planilla de Formatos de Conversiones Digitales."
'Listview1
ListView1.ToolTipText = "Bandeja de Entrada de Archivos Digitales."
'barra de formato selectiva
ProgressBar1.ToolTipText = "Barra de Progreso  selectiva en Porcentaje Máximo 100%"
'barra de espera
ProgressBar2.ToolTipText = "Barra de Espera..."
'bóton de Cargar Archivos
cmdCargarArchivos.ToolTipText = "Te permite Cargar 1 o más archivos en Bandeja."
'bóton Eliminar Seleciónado
eliminarselecionado.ToolTipText = "Te permite Eliminar 1 o más elementos de la bandeja de entrada."
'bóton eliminar todo
cmdeleimartodo.ToolTipText = "Elimina Todos los Elementos de la Bandeja de Entrada"
'bóton ver solo graficos para windows
cmdgraficos.ToolTipText = "Te permite solo ver la o las Ventanas y elementos Visuales en el sistema Operativo Windows."
'Bóton ver vewntanas de comandos de ms-dos
cmdConsola.ToolTipText = "Te permite ver  las ventanas y elementos Visuales y también las Ventanas de Comandos de MS-DOS."
'Bóton Decodificar
cmdConvert.ToolTipText = "Te permite decodificar el Archivo o los Archivos de Entrada a el Formato indicado por el usuario."
'Bóton Detener
cmdDetener.ToolTipText = "Te permite detener la decodificación."
'Directorio de Archivos de Salida
txtOutPut.ToolTipText = "Directorio de Archivos donde saldrán los archivos decodificados."
'interación de comandos
txtCL.ToolTipText = "Te permite ver la interacción de comandos con el programa de consola ffmpeg.exe"
'bóton Acualizar
cmdAcualizar.ToolTipText = "Te permite actualizar el software a una versión más reciente de la misma."
'bóton Abrir Carpeta
cmdcarpeta.ToolTipText = "Te permite Abrir la carpeta con el directorio de archivos Almacenados o en estado vacío."
'bóton cambiar directorio de archivos
cmdDirectorio.ToolTipText = "Te permite Cambiar el Directorio de Almacenamiento de Archivos."
'StatusBar1
With StatusBar1
    .Panels(1).ToolTipText = "Te permite Visualizar la Fecha en Tu sistema."
    .Panels(2).ToolTipText = "Te permite Visualizar la Hora  en Tu sistema."
    .Panels(3).ToolTipText = "Logo del Software."
End With
'bóton visualizar fb.
cmdMiRegistroDeFacebook.ToolTipText = "Visualizar Mi Facebook."
'bóton de recursos utilizados para el proyecto.
cmdWindows.ToolTipText = "Recursos de Preset utilizados Para el proyecto."
End Sub

Private Sub Twitter_Click()
On Error GoTo nose:
Dim x As String
 x = AbrirWeb(Me.hwnd, "Open" _
 , "https://www.twitter.com/Martinxd_0", _
 &O0, &O0, 0)
nose:
End Sub

Private Sub txtCL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub txtOutPut_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ListView1_MouseUp Button, Shift, x, Y
End Sub

Private Sub VerGraficosyConsola_Click()
cmdConsola_Click
End Sub

Private Sub VerSoloGraficosparaWindows_Click()
cmdgraficos_Click
End Sub

Private Sub VisaulizarMotor_Click()
picPrograma_Click
End Sub

Private Sub Youtube_Click()
Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "https://www.youtube.com/channel/UCEL746zBrw1bJMMkyDxgQAQ", _
 &O0, &O0, 0)
End Sub
