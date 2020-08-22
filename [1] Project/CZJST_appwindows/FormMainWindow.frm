VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KanaMaster　v0.10　by Sam Toki"
   ClientHeight    =   11670
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   18600
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   11670
   ScaleWidth      =   18600
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimerSettingsRefresher 
      Interval        =   100
      Left            =   2625
      Top             =   0
   End
   Begin VB.Timer TimerSpinningSakuraAnimation 
      Interval        =   1
      Left            =   18270
      Top             =   11340
   End
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   17850
      Top             =   11340
   End
   Begin VB.Timer TimerCalculator 
      Interval        =   90
      Left            =   3780
      Top             =   1470
   End
   Begin VB.TextBox TextboxInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1575
      MaxLength       =   1
      MouseIcon       =   "FormMainWindow.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   735
      Width           =   435
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "STOP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      MouseIcon       =   "FormMainWindow.frx":2676
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   735
      Width           =   1275
   End
   Begin VB.CommandButton CmdOption3 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   12390
      MouseIcon       =   "FormMainWindow.frx":27C8
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdOption1 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   840
      MouseIcon       =   "FormMainWindow.frx":291A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdOption2 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   6615
      MouseIcon       =   "FormMainWindow.frx":2A6C
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16590
      MouseIcon       =   "FormMainWindow.frx":2BBE
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   210
      Width           =   1800
   End
   Begin VB.CommandButton CmdStartPauseResume 
      Caption         =   "START"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      MouseIcon       =   "FormMainWindow.frx":2D10
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   210
      Width           =   1800
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   18060
      Top             =   945
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   1470
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Label LabelGameDifficultyIndexIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   11
      Top             =   5775
      Width           =   2955
   End
   Begin VB.Label LabelGameDifficultyIndexTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      TabIndex        =   10
      Top             =   5145
      Width           =   2955
   End
   Begin VB.Label LabelGameAverageReactionTimeIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   14385
      TabIndex        =   17
      Top             =   5775
      Width           =   2955
   End
   Begin VB.Label LabelGameAverageReactionTimeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Avg. React. Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   14385
      TabIndex        =   16
      Top             =   5145
      Width           =   2955
   End
   Begin VB.Label LabelGameTimeElapsedIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   14385
      TabIndex        =   15
      Top             =   4200
      Width           =   2955
   End
   Begin VB.Label LabelGameTimeElapsedTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Elapsed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   14385
      TabIndex        =   14
      Top             =   3570
      Width           =   2955
   End
   Begin VB.Label LabelGameProgressIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   9
      Top             =   4200
      Width           =   2955
   End
   Begin VB.Label LabelGameProgressTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      TabIndex        =   8
      Top             =   3570
      Width           =   2955
   End
   Begin VB.Line LineSpinningSakura5 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   5
      X1              =   735
      X2              =   735
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura4 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   5
      X1              =   630
      X2              =   630
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura3 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   5
      X1              =   525
      X2              =   525
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura2 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   5
      X1              =   420
      X2              =   420
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSpinningSakura1 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   5
      X1              =   315
      X2              =   315
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Shape ShapeGameCurrentTimeLeftProgressbar 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6000
      Left            =   13965
      Top             =   2205
      Width           =   120
   End
   Begin VB.Shape ShapeGameCurrentDifficultyProgressbar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6000
      Left            =   4515
      Top             =   2205
      Width           =   120
   End
   Begin VB.Label LabelGameCurrentTimeLeftIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   14385
      TabIndex        =   13
      Top             =   2625
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentTimeLeftTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   14385
      TabIndex        =   12
      Top             =   1995
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentDifficultyIndicator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1260
      TabIndex        =   7
      Top             =   2625
      Width           =   2955
   End
   Begin VB.Label LabelGameCurrentDifficultyTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      TabIndex        =   6
      Top             =   1995
      Width           =   2955
   End
   Begin VB.Label LabelOption3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   12390
      TabIndex        =   24
      Top             =   10815
      Width           =   5370
   End
   Begin VB.Label LabelOption2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   6615
      TabIndex        =   23
      Top             =   10815
      Width           =   5370
   End
   Begin VB.Shape ShapeGameCurrentTimeLeftBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6210
      Left            =   13965
      Top             =   1995
      Width           =   120
   End
   Begin VB.Shape ShapeGameCurrentDifficultyBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6210
      Left            =   4515
      Top             =   1995
      Width           =   120
   End
   Begin VB.Label LabelStatusbar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2625
      TabIndex        =   3
      Top             =   210
      Width           =   13350
   End
   Begin VB.Label LabelOption1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   840
      TabIndex        =   22
      Top             =   10815
      Width           =   5370
   End
   Begin VB.Label LabelClock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   16590
      TabIndex        =   5
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label LabelKanaDashboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   320.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6600
      Left            =   5260
      TabIndex        =   18
      Top             =   1800
      Width           =   8070
   End
   Begin VB.Shape ShapeLightIndicatorOption1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   735
      Top             =   8715
      Width           =   5580
   End
   Begin VB.Shape ShapeLightIndicatorOption2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   6510
      Top             =   8715
      Width           =   5580
   End
   Begin VB.Shape ShapeLightIndicatorOption3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   2010
      Left            =   12285
      Top             =   8715
      Width           =   5580
   End
   Begin VB.Shape ShapeGameProgressProgressbar 
      BackColor       =   &H00FF8800&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   2625
      Top             =   1050
      Width           =   13140
   End
   Begin VB.Shape ShapeGameProgressBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   2625
      Top             =   1050
      Width           =   13350
   End
   Begin VB.Menu MenuGame 
      Caption         =   "&Game"
      Begin VB.Menu MenuGameStartPauseResume 
         Caption         =   "Start"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuGameStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuGame1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuGameChooseOption1 
         Caption         =   "Choose Option 1"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuGameChooseOption2 
         Caption         =   "Choose Option 2"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuGameChooseOption3 
         Caption         =   "Choose Option 3"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuSoundSwitch 
      Caption         =   "Soun&d ON"
   End
   Begin VB.Menu MenuSettings 
      Caption         =   "&Settings..."
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "&About..."
   End
   Begin VB.Menu Menu2_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Ａ字あ (&L)"
      Begin VB.Menu MenuLanguageENG 
         Caption         =   "English (United States)"
         Checked         =   -1  'True
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MenuLanguageCHS 
         Caption         =   "中文（简体）"
         Enabled         =   0   'False
         Shortcut        =   +{F2}
      End
      Begin VB.Menu MenuLanguageCHT 
         Caption         =   "中文（繁體）"
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu MenuLanguageJPN 
         Caption         =   "日本語"
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu Menu3_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuEXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  KanaMaster
'
'  Powered by Sam Toki
'  Version: v0.10 Beta Version MuiltLang
'  Date:    2020/08/22 (Sat.)
'  History: First version v0.10 Beta was built on 2020/03/18.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2020 SAM TOKI STUDIO. All rights reserved. KanaMaster (TM).
'             SAM TOKI STUDIO and KanaMaster are trademarks of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public setlanguage As String
Public soundswitch As Boolean

'Declare Game...
Public gamestatus As Integer  '9-Welcome, 0-Initial, 3-Ready, 1-Ongoing, 2-Interval, 7-Paused, 4-Stopped.
Public gameresult As Integer  '0-None, 1-Winner, 2-Loser.
Public gameprogress As Single  'Range: 0~100.00
Public gametotalkana As Integer
Public gamekanarepeatedtimescount As Integer
Public gametotalcount As Integer
Public gamecombocount As Integer
Public gamecombobest As Integer
Public gamemistakecount As Integer
Public gametimeelapsed As Long  'Unit: 0.1s. eg. 10 -> 1 sec.
Public gamedifficultyindex As Integer  'Range: 0~1000
Public gamecurrentdifficulty As Integer  'Unit: 0.1. eg. 50 -> 5 sec.
Public gamecurrenttimeleft As Integer  'Unit: 0.1s. eg. 10 -> 1 sec.
Public gameaveragereactiontime As Single  'Unit: 0.1s. eg. 10 -> 1 sec.

Public lotterytotal As Integer
Public lotterynumber As Integer

Public lotterykana As String
Public lotterykanalocationX As Integer
Public lotterykanalocationY As Integer
Public kanadata As Variant  '(1 To 11, 1 To 16)
Public kanarepeatedtimesdata As Variant  '(1 To 11, 1 To 16)

Public correspondingromaji As String
Public lotteryromajilocationX As Integer
Public lotteryromajilocationY As Integer
Public romajidata As Variant  '(1 To 11, 1 To 16)

Public correctanswer As Integer
Public chosenanswer As Integer

'Declare Display...
Public gameprogressprogressbaranimationtarget As Long  'Range: 0~13350
Public gamecurrentdifficultyprogressbaranimationtarget As Long  'Range: 0~6210
Public gamecurrenttimeleftprogressbaranimationtarget As Long  'Range: 0~6210
Public spinningsakuracurrentangle As Single  'Range: -180.000~180.000. Note: 90.000 means straight up.
Public spinningsakuracurrentangle2 As Single
Public spinningsakuracurrentangle3 As Single
Public spinningsakuracurrentangle4 As Single
Public spinningsakuracurrentangle5 As Single
Public spinningsakuracurrentspeed As Single  'Range: 0.00~10.00
Public spinningsakuratargetspeed As Single  'Range: 0.00~10.00. Note: The maximum spinning speed is based on the current difficulty.

'Declare Settings...
Public gamedifficultyindexindicatordescription As String

Public setinputoption As Variant  '(1 To 3)

Public setkanaswitch As Variant  '(1 To 11)

Public setgamemode As Integer  '1-Kana, 2-Time.
Public setrepeatedtimes As Integer  'Range: 1~10
Public setspecifiedtime As Integer  'Unit: min. Range: 1~30 min.

Public setnormaldifficulty As Integer  'Unit: 0.1. eg. 20 -> 2 sec. Range: 2~50
Public setincreasedifficultygraduallyswitch As Boolean
Public setinitialdifficulty As Integer  'Unit: 0.1. eg. 50 -> 5 sec. Range: 2~50
Public setreachnormaldifficultyat As Integer  'Range: 0~100
Public setinterval As Integer  'Unit: 0.1. eg. 10 -> 1.0 sec. Range: 1~30
Public setmistakeallowedamount As Integer  'Range: 0~10

Public setblackonwhite As Boolean
Public setreducecontrast As Boolean
Public setanimationswitch As Boolean
Public sethideunnecessaryinfo As Boolean
Public setspinningsakuraswitch As Boolean

Public setcheatingswitch As Boolean
Public setcheatingshowcorrectanswer As Boolean

Public setfontswitch As Boolean

'Declare Others...
Public forloop1 As Integer
Public forloop2 As Integer
Public forloop3 As Integer

'Declare Dialog...
Public answer

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Load and Initialization...

        'Initialize Menu...
        setlanguage = "ENG"
        soundswitch = True

        'Initialize Game...
        gamestatus = 9
        gameresult = 0
        gameprogress = 0
        gametotalkana = 0
        gamekanarepeatedtimescount = 0
        gametotalcount = 0
        gamecombocount = 0
        gamecombobest = 0
        gamemistakecount = 0
        gametimeelapsed = 0
        gamedifficultyindex = 0
        gamecurrentdifficulty = 0
        gamecurrenttimeleft = 0
        gameaveragereactiontime = 0

        lotterytotal = 0
        lotterynumber = 0
        lotterykana = "??"
        kanadata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                         Array("!!", "あ", "か", "さ", "た", "な", "は", "ま", "や", "ら", "わ", "ん", "が", "ざ", "だ", "ば", "ぱ"), _
                         Array("!!", "い", "き", "し", "ち", "に", "ひ", "み", "--", "り", "--", "--", "ぎ", "じ", "ぢ", "び", "ぴ"), _
                         Array("!!", "う", "く", "す", "つ", "ぬ", "ふ", "む", "ゆ", "る", "--", "--", "ぐ", "ず", "づ", "ぶ", "ぷ"), _
                         Array("!!", "え", "け", "せ", "て", "ね", "へ", "め", "--", "れ", "--", "--", "げ", "ぜ", "で", "べ", "ぺ"), _
                         Array("!!", "お", "こ", "そ", "と", "の", "ほ", "も", "よ", "ろ", "を", "--", "ご", "ぞ", "ど", "ぼ", "ぽ"), _
 _
                         Array("!!", "ア", "カ", "サ", "タ", "ナ", "ハ", "マ", "ヤ", "ラ", "ワ", "ン", "ガ", "ザ", "ダ", "バ", "パ"), _
                         Array("!!", "イ", "キ", "シ", "チ", "ニ", "ヒ", "ミ", "--", "リ", "--", "--", "ギ", "ジ", "ヂ", "ビ", "ピ"), _
                         Array("!!", "ウ", "ク", "ス", "ツ", "ヌ", "フ", "ム", "ユ", "ル", "--", "ヴ", "グ", "ズ", "ヅ", "ブ", "プ"), _
                         Array("!!", "エ", "ケ", "セ", "テ", "ネ", "ヘ", "メ", "--", "レ", "--", "--", "ゲ", "ゼ", "デ", "ベ", "ペ"), _
                         Array("!!", "オ", "コ", "ソ", "ト", "ノ", "ホ", "モ", "ヨ", "ロ", "ヲ", "--", "ゴ", "ゾ", "ド", "ボ", "ポ"), _
 _
                         Array("!!", "ゐ", "ゑ", "ヰ", "ヱ", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                         )
        kanarepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                      Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                      )

        correspondingromaji = "??"
        romajidata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                           Array("!!", "a", "ka", "sa", "ta", "na", "ha", "ma", "ya", "ra", "wa", "n", "ga", "za", "da", "ba", "pa"), _
                           Array("!!", "i", "ki", "shi", "chi", "ni", "hi", "mi", "--", "ri", "--", "--", "gi", "ji", "ji", "bi", "pi"), _
                           Array("!!", "u", "ku", "su", "tsu", "nu", "fu", "mu", "yu", "ru", "--", "--", "gu", "zu", "zu", "bu", "pu"), _
                           Array("!!", "e", "ke", "se", "te", "ne", "he", "me", "--", "re", "--", "--", "ge", "ze", "de", "be", "pe"), _
                           Array("!!", "o", "ko", "so", "to", "no", "ho", "mo", "yo", "ro", "wo", "--", "go", "zo", "do", "bo", "po"), _
 _
                           Array("!!", "a", "ka", "sa", "ta", "na", "ha", "ma", "ya", "ra", "wa", "n", "ga", "za", "da", "ba", "pa"), _
                           Array("!!", "i", "ki", "shi", "chi", "ni", "hi", "mi", "--", "ri", "--", "--", "gi", "ji", "ji", "bi", "pi"), _
                           Array("!!", "u", "ku", "su", "tsu", "nu", "fu", "mu", "yu", "ru", "--", "v", "gu", "zu", "zu", "bu", "pu"), _
                           Array("!!", "e", "ke", "se", "te", "ne", "he", "me", "--", "re", "--", "--", "ge", "ze", "de", "be", "pe"), _
                           Array("!!", "o", "ko", "so", "to", "no", "ho", "mo", "yo", "ro", "wo", "--", "go", "zo", "do", "bo", "po"), _
 _
                           Array("!!", "wi", "we", "wi", "we", "wo", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                           )

        correctanswer = 0
        chosenanswer = 0

        'Initialize Display...
        gameprogressprogressbaranimationtarget = 0
        gamecurrentdifficultyprogressbaranimationtarget = 0
        gamecurrenttimeleftprogressbaranimationtarget = 0
        spinningsakuracurrentangle = 90
        spinningsakuracurrentspeed = 0
        spinningsakuratargetspeed = 0

        'Initialize Settings...
        gamedifficultyindexindicatordescription = "Description of the current difficulty index..."

        setinputoption = Array("!!", "1", "2", "3")

        setkanaswitch = Array("!!", True, True, True, True, True, True, True, True, True, True, False)

        setgamemode = 1
        setrepeatedtimes = 1
        setspecifiedtime = 3

        setnormaldifficulty = 20
        setincreasedifficultygraduallyswitch = True
        setinitialdifficulty = 30
        setreachnormaldifficultyat = 20
        setinterval = 10
        setmistakeallowedamount = 3

        setblackonwhite = False
        setreducecontrast = False
        setanimationswitch = True
        sethideunnecessaryinfo = False
        setspinningsakuraswitch = True

        setcheatingswitch = False
        setcheatingshowcorrectanswer = True

        setfontswitch = False
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Language...
    Public Sub MenuLanguageENG_Click()
        'Call ModuleLoadLanguage.LoadLanguageENG
    End Sub
    Public Sub MenuLanguageCHS_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHS
    End Sub
    Public Sub MenuLanguageCHT_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHT
    End Sub
    Public Sub MenuLanguageJPN_Click()
        'Call ModuleLoadLanguage.LoadLanguageJPN
    End Sub

    'CMD Menu...
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub CmdEXIT_Click()
        Call MenuEXIT_Click
    End Sub
    Public Sub MenuAbout_Click()
        FormAbout.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormAbout.windowanimationtargetleft = (Screen.Width / 2) - (12930 / 2)
        FormAbout.windowanimationtargettop = (Screen.Height / 2) - (7785 / 2)
        FormAbout.windowanimationtargetwidth = 12930
        FormAbout.windowanimationtargetheight = 7785
        FormAbout.Show
    End Sub
    Public Sub MenuSettings_Click()
        FormSettings.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormSettings.windowanimationtargetleft = (Screen.Width / 2) - (17970 / 2)
        FormSettings.windowanimationtargettop = (Screen.Height / 2) - (10725 / 2)
        FormSettings.windowanimationtargetwidth = 17970
        FormSettings.windowanimationtargetheight = 10725
        FormSettings.Show
    End Sub
    Public Sub MenuSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                MenuSoundSwitch.Caption = "Soun&d OFF"
            Case False
                soundswitch = True
                MenuSoundSwitch.Caption = "Soun&d ON"
        End Select
    End Sub

    'CMD Game...
    Public Sub MenuGameStartPauseResume_Click()
        Select Case gamestatus
            Case 9  'Status: Welcome...
                gamestatus = 9  'Hold it...
            Case 0  'Status: Initial...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Notification.wav"
                gamestatus = 3  'Into: Ready...
            Case 3  'Status: Ready...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 1  'Status: Ongoing...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 2  'Status: Interval...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 7  'Into: Paused...
            Case 7  'Status: Paused...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
                gamestatus = 2  'Into: Interval...
            Case 4  'Status: Stopped...
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Notification.wav"
                gamestatus = 3  'Into: Ready...
        End Select

        TextboxInput.SetFocus: Call GameStatusRefresher
    End Sub
    Public Sub CmdStartPauseResume_Click()
        Call MenuGameStartPauseResume_Click
    End Sub
    Public Sub MenuGameStop_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Speech Off.wav"
        gamestatus = 4: Call GameStatusRefresher: TextboxInput.SetFocus
    End Sub
    Public Sub CmdStop_Click()
        Call MenuGameStop_Click
    End Sub

    Public Sub MenuGameChooseOption1_Click()
        chosenanswer = 1
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption1_Click()
        Call MenuGameChooseOption1_Click
    End Sub
    Public Sub MenuGameChooseOption2_Click()
        chosenanswer = 2
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption2_Click()
        Call MenuGameChooseOption2_Click
    End Sub
    Public Sub MenuGameChooseOption3_Click()
        chosenanswer = 3
        If gamestatus = 1 Then Call GameRespondent
        'TextboxInput.SetFocus
    End Sub
    Public Sub CmdOption3_Click()
        Call MenuGameChooseOption3_Click
    End Sub

    Public Sub TextboxInput_Change()
        Select Case TextboxInput.Text
            Case setinputoption(1)
                Call MenuGameChooseOption1_Click
            Case setinputoption(2)
                Call MenuGameChooseOption2_Click
            Case setinputoption(3)
                Call MenuGameChooseOption3_Click
            Case ""
                Exit Sub
            Case Else
                MsgBox "CAUTION: Invalid input. You have pressed a wrong key. Please confirm that your fingers are on the right keys.", vbExclamation + vbOKOnly + vbDefaultButton1, "KanaMaster"
        End Select

        TextboxInput.Text = ""
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        LabelClock.Caption = Format((Hour(Time)), "00") & ":" & Format((Minute(Time)), "00") & ":" & Format((Second(Time)), "00")

        If gamestatus = 9 Then  'Initial welcome...
            gamestatus = 0: Call GameStatusRefresher
            FormWelcome.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
            FormWelcome.windowanimationtargetleft = (Screen.Width / 2) - (17970 / 2)
            FormWelcome.windowanimationtargettop = (Screen.Height / 2) - (10725 / 2)
            FormWelcome.windowanimationtargetwidth = 17970
            FormWelcome.windowanimationtargetheight = 10725
            FormWelcome.Show
        End If
    End Sub

    Public Sub TimerSettingsRefresher_Timer()
        'Game Difficulty Index Indicator...
            If gamedifficultyindex >= 0 Then gamedifficultyindexindicatordescription = "Is this much too easy?"
            If gamedifficultyindex >= 100 Then gamedifficultyindexindicatordescription = "Beginner Level"
            If gamedifficultyindex >= 300 Then gamedifficultyindexindicatordescription = "Friendly Level"
            If gamedifficultyindex >= 500 Then gamedifficultyindexindicatordescription = "Master Level"
            If gamedifficultyindex >= 700 Then gamedifficultyindexindicatordescription = "Even native Japanese cannot make it"
            If gamedifficultyindex >= 800 Then gamedifficultyindexindicatordescription = "MONSTER!! Level"

            FormSettings.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
            FormSettings.gamedifficultyindexprogressbaranimationtarget = gamedifficultyindex / 1000 * 8000
            FormSettings.LabelGameDifficultyIndexIndicator3.Caption = gamedifficultyindexindicatordescription

        'Input...
            If Not (FormSettings.TextboxInputOption1.Text = "") Then setinputoption(1) = FormSettings.TextboxInputOption1.Text: FormSettings.LabelInputOption1Indicator.Caption = setinputoption(1): LabelOption1.Caption = setinputoption(1): FormSettings.TextboxInputOption1.Text = ""
            If Not (FormSettings.TextboxInputOption2.Text = "") Then setinputoption(2) = FormSettings.TextboxInputOption2.Text: FormSettings.LabelInputOption2Indicator.Caption = setinputoption(2): LabelOption2.Caption = setinputoption(2): FormSettings.TextboxInputOption2.Text = ""
            If Not (FormSettings.TextboxInputOption3.Text = "") Then setinputoption(3) = FormSettings.TextboxInputOption3.Text: FormSettings.LabelInputOption3Indicator.Caption = setinputoption(3): LabelOption3.Caption = setinputoption(3): FormSettings.TextboxInputOption3.Text = ""

        'Kana Included...
            If FormSettings.CheckboxKanaIncluded01.Value = 1 Then setkanaswitch(1) = True Else setkanaswitch(1) = False
            If FormSettings.CheckboxKanaIncluded02.Value = 1 Then setkanaswitch(2) = True Else setkanaswitch(2) = False
            If FormSettings.CheckboxKanaIncluded03.Value = 1 Then setkanaswitch(3) = True Else setkanaswitch(3) = False
            If FormSettings.CheckboxKanaIncluded04.Value = 1 Then setkanaswitch(4) = True Else setkanaswitch(4) = False
            If FormSettings.CheckboxKanaIncluded05.Value = 1 Then setkanaswitch(5) = True Else setkanaswitch(5) = False
            If FormSettings.CheckboxKanaIncluded06.Value = 1 Then setkanaswitch(6) = True Else setkanaswitch(6) = False
            If FormSettings.CheckboxKanaIncluded07.Value = 1 Then setkanaswitch(7) = True Else setkanaswitch(7) = False
            If FormSettings.CheckboxKanaIncluded08.Value = 1 Then setkanaswitch(8) = True Else setkanaswitch(8) = False
            If FormSettings.CheckboxKanaIncluded09.Value = 1 Then setkanaswitch(9) = True Else setkanaswitch(9) = False
            If FormSettings.CheckboxKanaIncluded10.Value = 1 Then setkanaswitch(10) = True Else setkanaswitch(10) = False
            If FormSettings.CheckboxKanaIncluded11.Value = 1 Then setkanaswitch(11) = True Else setkanaswitch(11) = False

        'Game Mode...
            If FormSettings.RadioboxGameModeKana.Value = True Then setgamemode = 1
            If FormSettings.RadioboxGameModeTime.Value = True Then setgamemode = 2
            setrepeatedtimes = FormSettings.HScrollGameModeRepeatedTimes.Value
            setspecifiedtime = FormSettings.HScrollGameModeSpecifiedTime.Value
            FormSettings.LabelGameModeRepeatedTimesIndicator.Caption = setrepeatedtimes
            FormSettings.LabelGameModeSpecifiedTimeIndicator.Caption = setspecifiedtime & " min."

        'Difficulty...
            FormSettings.HScrollDifficultyNormalDifficulty.Max = FormSettings.HScrollDifficultyInitialDifficulty.Value

            setnormaldifficulty = FormSettings.HScrollDifficultyNormalDifficulty.Value
            FormSettings.LabelDifficultyNormalDifficultyIndicator.Caption = Format((setnormaldifficulty / 10), "0.0") & " sec."

            If FormSettings.CheckboxDifficultyIncreaseDifficultyGradually.Value = 1 Then setincreasedifficultygraduallyswitch = True Else setincreasedifficultygraduallyswitch = False
                setinitialdifficulty = FormSettings.HScrollDifficultyInitialDifficulty.Value
                FormSettings.LabelDifficultyInitialDifficultyIndicator.Caption = Format((setinitialdifficulty / 10), "0.0") & " sec."
                setreachnormaldifficultyat = FormSettings.HScrollDifficultyReachNormalDifficultyAt.Value
                FormSettings.LabelDifficultyReachNormalDifficultyAtIndicator.Caption = setreachnormaldifficultyat & "%"

            setinterval = FormSettings.HScrollDifficultyInterval.Value
            FormSettings.LabelDifficultyIntervalIndicator.Caption = Format((setinterval / 10), "0.0") & " sec."
            setmistakeallowedamount = FormSettings.HScrollDifficultyMistakeAllowedAmount.Value
            FormSettings.LabelDifficultyMistakeAllowedAmountIndicator.Caption = setmistakeallowedamount

        'Display...
            If FormSettings.CheckboxDisplayBlackOnWhite.Value = 1 Then
                setblackonwhite = True
                LabelKanaDashboard.BackColor = &H0&: LabelKanaDashboard.ForeColor = &HFFFFFF
            Else
                setblackonwhite = False
                LabelKanaDashboard.BackColor = &HFFFFFF: LabelKanaDashboard.ForeColor = &H0&
            End If

            If FormSettings.CheckboxDisplayReduceContrast.Value = 1 Then
                setreducecontrast = True
                If setblackonwhite = True Then LabelKanaDashboard.BackColor = &H404040 Else LabelKanaDashboard.BackColor = &HE0E0E0
            Else
                setreducecontrast = False
                If setblackonwhite = True Then LabelKanaDashboard.BackColor = &H0 Else LabelKanaDashboard.BackColor = &HFFFFFF
            End If

            If FormSettings.CheckboxDisplaySmoothAnimations.Value = 1 Then setanimationswitch = True Else setanimationswitch = False

            If FormSettings.CheckboxDisplayHideUnnecessaryInformation.Value = 1 Then
                sethideunnecessaryinfo = True
                LabelStatusbar.Visible = False
                LabelGameCurrentDifficultyTitle.Visible = False: LabelGameCurrentDifficultyIndicator.Visible = False
                LabelGameProgressTitle.Visible = False: LabelGameProgressIndicator.Visible = False
                LabelGameDifficultyIndexTitle.Visible = False: LabelGameDifficultyIndexIndicator.Visible = False
                LabelGameCurrentTimeLeftTitle.Visible = False: LabelGameCurrentTimeLeftIndicator.Visible = False
                LabelGameTimeElapsedTitle.Visible = False: LabelGameTimeElapsedIndicator.Visible = False
                LabelGameAverageReactionTimeTitle.Visible = False: LabelGameAverageReactionTimeIndicator.Visible = False
                LabelOption1.Visible = False: LabelOption2.Visible = False: LabelOption3.Visible = False
            Else
                sethideunnecessaryinfo = False
                LabelStatusbar.Visible = True
                LabelGameCurrentDifficultyTitle.Visible = True: LabelGameCurrentDifficultyIndicator.Visible = True
                LabelGameProgressTitle.Visible = True: LabelGameProgressIndicator.Visible = True
                LabelGameDifficultyIndexTitle.Visible = True: LabelGameDifficultyIndexIndicator.Visible = True
                LabelGameCurrentTimeLeftTitle.Visible = True: LabelGameCurrentTimeLeftIndicator.Visible = True
                LabelGameTimeElapsedTitle.Visible = True: LabelGameTimeElapsedIndicator.Visible = True
                LabelGameAverageReactionTimeTitle.Visible = True: LabelGameAverageReactionTimeIndicator.Visible = True
                LabelOption1.Visible = True: LabelOption2.Visible = True: LabelOption3.Visible = True
            End If

            If FormSettings.CheckboxDisplaySpinningSakura.Value = 1 Then
                setspinningsakuraswitch = True
                LineSpinningSakura1.Visible = True: LineSpinningSakura2.Visible = True: LineSpinningSakura3.Visible = True: LineSpinningSakura4.Visible = True: LineSpinningSakura5.Visible = True
            Else
                setspinningsakuraswitch = False
                LineSpinningSakura1.Visible = False: LineSpinningSakura2.Visible = False: LineSpinningSakura3.Visible = False: LineSpinningSakura4.Visible = False: LineSpinningSakura5.Visible = False
            End If

        'Cheating...
            If FormSettings.CheckboxCheatingSwitch.Value = 1 Then
                setcheatingswitch = True
                If FormSettings.CheckboxCheatingShowCorrectAnswer.Value = 1 Then setcheatingshowcorrectanswer = True Else setcheatingshowcorrectanswer = False
                FormSettings.CheckboxCheatingShowCorrectAnswer.Enabled = True
            Else
                setcheatingswitch = False: setcheatingshowcorrectanswer = False
                FormSettings.CheckboxCheatingShowCorrectAnswer.Enabled = False
            End If

        'Fonts (Beta)...
            If FormSettings.CheckboxFontsSwitch.Value = 1 Then
                FormSettings.TextboxFontsJpnFont.Enabled = True: FormSettings.TextboxFontsEngFont.Enabled = True: FormSettings.CmdFontsApply.Enabled = True
            Else
                FormSettings.TextboxFontsJpnFont.Enabled = False: FormSettings.TextboxFontsEngFont.Enabled = False: FormSettings.CmdFontsApply.Enabled = False
                FormMainWindow.LabelKanaDashboard.Font = "MS PGothic": FormMainWindow.CmdOption1.Font = "Microsoft Sans Serif": FormMainWindow.CmdOption2.Font = "Microsoft Sans Serif": FormMainWindow.CmdOption3.Font = "Microsoft Sans Serif"
                'MsgBox "Fonts reset to default!", vbInformation + vbOKOnly + vbDefaultButton1, "KanaMaster"
            End If
    End Sub

    Public Sub TimerCalculator_Timer()
        'Difficulty Index calculator...

            gamedifficultyindex = 0

            'Difficulty index calculation Part 1...
            For forloop1 = 1 To 5
                If setkanaswitch(forloop1) = True Then gamedifficultyindex = gamedifficultyindex + 16
            Next
            For forloop1 = 6 To 10
                If setkanaswitch(forloop1) = True Then gamedifficultyindex = gamedifficultyindex + 20
            Next
            If setkanaswitch(11) = True Then gamedifficultyindex = gamedifficultyindex + 20

            'Difficulty index calculaton Part 2...
            Select Case setgamemode
                Case 1
                    gamedifficultyindex = gamedifficultyindex + setrepeatedtimes * 40
                Case 2
                    gamedifficultyindex = gamedifficultyindex + setspecifiedtime * 10
            End Select

            'Difficulty index calculaton Part 3...
            gamedifficultyindex = gamedifficultyindex + 300 ^ ((50 - setnormaldifficulty) / 45)
            Select Case setincreasedifficultygraduallyswitch
                Case True
                    gamedifficultyindex = gamedifficultyindex + 60 ^ (1 - (setreachnormaldifficultyat / 100) * ((setinitialdifficulty - setnormaldifficulty) / 45))
                    gamedifficultyindex = gamedifficultyindex + 90 ^ (1 - (setreachnormaldifficultyat / 100) * ((setinitialdifficulty - setnormaldifficulty) / 45))
                Case False
                    gamedifficultyindex = gamedifficultyindex + 150
            End Select
            gamedifficultyindex = gamedifficultyindex + 50 ^ ((20 - setinterval) / 19)
            gamedifficultyindex = gamedifficultyindex + 100 ^ (1 - setmistakeallowedamount / 10)

            'Apply calculation result...
            LabelGameDifficultyIndexIndicator.Caption = gamedifficultyindex & " / 1000"

        'Game Progress calculator...

            gametotalkana = 0
            If setkanaswitch(1) = True Then gametotalkana = gametotalkana + 16
            If setkanaswitch(2) = True Then gametotalkana = gametotalkana + 13
            If setkanaswitch(3) = True Then gametotalkana = gametotalkana + 14
            If setkanaswitch(4) = True Then gametotalkana = gametotalkana + 13
            If setkanaswitch(5) = True Then gametotalkana = gametotalkana + 15
            If setkanaswitch(6) = True Then gametotalkana = gametotalkana + 16
            If setkanaswitch(7) = True Then gametotalkana = gametotalkana + 13
            If setkanaswitch(8) = True Then gametotalkana = gametotalkana + 15
            If setkanaswitch(9) = True Then gametotalkana = gametotalkana + 13
            If setkanaswitch(10) = True Then gametotalkana = gametotalkana + 15
            If setkanaswitch(11) = True Then gametotalkana = gametotalkana + 4

            'Prevent disabling all kanaswitch...
            If gametotalkana = 0 Then
                MsgBox "CAUTION: You are not allowed to exclude all parts of the kana. Resetting to default settings of [Kana Included].", vbExclamation + vbOKOnly + vbDefaultButton1, "KanaMaster"
                'setkanaswitch = Array("!!", True, True, True, True, True, True, True, True, True, True, False)
                FormSettings.CheckboxKanaIncluded01.Value = 1: FormSettings.CheckboxKanaIncluded02.Value = 1: FormSettings.CheckboxKanaIncluded03.Value = 1: FormSettings.CheckboxKanaIncluded04.Value = 1: FormSettings.CheckboxKanaIncluded05.Value = 1: FormSettings.CheckboxKanaIncluded06.Value = 1: FormSettings.CheckboxKanaIncluded07.Value = 1: FormSettings.CheckboxKanaIncluded08.Value = 1: FormSettings.CheckboxKanaIncluded09.Value = 1: FormSettings.CheckboxKanaIncluded10.Value = 1: FormSettings.CheckboxKanaIncluded11.Value = 0
                Exit Sub
            End If

            Select Case setgamemode
                Case 1
                    gamekanarepeatedtimescount = 0
                    For forloop1 = 1 To 11
                        For forloop2 = 1 To 16
                            If kanarepeatedtimesdata(forloop1)(forloop2) >= setrepeatedtimes Then gamekanarepeatedtimescount = gamekanarepeatedtimescount + 1
                        Next
                    Next
                    gameprogress = (gamekanarepeatedtimescount / gametotalkana) * 100
                Case 2
                    gameprogress = (gametimeelapsed / (setspecifiedtime * 600)) * 100
            End Select

            'Apply calculation result...
            LabelGameProgressIndicator.Caption = Format(gameprogress, "0.00") & "%"
            gameprogressprogressbaranimationtarget = gameprogress / 100 * 13350
            If gameprogressprogressbaranimationtarget < 0 Then gameprogressprogressbaranimationtarget = 0
            If gameprogressprogressbaranimationtarget > 13350 Then gameprogressprogressbaranimationtarget = 13350

        'Current Difficulty calculator...

            Select Case setincreasedifficultygraduallyswitch
                Case True
                    If gameprogress < setreachnormaldifficultyat Then
                        gamecurrentdifficulty = setinitialdifficulty - (setinitialdifficulty - setnormaldifficulty) * (gameprogress / setreachnormaldifficultyat)
                    Else
                        gamecurrentdifficulty = setnormaldifficulty
                    End If
                Case False
                    gamecurrentdifficulty = setnormaldifficulty
            End Select

            LabelGameCurrentDifficultyIndicator.Caption = Format((gamecurrentdifficulty / 10), "0.0")
            If (setinitialdifficulty = setnormaldifficulty) Or (setincreasedifficultygraduallyswitch = False) Then
                gamecurrentdifficultyprogressbaranimationtarget = 0.5 * 6210
            Else
                gamecurrentdifficultyprogressbaranimationtarget = (setinitialdifficulty - gamecurrentdifficulty) / (setinitialdifficulty - setnormaldifficulty) * 6210
            End If
            If gamecurrentdifficultyprogressbaranimationtarget < 0 Then gamecurrentdifficultyprogressbaranimationtarget = 0
            If gamecurrentdifficultyprogressbaranimationtarget > 6210 Then gamecurrentdifficultyprogressbaranimationtarget = 6210

        'Time Left, Time Elapsed, and Average Reaction Time calculator...

            Select Case gamestatus
                Case 3
                    'New Game initialization...
                    gameresult = 0: gameprogress = 0: gamekanarepeatedtimescount = 0: gametotalcount = 0: gamecombocount = 0: gamecombobest = 0: gamemistakecount = 0: gametimeelapsed = 0: gameaveragereactiontime = 0
                    lotterytotal = 0: lotterynumber = 0: lotterykana = "??": correspondingromaji = "??": correctanswer = 0: chosenanswer = 0
                    kanarepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                                  )
                    gamecurrenttimeleft = gamecurrenttimeleft + 1
                    LabelStatusbar.Caption = "Get Ready! --- " & Format(((30 - gamecurrenttimeleft) / 10), "0.0")
                    LabelKanaDashboard.Caption = Format(Int((40 - gamecurrenttimeleft) / 10), "0")

                    LabelGameCurrentTimeLeftIndicator.Caption = Format(((gamecurrenttimeleft / 30) * gamecurrentdifficulty / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / 30 * 6210
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 6210 Then gamecurrenttimeleftprogressbaranimationtarget = 6210

                    If gamecurrenttimeleft >= 30 Then
                        gamecurrenttimeleft = gamecurrentdifficulty
                        Call GameQuestioner: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If

                    ShapeLightIndicatorOption1.BackColor = &H808080
                    ShapeLightIndicatorOption2.BackColor = &H808080
                    ShapeLightIndicatorOption3.BackColor = &H808080
                Case 1
                    gametimeelapsed = gametimeelapsed + 1
                    gamecurrenttimeleft = gamecurrenttimeleft - 1
                    LabelStatusbar.Caption = "Cleared " & gamekanarepeatedtimescount & "/" & gametotalkana & " --- Count " & gametotalcount & " --- " & gamecombocount & " Combo --- Best " & gamecombobest & " Combo --- Mistakes " & gamemistakecount & "/" & setmistakeallowedamount

                    LabelGameCurrentTimeLeftIndicator.Caption = Format((gamecurrenttimeleft / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / gamecurrentdifficulty * 6210
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 6210 Then gamecurrenttimeleftprogressbaranimationtarget = 6210

                    'Time up judgement...
                    If gamecurrenttimeleft <= 0 Then
                        chosenanswer = 4: Call GameRespondent: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If
                Case 2
                    gametimeelapsed = gametimeelapsed + 1
                    gamecurrenttimeleft = gamecurrenttimeleft + 1
                    LabelStatusbar.Caption = "Cleared " & gamekanarepeatedtimescount & "/" & gametotalkana & " --- Count " & gametotalcount & " --- " & gamecombocount & " Combo --- Best " & gamecombobest & " Combo --- Mistakes " & gamemistakecount & "/" & setmistakeallowedamount

                    LabelGameCurrentTimeLeftIndicator.Caption = Format(((gamecurrenttimeleft / setinterval) * gamecurrentdifficulty / 10), "0.0")
                    gamecurrenttimeleftprogressbaranimationtarget = gamecurrenttimeleft / setinterval * 6210
                    If gamecurrenttimeleftprogressbaranimationtarget < 0 Then gamecurrenttimeleftprogressbaranimationtarget = 0
                    If gamecurrenttimeleftprogressbaranimationtarget > 6210 Then gamecurrenttimeleftprogressbaranimationtarget = 6210

                    'Time up judgement...
                    If gamecurrenttimeleft >= setinterval Then
                        Call GameQuestioner: GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    End If
                Case 7
                    GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                Case 9
                    GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                    Case 0
                        GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                        Case 4
                            GoTo TimerCalculator_ForceExitSelectCaseGameStatus_
                Case Else
                    MsgBox "ERROR: Game status is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
            End Select

TimerCalculator_ForceExitSelectCaseGameStatus_:

            LabelGameTimeElapsedIndicator.Caption = (Format(Int(gametimeelapsed / 600), "00")) & "' " & (Format((Int(gametimeelapsed / 10) Mod 60), "00")) & """ " & (Format((gametimeelapsed Mod 10), "0"))
            LabelGameAverageReactionTimeIndicator.Caption = Format((gameaveragereactiontime / 10), "0.000")
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerProgressbarAnimation_Timer()
        Select Case setanimationswitch
            Case True
                If ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
                If ShapeGameProgressProgressbar.Width > gameprogressprogressbaranimationtarget Then ShapeGameProgressProgressbar.Width = ShapeGameProgressProgressbar.Width - Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) / 4
                If ShapeGameProgressProgressbar.Width < gameprogressprogressbaranimationtarget Then ShapeGameProgressProgressbar.Width = ShapeGameProgressProgressbar.Width + Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) / 4
                If Abs(ShapeGameProgressProgressbar.Width - gameprogressprogressbaranimationtarget) < 10 Then ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget
TimerProgressbarAnimation_Skip1_:

                If ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
                If ShapeGameCurrentDifficultyProgressbar.Height > gamecurrentdifficultyprogressbaranimationtarget Then ShapeGameCurrentDifficultyProgressbar.Height = ShapeGameCurrentDifficultyProgressbar.Height - Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) / 4
                If ShapeGameCurrentDifficultyProgressbar.Height < gamecurrentdifficultyprogressbaranimationtarget Then ShapeGameCurrentDifficultyProgressbar.Height = ShapeGameCurrentDifficultyProgressbar.Height + Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) / 4
                If Abs(ShapeGameCurrentDifficultyProgressbar.Height - gamecurrentdifficultyprogressbaranimationtarget) < 10 Then ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget
                   ShapeGameCurrentDifficultyProgressbar.Top = 8205 - ShapeGameCurrentDifficultyProgressbar.Height
TimerProgressbarAnimation_Skip2_:

                If ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip3_
                If ShapeGameCurrentTimeLeftProgressbar.Height > gamecurrenttimeleftprogressbaranimationtarget Then ShapeGameCurrentTimeLeftProgressbar.Height = ShapeGameCurrentTimeLeftProgressbar.Height - Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) / 4
                If ShapeGameCurrentTimeLeftProgressbar.Height < gamecurrenttimeleftprogressbaranimationtarget Then ShapeGameCurrentTimeLeftProgressbar.Height = ShapeGameCurrentTimeLeftProgressbar.Height + Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) / 4
                If Abs(ShapeGameCurrentTimeLeftProgressbar.Height - gamecurrenttimeleftprogressbaranimationtarget) < 10 Then ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget
                   ShapeGameCurrentTimeLeftProgressbar.Top = 8205 - ShapeGameCurrentTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip3_:

            Case False
                If ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip4_
                ShapeGameProgressProgressbar.Width = gameprogressprogressbaranimationtarget
TimerProgressbarAnimation_Skip4_:
                If ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip5_
                ShapeGameCurrentDifficultyProgressbar.Height = gamecurrentdifficultyprogressbaranimationtarget: ShapeGameCurrentDifficultyProgressbar.Top = 8205 - ShapeGameCurrentDifficultyProgressbar.Height
TimerProgressbarAnimation_Skip5_:
                If ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip6_
                ShapeGameCurrentTimeLeftProgressbar.Height = gamecurrenttimeleftprogressbaranimationtarget: ShapeGameCurrentTimeLeftProgressbar.Top = 8205 - ShapeGameCurrentTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip6_:

        End Select
    End Sub

    Public Sub TimerSpinningSakuraAnimation_Timer()
        If (gamestatus = 1 Or gamestatus = 2 Or gamestatus = 3) Then
            spinningsakuratargetspeed = (70 - gamecurrentdifficulty) / 68 * 10
        Else
            spinningsakuratargetspeed = 0
        End If

        'Locate (2625+ShapeGameProgressProgressbar.Width, 1110) ...
        LineSpinningSakura1.X1 = 2625 + ShapeGameProgressProgressbar.Width: LineSpinningSakura1.Y1 = 1110
        LineSpinningSakura2.X1 = 2625 + ShapeGameProgressProgressbar.Width: LineSpinningSakura2.Y1 = 1110
        LineSpinningSakura3.X1 = 2625 + ShapeGameProgressProgressbar.Width: LineSpinningSakura3.Y1 = 1110
        LineSpinningSakura4.X1 = 2625 + ShapeGameProgressProgressbar.Width: LineSpinningSakura4.Y1 = 1110
        LineSpinningSakura5.X1 = 2625 + ShapeGameProgressProgressbar.Width: LineSpinningSakura5.Y1 = 1110

        'Make flower (Length set to 250) ...
        spinningsakuracurrentangle2 = spinningsakuracurrentangle - 360 / 5 * 1
        spinningsakuracurrentangle3 = spinningsakuracurrentangle - 360 / 5 * 2
        spinningsakuracurrentangle4 = spinningsakuracurrentangle - 360 / 5 * 3
        spinningsakuracurrentangle5 = spinningsakuracurrentangle - 360 / 5 * 4
        While spinningsakuracurrentangle2 < -180: spinningsakuracurrentangle2 = spinningsakuracurrentangle2 + 360: Wend
        While spinningsakuracurrentangle3 < -180: spinningsakuracurrentangle3 = spinningsakuracurrentangle3 + 360: Wend
        While spinningsakuracurrentangle4 < -180: spinningsakuracurrentangle4 = spinningsakuracurrentangle4 + 360: Wend
        While spinningsakuracurrentangle5 < -180: spinningsakuracurrentangle5 = spinningsakuracurrentangle5 + 360: Wend

        LineSpinningSakura1.X2 = LineSpinningSakura1.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle)
        LineSpinningSakura1.Y2 = LineSpinningSakura1.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle)
        LineSpinningSakura2.X2 = LineSpinningSakura2.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle2)
        LineSpinningSakura2.Y2 = LineSpinningSakura2.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle2)
        LineSpinningSakura3.X2 = LineSpinningSakura3.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle3)
        LineSpinningSakura3.Y2 = LineSpinningSakura3.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle3)
        LineSpinningSakura4.X2 = LineSpinningSakura4.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle4)
        LineSpinningSakura4.Y2 = LineSpinningSakura4.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle4)
        LineSpinningSakura5.X2 = LineSpinningSakura5.X1 + 250 * Cos(3.14 / 180 * spinningsakuracurrentangle5)
        LineSpinningSakura5.Y2 = LineSpinningSakura5.Y1 - 250 * Sin(3.14 / 180 * spinningsakuracurrentangle5)

        'Prevent constant blinking...
        If (spinningsakuratargetspeed = 0 And spinningsakuracurrentspeed = 0) Then Exit Sub

        'Spin...
        Select Case setanimationswitch
            Case True
                spinningsakuracurrentangle = spinningsakuracurrentangle - spinningsakuracurrentspeed
                If spinningsakuracurrentangle <= -180 Then spinningsakuracurrentangle = spinningsakuracurrentangle + 360
            Case False
                spinningsakuracurrentangle = 90
        End Select

        'Adjust spinning speed...
        If spinningsakuracurrentspeed < spinningsakuratargetspeed Then spinningsakuracurrentspeed = spinningsakuracurrentspeed + 0.1
        If spinningsakuracurrentspeed > spinningsakuratargetspeed Then spinningsakuracurrentspeed = spinningsakuracurrentspeed - 0.05
        If spinningsakuracurrentspeed < 0 Then spinningsakuracurrentspeed = 0
        If spinningsakuracurrentspeed > 10 Then spinningsakuracurrentspeed = 10
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] GAME ENGINE []

    Public Sub GameStatusRefresher()
        Select Case gamestatus
            Case 9
                LabelStatusbar.Caption = "Loading..."
            Case 0
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "Click [START] to Rock 'n' Roll!"
                MenuGameStartPauseResume.Caption = "Start": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = False
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = True: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "START": CmdStop.Enabled = False
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case 3
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "Ready": MenuGameStartPauseResume.Enabled = False: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = False: CmdStartPauseResume.Caption = "READY": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?"
            Case 1
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "Pause": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = True: MenuGameChooseOption2.Enabled = True: MenuGameChooseOption3.Enabled = True
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "PAUSE": CmdStop.Enabled = True
                CmdOption1.Enabled = True: CmdOption2.Enabled = True: CmdOption3.Enabled = True
            Case 2
                LabelStatusbar.Caption = ""
                MenuGameStartPauseResume.Caption = "Pause": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = False
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "PAUSE": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False
            Case 7
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "Game Paused --- Press [RESUME] to move on!"
                MenuGameStartPauseResume.Caption = "Resume": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = True
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = False: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "RESUME": CmdStop.Enabled = True
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case 4
                'Reset the time left...
                gamecurrenttimeleft = 0

                LabelStatusbar.Caption = "Game Stopped --- Press [START] to rock again!"
                MenuGameStartPauseResume.Caption = "Start": MenuGameStartPauseResume.Enabled = True: MenuGameStop.Enabled = False
                MenuGameChooseOption1.Enabled = False: MenuGameChooseOption2.Enabled = False: MenuGameChooseOption3.Enabled = False
                MenuSettings.Enabled = True: MenuAbout.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "START": CmdStop.Enabled = False
                CmdOption1.Enabled = False: CmdOption2.Enabled = False: CmdOption3.Enabled = False: CmdOption1.Caption = "?": CmdOption2.Caption = "?": CmdOption3.Caption = "?": LabelKanaDashboard.Caption = "?"
            Case Else
                MsgBox "ERROR: Game status is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
        End Select
    End Sub

    Public Sub RandomNumberGenerator()
        Randomize
        lotterynumber = Int((lotterytotal + 1) * Rnd)
        While lotterynumber = 0
            Randomize
            lotterynumber = Int((lotterytotal + 1) * Rnd)
        Wend
    End Sub

    Public Sub GameQuestioner()
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Menu Command.wav"

        gamestatus = 1: Call GameStatusRefresher: gamecurrenttimeleft = gamecurrentdifficulty
        If gameprogress >= 100 Then Exit Sub

        'Clear contents...
        LabelStatusbar.BackColor = &HC0C0C0
        LabelKanaDashboard.Caption = ""
        CmdOption1.Caption = ""
        CmdOption2.Caption = ""
        CmdOption3.Caption = ""
        ShapeLightIndicatorOption1.BackColor = &H808080
        ShapeLightIndicatorOption2.BackColor = &H808080
        ShapeLightIndicatorOption3.BackColor = &H808080

        'Step 1: Kana...
            lotterytotal = 0: lotterynumber = 0: lotterykanalocationX = 0: lotterykanalocationY = 0
            Do Until Not (kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) >= setrepeatedtimes Or kanadata(lotterykanalocationX)(lotterykanalocationY) = "!!" Or kanadata(lotterykanalocationX)(lotterykanalocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                lotterykanalocationX = lotterynumber
                lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                lotterykanalocationY = lotterynumber
            Loop
            lotterykana = kanadata(lotterykanalocationX)(lotterykanalocationY)
            correspondingromaji = romajidata(lotterykanalocationX)(lotterykanalocationY)
            LabelKanaDashboard.Caption = lotterykana

        'Step 2: The correct option...
            lotterytotal = 3: lotterynumber = 0: Call RandomNumberGenerator: correctanswer = lotterynumber
            Select Case correctanswer
                Case 1
                    CmdOption1.Caption = correspondingromaji
                Case 2
                    CmdOption2.Caption = correspondingromaji
                Case 3
                    CmdOption3.Caption = correspondingromaji
                Case Else
                    MsgBox "ERROR: Correct answer is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
            End Select

        'Step 3: Other option 1...
            Select Case correctanswer
                Case 1
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption2.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 2
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption1.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 3
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption3.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption1.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
            End Select

        'Step 4: Other option 2...
            Select Case correctanswer
                Case 1
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption3.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 2
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption2.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption3.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
                Case 3
                    lotterytotal = 0: lotterynumber = 0: lotteryromajilocationX = 0: lotteryromajilocationY = 0
                    Do Until Not (romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption3.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = CmdOption1.Caption Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "!!" Or romajidata(lotteryromajilocationX)(lotteryromajilocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                        lotterytotal = 11: lotterynumber = 0: Do Until setkanaswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                        lotteryromajilocationX = lotterynumber
                        lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                        lotteryromajilocationY = lotterynumber
                    Loop
                    CmdOption2.Caption = romajidata(lotteryromajilocationX)(lotteryromajilocationY)
            End Select

        'Cheating...
            If setcheatingswitch = True Then
                LabelStatusbar.BackColor = &HFFFF&

                If setcheatingshowcorrectanswer = True Then
                    Select Case correctanswer
                        Case 1
                            ShapeLightIndicatorOption1.BackColor = &HFFFF&
                        Case 2
                            ShapeLightIndicatorOption2.BackColor = &HFFFF&
                        Case 3
                            ShapeLightIndicatorOption3.BackColor = &HFFFF&
                        Case Else
                            MsgBox "ERROR: Correct answer is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
                    End Select
                End If
            End If
    End Sub

    Public Sub GameRespondent()
        'Average reaction time calculation...
        gametotalcount = gametotalcount + 1
        gameaveragereactiontime = (gameaveragereactiontime * (gametotalcount - 1) + (gamecurrentdifficulty - gamecurrenttimeleft)) / gametotalcount

        'Switch game status...
        gamestatus = 2: Call GameStatusRefresher: gamecurrenttimeleft = 0

        'Judgement...
        Select Case correctanswer
            Case 1
                ShapeLightIndicatorOption1.BackColor = &HFF00&
                ShapeLightIndicatorOption2.BackColor = &H808080
                ShapeLightIndicatorOption3.BackColor = &H808080
            Case 2
                ShapeLightIndicatorOption1.BackColor = &H808080
                ShapeLightIndicatorOption2.BackColor = &HFF00&
                ShapeLightIndicatorOption3.BackColor = &H808080
            Case 3
                ShapeLightIndicatorOption1.BackColor = &H808080
                ShapeLightIndicatorOption2.BackColor = &H808080
                ShapeLightIndicatorOption3.BackColor = &HFF00&
            Case Else
                MsgBox "ERROR: Correct answer is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
        End Select

        If chosenanswer = correctanswer Then
            'Answer correct sound...
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Startup.wav"

            'Combo count...
            gamecombocount = gamecombocount + 1
            If gamecombobest < gamecombocount Then gamecombobest = gamecombocount

            If setgamemode = 1 Then kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) = kanarepeatedtimesdata(lotterykanalocationX)(lotterykanalocationY) + 1

            'Winner judgement...
            Call TimerCalculator_Timer
            If gameprogress >= 100 Then
                FormGameReport.LabelGameReportWinnerLoser.Caption = "WINNER!"
                If setcheatingswitch = True Then
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = True
                Else
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = False
                End If
                FormGameReport.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
                FormGameReport.LabelGameDifficultyIndexIndicator3.Caption = FormSettings.LabelGameDifficultyIndexIndicator3.Caption
                FormGameReport.LabelGameProgressIndicator.Caption = LabelGameProgressIndicator.Caption
                FormGameReport.LabelGameCurrentDifficultyIndicator.Caption = LabelGameCurrentDifficultyIndicator.Caption & "s"
                FormGameReport.LabelGameAverageReactionTimeIndicator.Caption = LabelGameAverageReactionTimeIndicator.Caption & "s"
                FormGameReport.LabelGameTimeElapsedIndicator.Caption = LabelGameTimeElapsedIndicator.Caption
                FormGameReport.LabelGameTotalCountIndicator.Caption = gametotalcount
                FormGameReport.LabelGameComboBestIndicator.Caption = gamecombobest
                FormGameReport.LabelGameMistakeCountIndicator.Caption = gamemistakecount

                gameresult = 1: gamestatus = 0: Call GameStatusRefresher
                MsgBox "Congratulations!! You won the game." & vbCrLf & "We will show you a game report later.", vbInformation + vbOKOnly + vbDefaultButton1, "KanaMaster"

                FormGameReport.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
                FormGameReport.windowanimationtargetleft = (Screen.Width / 2) - (17970 / 2)
                FormGameReport.windowanimationtargettop = (Screen.Height / 2) - (10725 / 2)
                FormGameReport.windowanimationtargetwidth = 17970
                FormGameReport.windowanimationtargetheight = 10725
                FormGameReport.Show
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\tada.wav"
            End If
        Else
            'Answer incorrect sound...
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\chord.wav"

            Select Case chosenanswer
                Case 1
                    ShapeLightIndicatorOption1.BackColor = &HFF&
                Case 2
                    ShapeLightIndicatorOption2.BackColor = &HFF&
                Case 3
                    ShapeLightIndicatorOption3.BackColor = &HFF&
                Case 4
                    If ShapeLightIndicatorOption1.BackColor = &H808080 Then ShapeLightIndicatorOption1.BackColor = &H80FF&
                    If ShapeLightIndicatorOption2.BackColor = &H808080 Then ShapeLightIndicatorOption2.BackColor = &H80FF&
                    If ShapeLightIndicatorOption3.BackColor = &H808080 Then ShapeLightIndicatorOption3.BackColor = &H80FF&
                Case Else
                    MsgBox "ERROR: Chosen answer is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "KanaMaster"
            End Select

            'Combo reset... But do not reset best combo (gamecombobest)...
            gamecombocount = 0: gamemistakecount = gamemistakecount + 1

            'Loser judgement...
            Call TimerCalculator_Timer
            If gamemistakecount > setmistakeallowedamount Then
                FormGameReport.LabelGameReportWinnerLoser.Caption = "LOSER!"
                If setcheatingswitch = True Then
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = True
                Else
                    FormGameReport.LabelGameReportCheatedIndicator.Visible = False
                End If
                FormGameReport.LabelGameDifficultyIndexIndicator1.Caption = gamedifficultyindex
                FormGameReport.LabelGameDifficultyIndexIndicator3.Caption = FormSettings.LabelGameDifficultyIndexIndicator3.Caption
                FormGameReport.LabelGameProgressIndicator.Caption = LabelGameProgressIndicator.Caption
                FormGameReport.LabelGameCurrentDifficultyIndicator.Caption = LabelGameCurrentDifficultyIndicator.Caption & "s"
                FormGameReport.LabelGameAverageReactionTimeIndicator.Caption = LabelGameAverageReactionTimeIndicator.Caption & "s"
                FormGameReport.LabelGameTimeElapsedIndicator.Caption = LabelGameTimeElapsedIndicator.Caption
                FormGameReport.LabelGameTotalCountIndicator.Caption = gametotalcount
                FormGameReport.LabelGameComboBestIndicator.Caption = gamecombobest
                FormGameReport.LabelGameMistakeCountIndicator.Caption = gamemistakecount

                gameresult = 2: gamestatus = 0: Call GameStatusRefresher
                MsgBox "Game over..." & vbCrLf & "Unfortunately, you lost the game." & vbCrLf & "You have finished progress " & Format(gameprogress, "0.00") & "%.", vbInformation + vbOKOnly + vbDefaultButton1, "KanaMaster"

                FormGameReport.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
                FormGameReport.windowanimationtargetleft = (Screen.Width / 2) - (17970 / 2)
                FormGameReport.windowanimationtargettop = (Screen.Height / 2) - (10725 / 2)
                FormGameReport.windowanimationtargetwidth = 17970
                FormGameReport.windowanimationtargetheight = 10725
                FormGameReport.Show
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
            End If
        End If
    End Sub
