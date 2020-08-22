VERSION 5.00
Begin VB.Form FormWelcome 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "KanaMaster"
   ClientHeight    =   10725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
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
   Icon            =   "FormWelcome.frx":0000
   LinkTopic       =   "FormWelcome"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormWelcome.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   10725
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton RadioboxLanguageJPN 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "日本語"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11970
      MouseIcon       =   "FormWelcome.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1050
      Width           =   3480
   End
   Begin VB.OptionButton RadioboxLanguageCHT 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "中文（繁體）"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   8085
      MouseIcon       =   "FormWelcome.frx":2676
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1050
      Width           =   3480
   End
   Begin VB.OptionButton RadioboxLanguageCHS 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "中文（简体）"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4200
      MouseIcon       =   "FormWelcome.frx":27C8
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1050
      Width           =   3480
   End
   Begin VB.OptionButton RadioboxLanguageENG 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "English (United States)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   315
      MouseIcon       =   "FormWelcome.frx":291A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1050
      Value           =   -1  'True
      Width           =   3480
   End
   Begin VB.TextBox TextboxWelcome 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8550
      Left            =   315
      Locked          =   -1  'True
      MouseIcon       =   "FormWelcome.frx":2A6C
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FormWelcome.frx":2BBE
      Top             =   1785
      Width           =   17340
   End
   Begin VB.CommandButton CmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   16275
      MouseIcon       =   "FormWelcome.frx":2ED9
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   17640
      Top             =   10395
   End
   Begin VB.Label LabelWelcomeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
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
      Height          =   645
      Left            =   315
      TabIndex        =   0
      Top             =   210
      Width           =   15555
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   10725
      Left            =   0
      Top             =   0
      Width           =   17970
   End
End
Attribute VB_Name = "FormWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    Public Sub CmdOK_Click()
        windowanimationtargetleft = (Screen.Width / 2)
        windowanimationtargettop = (Screen.Height / 2)
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerWindowAnimation_Timer()
        If ((Me.Left = windowanimationtargetleft) And (Me.Top = windowanimationtargettop) And (Me.Width = windowanimationtargetwidth) And (Me.Height = windowanimationtargetheight)) Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If Me.Left > windowanimationtargetleft Then Me.Left = Me.Left - Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Left < windowanimationtargetleft Then Me.Left = Me.Left + Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Top > windowanimationtargettop Then Me.Top = Me.Top - Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Top < windowanimationtargettop Then Me.Top = Me.Top + Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Width > windowanimationtargetwidth Then Me.Width = Me.Width - Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Width < windowanimationtargetwidth Then Me.Width = Me.Width + Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Height > windowanimationtargetheight Then Me.Height = Me.Height - Abs(Me.Height - windowanimationtargetheight) / 4
                If Me.Height < windowanimationtargetheight Then Me.Height = Me.Height + Abs(Me.Height - windowanimationtargetheight) / 4
                If Abs(Me.Left - windowanimationtargetleft) < 10 Then Me.Left = windowanimationtargetleft
                If Abs(Me.Top - windowanimationtargettop) < 10 Then Me.Top = windowanimationtargettop
                If Abs(Me.Width - windowanimationtargetwidth) < 10 Then Me.Width = windowanimationtargetwidth
                If Abs(Me.Height - windowanimationtargetheight) < 10 Then Me.Height = windowanimationtargetheight
            Case False
                Me.Move windowanimationtargetleft, windowanimationtargettop, windowanimationtargetwidth, windowanimationtargetheight
        End Select

        If windowanimationtargetheight = 0 And Me.Height < 10 Then Me.Hide
    End Sub
