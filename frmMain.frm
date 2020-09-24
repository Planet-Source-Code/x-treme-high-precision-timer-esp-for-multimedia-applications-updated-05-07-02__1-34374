VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "High Precision Timer - Demo"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4815
   StartUpPosition =   2  'Bildschirmmitte
   Begin HiPerfTimer.ctlHiPreTimer MyTimer 
      Index           =   1
      Left            =   4920
      Top             =   1680
      _extentx        =   4101
      _extenty        =   794
      interval        =   500
      periodic        =   -1  'True
      resolution      =   1
      enabled         =   -1  'True
   End
   Begin VB.TextBox label_Profiling 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   22
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox label_Profiling 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdRestartTimers 
      Caption         =   "Restart Both Timers"
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Index           =   2
      Left            =   3600
      MaxLength       =   7
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdResetCounter 
      Caption         =   "Reset Counter"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdResetCounter 
      Caption         =   "Reset Counter"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdStartTimer 
      Caption         =   "Start"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdStartTimer 
      Caption         =   "Start"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdStopTimer 
      Caption         =   "Stop"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame 
      Caption         =   "Timer 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      Begin VB.Label label_Ticks 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Timer 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
      Begin VB.Label label_Ticks 
         Alignment       =   1  'Rechts
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdStopTimer 
      Caption         =   "Stop"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin HiPerfTimer.ctlHiPreTimer MyTimer 
      Index           =   2
      Left            =   4920
      Top             =   2160
      _extentx        =   4101
      _extenty        =   794
      periodic        =   -1  'True
      resolution      =   1
      enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   -120
      TabIndex        =   12
      Top             =   3240
      Width           =   6615
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   3960
         TabIndex        =   15
         ToolTipText     =   "Exits the application."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdHomePage 
         Caption         =   "Visit Home Page"
         Height          =   615
         Left            =   2520
         TabIndex        =   14
         ToolTipText     =   "Opens the author's home page in your web browser."
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdVote 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Grafisch
         TabIndex        =   13
         ToolTipText     =   "Opens the appropriate article at www.Planet-Source-Code.com in your web browser."
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label label_Interval 
      Caption         =   "Interval (ms)"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label label_Interval 
      Caption         =   "Interval (ms)"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label label_Title 
      Caption         =   "High Precision Timer v"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label label_Copyright 
      Caption         =   "(c) 2002 by Sebastian Thomschke          http://www.sebthom.de"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
' High Precision Timer - Demo
' Copyright ©2002 by Sebastian Thomschke, All Rights Reserved.
' http://www.sebthom.de
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34374&lngWId=1
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing this
' source code without prior written consent.
'*********************************************************************
' WARNING: ANY USE BY YOU IS AT YOUR OWN RISK. I provide this code
' "as is" without warranty of any kind, either express or implied,
' including but not limited to the implied warranties of
' merchantability and/or fitness for a particular purpose.
'*********************************************************************
Option Explicit


'*********************************************************************
' Private API Declarations
'*********************************************************************
' used for opening URLs in the standard web browser
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
   As Long

' used for only allowing numbers in the interval textboxes
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) _
   As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&


'*********************************************************************
' Private Vars
'*********************************************************************
Private Profiler(1 To 2) As New clsProfiler


'*********************************************************************
' Private Form Events
'*********************************************************************
Private Sub Form_Load()
   label_Title.Caption = label_Title.Caption & App.Major & "." & App.Minor & App.Revision
   
   txtInterval(1).Text = CStr(MyTimer(1).Interval)
   txtInterval(2).Text = CStr(MyTimer(2).Interval)
   
   NumbersOnly txtInterval(1), True
   NumbersOnly txtInterval(2), True
   
   Debug.Print "Profiler's API Overhead: " & Profiler(1).getAPIOverhead & " ms"
   
   Profiler(1).StartProfiler
   Profiler(2).StartProfiler
End Sub


'*********************************************************************
' Private Timer Events
'*********************************************************************
Private Sub MyTimer_Timer(Index As Integer)
   Dim t As Long
   
   t = Profiler(Index).StopProfiler
   frmMain.label_Profiling(Index).Text = "Timer " & Index & " fired after " & t & " ms"
   frmMain.label_Ticks(Index).Caption = Val(frmMain.label_Ticks(Index).Caption) + 1
   Profiler(Index).StartProfiler
End Sub


'*********************************************************************
' Private TextBox Events
'*********************************************************************
Private Sub txtInterval_Change(Index As Integer)
   MyTimer(Index).Interval = Val(txtInterval(Index).Text)
End Sub


'*********************************************************************
' Private Button Events
'*********************************************************************
Private Sub cmdStartTimer_Click(Index As Integer)
   MyTimer(Index).Enabled = True
End Sub

Private Sub cmdStopTimer_Click(Index As Integer)
   MyTimer(Index).Enabled = False
End Sub

Private Sub cmdResetCounter_Click(Index As Integer)
   label_Ticks(Index).Caption = "0"
End Sub

Private Sub cmdVote_Click()
   OpenURL "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34374&lngWId=1"
End Sub

Private Sub cmdHomePage_Click()
   OpenURL App.Comments
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdRestartTimers_Click()
   label_Ticks(1).Caption = "0"
   label_Ticks(2).Caption = "0"
   
   MyTimer(1).Enabled = True
   MyTimer(2).Enabled = True
End Sub


'*********************************************************************
' other methods
'*********************************************************************
Private Sub OpenURL(ByVal URL As String)
   ShellExecute 0&, "OPEN", URL, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub NumbersOnly(MyTextBox As TextBox, Flag As Boolean)
   Dim Style As Long

   Style = GetWindowLong(MyTextBox.hwnd, GWL_STYLE)

   Call SetWindowLong( _
      MyTextBox.hwnd, _
      GWL_STYLE, _
      IIf(Flag, Style Or ES_NUMBER, Style And (Not ES_NUMBER)) _
   )
End Sub
