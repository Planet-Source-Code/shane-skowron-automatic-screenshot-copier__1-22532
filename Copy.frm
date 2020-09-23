VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screenshot Copier"
   ClientHeight    =   3450
   ClientLeft      =   5355
   ClientTop       =   5025
   ClientWidth     =   6570
   ClipControls    =   0   'False
   Icon            =   "Copy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6570
   Begin VB.CheckBox Check1 
      Caption         =   "Minimize me"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2138
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Kill Process"
      Height          =   375
      Left            =   2633
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3698
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Maximized"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2093
      TabIndex        =   12
      Top             =   2040
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2093
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "5"
      Top             =   600
      Width           =   720
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2093
      TabIndex        =   5
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2093
      TabIndex        =   4
      Text            =   "1000"
      Top             =   1080
      Width           =   720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Browse"
      Height          =   285
      Left            =   5213
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5138
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1024
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start\Restart"
      Height          =   375
      Left            =   218
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   544
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   64
      Top             =   -120
   End
   Begin VB.Line Line5 
      X1              =   218
      X2              =   6338
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   218
      X2              =   6338
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   218
      X2              =   6338
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   218
      X2              =   6338
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label7 
      Caption         =   "My state"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   458
      TabIndex        =   15
      Top             =   2520
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   1973
      X2              =   1973
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   240
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label6 
      Caption         =   "Window state"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   413
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Path to run"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   413
      TabIndex        =   10
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Delay after exec."
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   413
      TabIndex        =   9
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(miliseconds)"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2978
      TabIndex        =   8
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Seconds to exec."
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   413
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Screenshot Copier"
      BeginProperty Font 
         Name            =   "Civic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2310
   End
   Begin VB.Menu program 
      Caption         =   "Program"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Begin VB.Menu aboutsc 
         Caption         =   "About Screenshot Copier"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3

Private Sub aboutsc_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
If Check1.Value = Checked Then
    Me.Hide
    frmTMinus.Show
End If

Dim winState As Integer
    If Option1.Value = True Then
        winState = 3
    ElseIf Option2.Value = True Then
        winState = 1
    End If
Text1.Text = 5
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Dim path As String
cd.ShowOpen
path = cd.FileName
Text3.Text = path
    If path = "" Or 0 Then
        Text3.Text = ""
        Exit Sub
    End If
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Timer1_Timer()
Text1.Text = Text1.Text - 1
End Sub

Private Sub Timer2_Timer()
Dim sleeptime As Integer
sleeptime = Text2.Text

If Text1.Text = "0" Then
    Timer1.Enabled = False
    ShellExecute Me.hwnd, vbNullString, Text3.Text, vbNullString, "C:\", winState
    Sleep sleeptime
    Call keybd_event(vbKeySnapshot, 1, 0, 0)
    Timer2.Enabled = False
    MsgBox "Screenshot capture is finished. Please check to make sure your screenshot is correct."
End If

End Sub
