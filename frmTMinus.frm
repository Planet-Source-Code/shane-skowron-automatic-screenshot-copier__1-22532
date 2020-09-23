VERSION 5.00
Begin VB.Form frmTMinus 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Civic"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2750
      Left            =   2760
      Top             =   120
   End
End
Attribute VB_Name = "frmTMinus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
Me.AutoRedraw = True
Me.Print "TMinus 5 seconds and counting..."
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
