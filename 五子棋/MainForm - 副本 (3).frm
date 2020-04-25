VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   Caption         =   "五子棋"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14445
   Icon            =   "MainForm - 副本 (3).frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   14445
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13560
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "UNDO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      ToolTipText     =   "can only undo 1 turn"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdRedraw 
      Caption         =   "Redraw field"
      BeginProperty Font 
         Name            =   "Monaco"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   10200
      Width           =   1695
   End
   Begin VB.PictureBox picField 
      BackColor       =   &H00A0C0C0&
      FillStyle       =   0  'Solid
      Height          =   10065
      Left            =   120
      ScaleHeight     =   10005
      ScaleWidth      =   10005
      TabIndex        =   10
      Top             =   120
      Width           =   10065
   End
   Begin VB.TextBox txtPct 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0/0"
      Top             =   10200
      Width           =   735
   End
   Begin VB.PictureBox AIColor 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11880
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      ToolTipText     =   "click to change"
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox PlyColor 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      ToolTipText     =   "click to change"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   4
      ToolTipText     =   "Player-AI-Random"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "@Fixedsys"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   1
      ToolTipText     =   "START THE GAME!!"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Ink Draft"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   0
      Text            =   "Player"
      Top             =   120
      Width           =   3975
   End
   Begin VB.ListBox lblOutPut 
      BeginProperty Font 
         Name            =   "Monaco"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7860
      ItemData        =   "MainForm - 副本 (3).frx":10CA
      Left            =   10320
      List            =   "MainForm - 副本 (3).frx":10D7
      TabIndex        =   15
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Wins/Games"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "If your game field turn blank anyway, just click"
      BeginProperty Font 
         Name            =   "Monaco"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   10200
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "AI   Color"
      Height          =   375
      Left            =   12360
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Player Color"
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lbl1st 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ask for Declaration
Dim Ply1st As Boolean
Dim Turns, A, B As Integer
Dim Dot(-4 To 24, -4 To 24) As Integer  '0-None; 1-Payer; 2-AI; 3- Wall
Dim Points(1 To 19, 1 To 19) As Long
Dim MaxPointX, MaxPointY, MaxPoint, AIX, AIY, PlyX, PlyY, Games, Wins As Integer
Dim PlyMove, AIMove, win As Boolean
Dim Space As Integer
'Five-In-A-Row with half-done but already powerful AI
'(c)2020 Darth Jesus Yan

Private Sub Case1()
On Error Resume Next
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A, B - 1) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B - 1) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A - 1, B + 1) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 1048576
End Sub

Private Sub Case2()
On Error Resume Next
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 1 And Dot(A - 4, B) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 1 And Dot(A + 1, B) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B) = 1 And Dot(A + 3, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A + 4, B) = 1 And Dot(A + 3, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 1 And Dot(A, B - 4) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 1 And Dot(A, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A, B - 1) = 1 And Dot(A, B + 3) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A, B + 4) = 1 And Dot(A, B + 3) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 1 And Dot(A - 4, B - 4) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 1 And Dot(A + 1, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B - 1) = 1 And Dot(A + 3, B + 3) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A + 4, B + 4) = 1 And Dot(A + 3, B + 3) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 1 And Dot(A - 4, B + 4) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 1 And Dot(A + 1, B - 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A - 1, B + 1) = 1 And Dot(A + 3, B - 3) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 1) Then Points(A, B) = Points(A, B) + 32768
If (Dot(A + 4, B - 4) = 1 And Dot(A + 3, B - 3) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 1) Then Points(A, B) = Points(A, B) + 32768
End Sub

Private Sub Case3()
On Error Resume Next
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 2, B) = 0 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 3, B) = 0 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 4, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 3) = 0 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A, B - 1) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 2) = 0 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 1) = 0 And Dot(A, B + 4) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 3, B - 3) = 0 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B - 1) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 1, B - 1) = 0 And Dot(A + 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 3, B + 3) = 0 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A - 1, B + 1) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8192
If (Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 1, B + 1) = 0 And Dot(A + 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 8192
End Sub

Private Sub Case4()
On Error Resume Next
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A - 4, B) Mod 2 = 1 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A - 3, B) Mod 2 = 1 And Dot(A + 2, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A - 2, B) Mod 2 = 1 And Dot(A + 3, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A - 1, B) Mod 2 = 1 And Dot(A + 4, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A + 1, B) Mod 2 = 1 And Dot(A - 4, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 2, B) Mod 2 = 1 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 3, B) Mod 2 = 1 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2 And Dot(A + 4, B) Mod 2 = 1 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B - 4) Mod 2 = 1 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 3) Mod 2 = 1 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 2) Mod 2 = 1 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 1) Mod 2 = 1 And Dot(A, B + 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B + 1) Mod 2 = 1 And Dot(A, B - 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B + 2) Mod 2 = 1 And Dot(A, B - 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B + 3) Mod 2 = 1 And Dot(A, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2 And Dot(A, B + 4) Mod 2 = 1 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) Mod 2 = 1 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 3, B - 3) Mod 2 = 1 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 2, B - 2) Mod 2 = 1 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 1, B - 1) Mod 2 = 1 And Dot(A + 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A + 1, B + 1) Mod 2 = 1 And Dot(A - 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A + 2, B + 2) Mod 2 = 1 And Dot(A - 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A + 3, B + 3) Mod 2 = 1 And Dot(A - 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A + 4, B + 4) Mod 2 = 1 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) Mod 2 = 1 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 3, B + 3) Mod 2 = 1 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 2, B + 2) Mod 2 = 1 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 1, B + 1) Mod 2 = 1 And Dot(A + 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A + 1, B - 1) Mod 2 = 1 And Dot(A - 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A + 2, B - 2) Mod 2 = 1 And Dot(A - 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A + 3, B - 3) Mod 2 = 1 And Dot(A - 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 1, B + 1) Mod 2 = 1 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 0 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 0 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 0 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 0 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 0 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 0 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 0 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 0 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 0 And Dot(A - 4, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B) = 0 And Dot(A - 2, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 0 And Dot(A, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B - 1) = 0 And Dot(A, B - 2) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 0 And Dot(A - 4, B - 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 0 And Dot(A - 4, B + 4) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A - 1, B + 1) = 0 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 2) Then Points(A, B) = Points(A, B) + 2048
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 2048
End Sub

Private Sub Case5()
On Error Resume Next
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 1 And Dot(A - 4, B) = 0 And (Dot(A - 5, B) = 0 Or Dot(A + 1, B) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 3, B) = 1 And Dot(A + 4, B) = 0 And (Dot(A + 5, B) = 0 Or Dot(A - 1, B) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 1 And Dot(A, B - 4) = 0 And (Dot(A, B - 5) = 0 Or Dot(A, B + 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B + 1) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 3) = 1 And Dot(A, B + 4) = 0 And (Dot(A, B + 5) = 0 Or Dot(A, B - 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 1 And Dot(A - 4, B - 4) = 0 And (Dot(A - 5, B - 5) = 0 Or Dot(A + 1, B + 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 3, B + 3) = 1 And Dot(A + 4, B + 4) = 0 And (Dot(A + 5, B + 5) = 0 Or Dot(A - 1, B - 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 1 And Dot(A - 4, B + 4) = 0 And (Dot(A - 5, B + 5) = 0 Or Dot(A + 1, B - 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 3, B - 3) = 1 And Dot(A + 4, B - 4) = 0 And (Dot(A + 5, B - 5) = 0 Or Dot(A - 1, B + 1) = 0)) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 3, B) = 0 And Dot(A + 4, B) = 1 And Dot(A + 5, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 0 And Dot(A - 4, B) = 1 And Dot(A - 5, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B + 1) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 3) = 0 And Dot(A, B + 4) = 1 And Dot(A, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 0 And Dot(A, B - 4) = 1 And Dot(A, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 3, B + 3) = 0 And Dot(A + 4, B + 4) = 1 And Dot(A + 5, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 0 And Dot(A - 4, B - 4) = 1 And Dot(A - 5, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 3, B - 3) = 0 And Dot(A + 4, B - 4) = 1 And Dot(A + 5, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 0 And Dot(A - 4, B + 4) = 1 And Dot(A - 5, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 1 And Dot(A + 4, B) = 1 And Dot(A + 5, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 1 And Dot(A - 4, B) = 1 And Dot(A - 5, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 1 And Dot(A, B + 4) = 1 And Dot(A, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 1 And Dot(A, B - 4) = 1 And Dot(A, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 0 And Dot(A + 3, B + 3) = 1 And Dot(A + 4, B + 4) = 1 And Dot(A + 5, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 1 And Dot(A - 4, B - 4) = 1 And Dot(A - 5, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 1 And Dot(A + 4, B - 4) = 1 And Dot(A + 5, B - 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 1 And Dot(A - 4, B + 4) = 1 And Dot(A - 5, B + 5) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B - 3) = 0 And Dot(A, B - 2) = 1 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A, B + 3) = 0 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 1 And Dot(A, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 3, B) = 0 And Dot(A - 2, B) = 1 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 3, B) = 0 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 1 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 3, B + 3) = 0 And Dot(A - 2, B + 2) = 1 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A - 3, B - 3) = 0 And Dot(A - 2, B - 2) = 1 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 3, B + 3) = 0 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 2056
If (Dot(A + 3, B - 3) = 0 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 2056
End Sub

Private Sub Case6()
On Error Resume Next
If (Dot(A + 1, B) = 2 And Dot(A + 2, B) = 2 And Dot(A + 3, B) = 0 And Dot(A - 1, B) = 0 And (Dot(A - 2, B) = 0 Or Dot(A + 4, B) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 0 And Dot(A + 1, B) = 0 And (Dot(A + 2, B) = 0 Or Dot(A - 4, B) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B + 1) = 2 And Dot(A, B + 2) = 2 And Dot(A, B + 3) = 0 And Dot(A, B - 1) = 0 And (Dot(A, B - 2) = 0 Or Dot(A, B + 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 0 And Dot(A, B + 1) = 0 And (Dot(A, B + 2) = 0 Or Dot(A, B - 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B + 1) = 2 And Dot(A + 2, B + 2) = 2 And Dot(A + 3, B + 3) = 0 And Dot(A - 1, B - 1) = 0 And (Dot(A - 2, B - 2) = 0 Or Dot(A + 4, B + 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 0 And Dot(A + 1, B + 1) = 0 And (Dot(A + 2, B + 2) = 0 Or Dot(A - 4, B - 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B - 1) = 2 And Dot(A + 2, B - 2) = 2 And Dot(A + 3, B - 3) = 0 And Dot(A - 1, B + 1) = 0 And (Dot(A - 2, B + 2) = 0 Or Dot(A + 4, B - 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 0 And Dot(A + 1, B - 1) = 0 And (Dot(A + 2, B - 2) = 0 Or Dot(A - 4, B + 4) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B) = 2 And Dot(A + 1, B) = 2 And Dot(A - 2, B) = 0 And Dot(A + 2, B) = 0 And (Dot(A - 3, B) = 0 Or Dot(A + 3, B) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B - 1) = 2 And Dot(A, B + 1) = 2 And Dot(A, B - 2) = 0 And Dot(A, B + 2) = 0 And (Dot(A, B - 3) = 0 Or Dot(A, B + 3) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B - 1) = 2 And Dot(A + 1, B + 1) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A + 2, B + 2) = 0 And (Dot(A - 3, B - 3) = 0 Or Dot(A + 3, B + 3) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B + 1) = 2 And Dot(A + 1, B - 1) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A + 2, B - 2) = 0 And (Dot(A - 3, B + 3) = 0 Or Dot(A + 3, B - 3) = 0)) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B) = 0 And Dot(A + 1, B) = 2 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 2 And Dot(A + 4, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B) = 0 And Dot(A - 1, B) = 2 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B - 1) = 0 And Dot(A, B + 1) = 2 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 2 And Dot(A, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B + 1) = 0 And Dot(A, B - 1) = 2 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B - 1) = 0 And Dot(A + 1, B + 1) = 2 And Dot(A + 2, B + 2) = 0 And Dot(A + 3, B + 3) = 2 And Dot(A + 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B + 1) = 0 And Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B + 1) = 0 And Dot(A + 1, B - 1) = 2 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 2 And Dot(A + 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B - 1) = 0 And Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B) = 0 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 2 And Dot(A + 3, B) = 2 And Dot(A + 4, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B) = 0 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 2 And Dot(A - 4, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B - 1) = 0 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 2 And Dot(A, B + 3) = 2 And Dot(A, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B + 1) = 0 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 2 And Dot(A, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B - 1) = 0 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 2 And Dot(A + 3, B + 3) = 2 And Dot(A + 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B + 1) = 0 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B + 1) = 0 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 2 And Dot(A + 3, B - 3) = 2 And Dot(A + 4, B - 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B - 1) = 0 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B) = 2 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 2 And Dot(A + 3, B) = 0 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B) = 2 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 2 And Dot(A - 3, B) = 0 And Dot(A + 2, B) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B - 1) = 2 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 2 And Dot(A, B + 3) = 0 And Dot(A, B - 2) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A, B + 1) = 2 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 2 And Dot(A, B - 3) = 0 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B - 1) = 2 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 2 And Dot(A + 3, B + 3) = 0 And Dot(A - 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B + 1) = 2 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 0 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A - 1, B + 1) = 2 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 2 And Dot(A + 3, B - 3) = 0 And Dot(A - 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 64
If (Dot(A + 1, B - 1) = 2 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 0 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 64
End Sub
'(c)2020 Darth Jesus Yan

Private Sub Case7()
On Error Resume Next

'then points(a,b)=points(a,b)+2047
End Sub

Private Sub Case8()
On Error Resume Next

'then points(a,b)=points(a,b)+1024
End Sub

Private Sub Case9()
On Error Resume Next
If (Dot(A + 1, B) = 1 And Dot(A + 2, B) = 1 And Dot(A + 3, B) = 0 And Dot(A + 4, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 0 And Dot(A - 4, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B + 1) = 1 And Dot(A, B + 2) = 1 And Dot(A, B + 3) = 0 And Dot(A, B + 4) = 0 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 0 And Dot(A, B - 4) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 1 And Dot(A + 3, B + 3) = 0 And Dot(A + 4, B + 4) = 0 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 0 And Dot(A - 4, B - 4) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 1 And Dot(A + 3, B - 3) = 0 And Dot(A + 4, B - 4) = 0 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 0 And Dot(A - 4, B + 4) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 1 And Dot(A + 4, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B) = 1 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 1 And Dot(A - 4, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 1 And Dot(A, B + 4) = 0 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B - 1) = 1 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 1 And Dot(A, B - 4) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 0 And Dot(A + 3, B + 3) = 1 And Dot(A + 4, B + 4) = 0 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 1 And Dot(A - 4, B - 4) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 1 And Dot(A + 4, B - 4) = 0 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 1 And Dot(A - 4, B + 4) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 2, B) = 0 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0 And (Dot(A - 3, B) = 0 Or Dot(A + 3, B) = 0)) Then Points(A, B) = Points(A, B) + 1025
If (Dot(A, B - 2) = 0 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0 And (Dot(A, B - 3) = 0 Or Dot(A, B + 3) = 0)) Then Points(A, B) = Points(A, B) + 1025
If (Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 0 And (Dot(A - 3, B - 3) = 0 Or Dot(A + 3, B + 3) = 0)) Then Points(A, B) = Points(A, B) + 1025
If (Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0 And (Dot(A - 3, B + 3) = 0 Or Dot(A + 3, B - 3) = 0)) Then Points(A, B) = Points(A, B) + 1025
If (Dot(A - 2, B) = 0 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 1 And Dot(A + 3, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 2, B) = 0 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B - 2) = 0 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 1 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A, B + 2) = 0 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 1 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 1 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 32
If (Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 32
End Sub

Private Sub Case10()
On Error Resume Next
If Dot(A + 1, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B) = 0 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A, B + 1) = 0 And Dot(A, B - 2) = 0 And Dot(A, B - 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A, B - 1) = 0 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B + 1) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B - 1) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B - 1) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B + 1) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 2 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B) = 0 And Dot(A - 3, B) = 0 And Dot(A - 2, B) = 2 And Dot(A - 1, B) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B) = 0 And Dot(A + 3, B) = 0 And Dot(A + 2, B) = 2 And Dot(A + 1, B) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A, B + 1) = 0 And Dot(A, B - 3) = 0 And Dot(A, B - 2) = 2 And Dot(A, B - 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A, B - 1) = 0 And Dot(A, B + 3) = 0 And Dot(A, B + 2) = 2 And Dot(A, B + 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B + 1) = 0 And Dot(A - 3, B - 3) = 0 And Dot(A - 2, B - 2) = 2 And Dot(A - 1, B - 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B - 1) = 0 And Dot(A + 3, B + 3) = 0 And Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B - 1) = 0 And Dot(A - 3, B + 3) = 0 And Dot(A - 2, B + 2) = 2 And Dot(A - 1, B + 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A - 1, B + 1) = 0 And Dot(A + 3, B - 3) = 0 And Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 0 Then Points(A, B) = Points(A, B) + 33
If Dot(A + 1, B) = 0 And Dot(A - 4, B) = 0 And Dot(A - 3, B) = 2 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A - 1, B) = 0 And Dot(A + 4, B) = 0 And Dot(A + 3, B) = 2 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A, B + 1) = 0 And Dot(A, B - 4) = 0 And Dot(A, B - 3) = 2 And Dot(A, B - 2) = 0 And Dot(A, B - 1) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A, B - 1) = 0 And Dot(A, B + 4) = 0 And Dot(A, B + 3) = 2 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A + 1, B + 1) = 0 And Dot(A - 4, B - 4) = 0 And Dot(A - 3, B - 3) = 2 And Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A - 1, B - 1) = 0 And Dot(A + 4, B + 4) = 0 And Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A + 1, B - 1) = 0 And Dot(A - 4, B + 4) = 0 And Dot(A - 3, B + 3) = 2 And Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 0 Then Points(A, B) = Points(A, B) + 32
If Dot(A - 1, B + 1) = 0 And Dot(A + 4, B - 4) = 0 And Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 0 Then Points(A, B) = Points(A, B) + 32
End Sub

Private Sub Case11()
On Error Resume Next
If (Dot(A - 3, B) = 2 And Dot(A - 2, B) = 1 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B - 3) = 2 And Dot(A - 2, B - 2) = 1 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B + 3) = 2 And Dot(A - 2, B + 2) = 1 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B) = 2 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 3) = 2 And Dot(A, B - 2) = 1 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 3) = 2 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B) = 2 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B) = 2 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 1 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 2) = 2 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 1 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 2) = 2 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B - 2) = 2 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 3) = 0 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B + 2) = 2 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 4, B) = 2 And Dot(A - 3, B) = 1 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 1 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 4) = 2 And Dot(A, B - 3) = 1 And Dot(A, B - 2) = 0 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 1 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 4, B - 4) = 2 And Dot(A - 3, B - 3) = 1 And Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 1 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 4, B + 4) = 2 And Dot(A - 3, B + 3) = 1 And Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 1 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B) = 2 And Dot(A - 1, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 1 And Dot(A + 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B) = 2 And Dot(A + 1, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 1 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 2) = 2 And Dot(A, B - 1) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 1 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 2) = 2 And Dot(A, B + 1) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 1 And Dot(A, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B - 2) = 2 And Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 1 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B - 2) = 2 And Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 1 And Dot(A - 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B + 2) = 2 And Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 1 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B + 2) = 2 And Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 1 And Dot(A - 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B) = 2 And Dot(A - 2, B) = 1 And Dot(A - 1, B) = 0 And Dot(A + 1, B) = 1 And Dot(A + 2, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B) = 2 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 0 And Dot(A - 1, B) = 1 And Dot(A - 2, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 3) = 2 And Dot(A, B - 2) = 1 And Dot(A, B - 1) = 0 And Dot(A, B + 1) = 1 And Dot(A, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 3) = 2 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 0 And Dot(A, B - 1) = 1 And Dot(A, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B - 3) = 2 And Dot(A - 2, B - 2) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A + 1, B + 1) = 1 And Dot(A + 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B - 3) = 2 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A - 1, B + 1) = 1 And Dot(A - 2, B + 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B + 3) = 2 And Dot(A - 2, B + 2) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A + 1, B - 1) = 1 And Dot(A + 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B + 3) = 2 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A - 1, B - 1) = 1 And Dot(A - 2, B - 2) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 5, B) = 2 And Dot(A - 4, B) = 1 And Dot(A - 3, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 5, B) = 2 And Dot(A + 4, B) = 1 And Dot(A + 3, B) = 0 And Dot(A + 2, B) = 1 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 5) = 2 And Dot(A, B - 4) = 1 And Dot(A, B - 3) = 0 And Dot(A, B - 2) = 1 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 5) = 2 And Dot(A, B + 4) = 1 And Dot(A, B + 3) = 0 And Dot(A, B + 2) = 1 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 5, B - 5) = 2 And Dot(A - 4, B - 2) = 1 And Dot(A - 3, B - 3) = 0 And Dot(A - 2, B - 2) = 1 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 5, B - 5) = 2 And Dot(A + 4, B - 4) = 1 And Dot(A + 3, B - 3) = 0 And Dot(A + 2, B - 2) = 1 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 5, B + 5) = 2 And Dot(A - 4, B + 4) = 1 And Dot(A - 3, B + 3) = 0 And Dot(A - 2, B + 2) = 1 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 5, B + 5) = 2 And Dot(A + 4, B + 4) = 1 And Dot(A + 3, B + 3) = 0 And Dot(A + 2, B + 2) = 1 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 1, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 1) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 1) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B + 1) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B - 1) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A + 3, B + 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 1, B + 1) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 1, B - 1) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B) = 1 And Dot(A - 1, B) = 0 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 2) = 1 And Dot(A, B - 1) = 0 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B - 2) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 1) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B + 2) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 1) Then Points(A, B) = Points(A, B) + 8
End Sub

Private Sub Case12()
On Error Resume Next
If (Dot(A + 1, B) = 2 And Dot(A + 2, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B) = 2 And Dot(A + 3, B) = 1 And Dot(A - 1, B) = 0 And Dot(A - 2, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B) = 2 And Dot(A + 4, B) = 1 And Dot(A - 1, B) = 0 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B) = 2 And Dot(A - 2, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 0 And Dot(A + 3, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B) = 2 And Dot(A - 3, B) = 1 And Dot(A + 1, B) = 0 And Dot(A + 2, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B) = 2 And Dot(A - 4, B) = 1 And Dot(A + 1, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 1, B + 1) = 2 And Dot(A + 2, B + 2) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A - 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B + 2) = 2 And Dot(A + 3, B + 3) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B + 3) = 2 And Dot(A + 4, B + 4) = 1 And Dot(A - 1, B - 1) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B - 1) = 2 And Dot(A - 2, B - 2) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A + 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B - 2) = 2 And Dot(A - 3, B - 3) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B - 3) = 2 And Dot(A - 4, B - 4) = 1 And Dot(A + 1, B + 1) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 1) = 2 And Dot(A, B + 2) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 0 And Dot(A, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 2) = 2 And Dot(A, B + 3) = 1 And Dot(A, B - 1) = 0 And Dot(A, B - 2) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B + 3) = 2 And Dot(A, B + 4) = 1 And Dot(A, B - 1) = 0 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 1) = 2 And Dot(A, B - 2) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 0 And Dot(A, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 2) = 2 And Dot(A, B - 3) = 1 And Dot(A, B + 1) = 0 And Dot(A, B + 2) = 0 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A, B - 3) = 2 And Dot(A, B - 4) = 1 And Dot(A, B + 1) = 0 And Dot(A, B - 2) = 0 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 1, B + 1) = 2 And Dot(A - 2, B + 2) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A + 3, B - 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 2, B + 2) = 2 And Dot(A - 3, B + 3) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A - 3, B + 3) = 2 And Dot(A - 4, B + 4) = 1 And Dot(A + 1, B - 1) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 1, B - 1) = 2 And Dot(A + 2, B - 2) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A - 3, B + 3) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 2, B - 2) = 2 And Dot(A + 3, B - 3) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 3, B - 3) = 2 And Dot(A + 4, B - 4) = 1 And Dot(A - 1, B + 1) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 8
If (Dot(A + 4, B) = 2 And Dot(A + 3, B) = 0 And Dot(A + 2, B) = 0 And Dot(A + 1, B) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A - 4, B) = 2 And Dot(A - 3, B) = 0 And Dot(A - 2, B) = 0 And Dot(A - 1, B) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A, B - 4) = 2 And Dot(A, B - 3) = 0 And Dot(A, B - 2) = 0 And Dot(A, B - 1) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A, B + 4) = 2 And Dot(A, B + 3) = 0 And Dot(A, B + 2) = 0 And Dot(A, B + 1) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A + 4, B + 4) = 2 And Dot(A + 3, B + 3) = 0 And Dot(A + 2, B + 2) = 0 And Dot(A + 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A + 4, B - 4) = 2 And Dot(A + 3, B - 3) = 0 And Dot(A + 2, B - 2) = 0 And Dot(A + 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A - 4, B - 4) = 2 And Dot(A - 3, B - 3) = 0 And Dot(A - 2, B - 2) = 0 And Dot(A - 1, B - 1) = 0) Then Points(A, B) = Points(A, B) + 7
If (Dot(A - 4, B + 4) = 2 And Dot(A - 3, B + 3) = 0 And Dot(A - 2, B + 2) = 0 And Dot(A - 1, B + 1) = 0) Then Points(A, B) = Points(A, B) + 7
End Sub

Private Sub Case13()
On Error Resume Next
If Dot(A + 1, B + 1) = 1 Then Points(A, B) = Points(A, B) + 2
If Dot(A - 1, B + 1) = 1 Then Points(A, B) = Points(A, B) + 2
If Dot(A + 1, B - 1) = 1 Then Points(A, B) = Points(A, B) + 2
If Dot(A - 1, B - 1) = 1 Then Points(A, B) = Points(A, B) + 2
If Dot(A - 1, B) = 1 Then Points(A, B) = Points(A, B) + 1
If Dot(A + 1, B) = 1 Then Points(A, B) = Points(A, B) + 1
If Dot(A, B - 1) = 1 Then Points(A, B) = Points(A, B) + 1
If Dot(A, B + 1) = 1 Then Points(A, B) = Points(A, B) + 1
End Sub
Private Sub DrawField()  'Draw game field
  Dim X, Y, r As Integer 'Declare X, Y for horizontal and vertical, r for radius
  picField.Cls
  For X = 500 To 9500 Step 500
    picField.Line (X, 500)-(X, 9500)
    picField.CurrentX = X - 150
    picField.CurrentY = 300
    picField.Print X \ 500
    picField.CurrentX = X - 150
    picField.CurrentY = 9600
    picField.Print X \ 500
  Next X
  For Y = 500 To 9500 Step 500
    picField.Line (500, Y)-(9500, Y)
    picField.CurrentY = Y - 100
    picField.CurrentX = 200
    picField.Print Y \ 500
    picField.CurrentY = Y - 100
    picField.CurrentX = 9500
    picField.Print Y \ 500
  Next Y
  picField.FillColor = vbBlack
  For X = 2000 To 8000 Step 3000
    For Y = 2000 To 8000 Step 3000
      picField.Circle (X, Y), 50
    Next Y
  Next X
End Sub

'(c)2020 Darth Jesus Yan

Private Sub cmdChange_Click() 'click change button
  Select Case lbl1st.Caption  'change text according to the text
    Case "Random"
      lbl1st.Caption = "Player"
    Case "Player"
      lbl1st.Caption = "AI"
    Case "AI"
      lbl1st.Caption = "Random"
  End Select
End Sub

Private Sub cmdRedraw_Click()
  DrawField
  Dim X, Y As Integer
  For X = 1 To 19
    For Y = 1 To 19
      If Dot(X, Y) = 1 Then
        picField.FillColor = PlyColor.BackColor
        picField.Circle (X * 500, Y * 500), 150, PlyColor.BackColor
      End If
      If Dot(X, Y) = 2 Then
        picField.FillColor = AIColor.BackColor
        picField.Circle (X * 500, Y * 500), 150, AIColor.BackColor
      End If
    Next Y
  Next X
End Sub

Private Sub cmdStart_Click()  'click start button
  cmdUndo.Enabled = False
  Dim P, Q As Integer
  Games = Games + 1
  Turns = 0
  Call DrawField
  win = False
  cmdStart.Enabled = False
  For P = 1 To 19
    For Q = 1 To 19
     Dot(P, Q) = 0
    Next Q
  Next P
  Call Judge1st
  If Ply1st = False Then
    Dot(10, 10) = 2
    picField.FillColor = AIColor.BackColor
    picField.Circle (5000, 5000), 150, AIColor.BackColor
    picField.Line (4998, 4925)-(5002, 5075), 16777215 - AIColor.BackColor, BF
    picField.Line (4925, 4998)-(5075, 5002), 16777215 - AIColor.BackColor, BF
    lblOutPut.Clear
    lblOutPut.AddItem "start game", 0
    lblOutPut.AddItem "AI First.", 0
    lblOutPut.AddItem "AI - (10,10) #1", 0
    Turns = Turns + 1
  Else
    lblOutPut.Clear
    lblOutPut.AddItem "start game", 0
    lblOutPut.AddItem txtName.Text & " First.", 0
  End If
    PlyMove = True
End Sub

Private Sub Judge1st() 'judge who 1st
  Select Case lbl1st.Caption
    Case "Player"
      Ply1st = True
    Case "AI"
      Ply1st = False
    Case "Random"
      If Int(2 * Rnd - 1) = 0 Then
        Ply1st = True
      Else
        Ply1st = False
      End If
  End Select
End Sub

Private Function ValidMove(ByVal X As Integer, ByVal Y As Integer)
If X <> 0 And Y <> 0 And Dot(X, Y) = 0 Then ValidMove = True
End Function

Private Sub getAIMove() 'get AI's move
  Call cmdRedraw_Click
  For A = 1 To 19
    For B = 1 To 19
      If Dot(A, B) <> 0 Then GoTo jump1
      Call Case1
      Call Case2
      Call Case3
      Call Case4
      Call Case5
      Call Case6
      Call Case8
      Call Case9
      Call Case10
      Call Case11
      Call Case12
      Call Case13
      If Points(A, B) > MaxPoint Then
        MaxPoint = Points(A, B)
        MaxPointX = A
        MaxPointY = B
      ElseIf Points(A, B) = MaxPoint Then
        If Rnd > 0.5 Then
          MaxPoint = Points(A, B)
          MaxPointX = A
          MaxPointY = B
        End If
      End If
jump1:
      Points(A, B) = 0
    Next B
  Next A
  MaxPoint = 0
  If ValidMove(MaxPointX, MaxPointY) Then
    AIX = MaxPointX
    AIY = MaxPointY
    picField.FillColor = AIColor.BackColor
    picField.Circle (AIX * 500, AIY * 500), 150, AIColor.BackColor
    picField.Line (AIX * 500 - 2, AIY * 500 - 75)-(AIX * 500 + 2, AIY * 500 + 75), 16777215 - AIColor.BackColor, BF
    picField.Line (AIX * 500 - 75, AIY * 500 - 2)-(AIX * 500 + 75, AIY * 500 + 2), 16777215 - AIColor.BackColor, BF
    Dot(AIX, AIY) = 2
    Turns = Turns + 1
    lblOutPut.AddItem "AI - (" & AIX & "," & AIY & ") #" & Turns, 0
    If IfWin(AIX, AIY) Then
      PlyMove = False
      lblOutPut.AddItem "Game Over. AI Wins."
  Call cmdRedraw_Click
      cmdStart.Enabled = True
      txtPct = Wins & "/" & Games
      MsgBox "AI Wins !!!"
      cmdUndo.Enabled = False
      Exit Sub
    End If
jump2:
    PlyMove = True
  End If
End Sub






Private Sub cmdUndo_Click()
  Dot(AIX, AIY) = 0
  Dot(PlyX, PlyY) = 0
  Call cmdRedraw_Click
  cmdUndo.Enabled = False
  Turns = Turns - 2
  lblOutPut.RemoveItem 0
  lblOutPut.RemoveItem 0
End Sub

Private Sub Form_Load()
 Space = 500
  AIMove = True
  Dim i As Integer
  For i = -4 To 24  'set wall
    Dot(0, i) = 3
    Dot(-1, i) = 3
    Dot(-2, i) = 3
    Dot(-3, i) = 3
    Dot(-4, i) = 3
    Dot(i, 0) = 3
    Dot(i, -1) = 3
    Dot(i, -2) = 3
    Dot(i, -3) = 3
    Dot(i, -4) = 3
    Dot(20, i) = 3
    Dot(21, i) = 3
    Dot(22, i) = 3
    Dot(23, i) = 3
    Dot(24, i) = 3
    Dot(i, 20) = 3
    Dot(i, 21) = 3
    Dot(i, 22) = 3
    Dot(i, 23) = 3
    Dot(i, 24) = 3
  Next i
End Sub
'(c)2020 Darth Jesus Yan

Private Sub picField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Gap, xOut, yOut As Integer
  On Error Resume Next
  If Button <> vbLeftButton Then Exit Sub
  If PlyMove = True Then
    For Gap = 500 To 9500 Step 500
      If X >= Gap - 100 And X <= Gap + 100 Then PlyX = Gap \ 500
      If X > Gap + 100 Or X < Gap - 100 Then xOut = xOut + 1
      If Y >= Gap - 100 And Y <= Gap + 100 Then PlyY = Gap \ 500
      If Y > Gap + 100 Or Y < Gap - 100 Then yOut = yOut + 1
    Next Gap
    If xOut = 19 Then PlyX = 0
    If yOut = 19 Then PlyY = 0
    xOut = 0
    yOut = 0
    If ValidMove(PlyX, PlyY) = True Then
      Call cmdRedraw_Click
      Dot(PlyX, PlyY) = 1
      Turns = Turns + 1
      lblOutPut.AddItem txtName.Text & " - (" & PlyX & "," & PlyY & ") #" & Turns, 0
      picField.FillColor = PlyColor.BackColor
      picField.Circle (PlyX * 500, PlyY * 500), 150, PlyColor.BackColor
      picField.Line (PlyX * 500 - 2, PlyY * 500 - 75)-(AIX * 500 + 2, AIY * 500 + 75), 16777215 - PlyColor.BackColor, BF
      picField.Line (PlyX * 500 - 75, PlyY * 500 - 2)-(AIX * 500 + 75, AIY * 500 + 2), 16777215 - PlyColor.BackColor, BF
      PlyMove = False
      cmdUndo.Enabled = True
      If IfWin(PlyX, PlyY) Then
        PlyMove = False
        lblOutPut.AddItem txtName.Text & " WINS!", 0
  Call cmdRedraw_Click
        cmdStart.Enabled = True
        Wins = Wins + 1
        txtPct = Wins & "/" & Games
        MsgBox txtName + " WINS !!!!!!"
        cmdUndo.Enabled = False
        Exit Sub
      End If
      Call getAIMove
    End If
  End If
End Sub


Private Sub PlyColor_Click()
On Error Resume Next
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlCCRGBInit
CommonDialog1.ShowColor
PlyColor.BackColor = CommonDialog1.Color
End Sub
Private Sub aiColor_Click()
On Error Resume Next
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlCCRGBInit
CommonDialog1.ShowColor
AIColor.BackColor = CommonDialog1.Color
End Sub
Private Function IfWin(ByVal A As Integer, ByVal B As Integer)
  If Dot(A, B) = Dot(A + 1, B) And Dot(A, B) = Dot(A + 2, B) And Dot(A, B) = Dot(A + 3, B) And Dot(A, B) = Dot(A + 4, B) Then IfWin = True
  If Dot(A, B) = Dot(A, B + 1) And Dot(A, B) = Dot(A, B + 2) And Dot(A, B) = Dot(A, B + 3) And Dot(A, B) = Dot(A, B + 4) Then IfWin = True
  If Dot(A, B) = Dot(A + 1, B + 1) And Dot(A, B) = Dot(A + 2, B + 2) And Dot(A, B) = Dot(A + 3, B + 3) And Dot(A, B) = Dot(A + 4, B + 4) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B) And Dot(A, B) = Dot(A - 2, B) And Dot(A, B) = Dot(A - 3, B) And Dot(A, B) = Dot(A - 4, B) Then IfWin = True
  If Dot(A, B) = Dot(A, B - 1) And Dot(A, B) = Dot(A, B - 2) And Dot(A, B) = Dot(A, B - 3) And Dot(A, B) = Dot(A, B - 4) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B - 1) And Dot(A, B) = Dot(A - 2, B - 2) And Dot(A, B) = Dot(A - 3, B - 3) And Dot(A, B) = Dot(A - 4, B - 4) Then IfWin = True
  If Dot(A, B) = Dot(A + 1, B - 1) And Dot(A, B) = Dot(A + 2, B - 2) And Dot(A, B) = Dot(A + 3, B - 3) And Dot(A, B) = Dot(A + 4, B - 4) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B + 1) And Dot(A, B) = Dot(A - 2, B + 2) And Dot(A, B) = Dot(A - 3, B + 3) And Dot(A, B) = Dot(A - 4, B + 4) Then IfWin = True
  If Dot(A, B) = Dot(A + 1, B) And Dot(A, B) = Dot(A + 2, B) And Dot(A, B) = Dot(A + 3, B) And Dot(A, B) = Dot(A - 1, B) Then IfWin = True
  If Dot(A, B) = Dot(A, B + 1) And Dot(A, B) = Dot(A, B + 2) And Dot(A, B) = Dot(A, B + 3) And Dot(A, B) = Dot(A, B - 1) Then IfWin = True
  If Dot(A, B) = Dot(A + 1, B + 1) And Dot(A, B) = Dot(A + 2, B + 2) And Dot(A, B) = Dot(A + 3, B + 3) And Dot(A, B) = Dot(A - 1, B - 1) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B) And Dot(A, B) = Dot(A - 2, B) And Dot(A, B) = Dot(A - 3, B) And Dot(A, B) = Dot(A + 1, B) Then IfWin = True
  If Dot(A, B) = Dot(A, B - 1) And Dot(A, B) = Dot(A, B - 2) And Dot(A, B) = Dot(A, B - 3) And Dot(A, B) = Dot(A, B + 1) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B - 1) And Dot(A, B) = Dot(A - 2, B - 2) And Dot(A, B) = Dot(A - 3, B - 3) And Dot(A, B) = Dot(A + 1, B + 1) Then IfWin = True
  If Dot(A, B) = Dot(A + 1, B - 1) And Dot(A, B) = Dot(A + 2, B - 2) And Dot(A, B) = Dot(A + 3, B - 3) And Dot(A, B) = Dot(A - 1, B + 1) Then IfWin = True
  If Dot(A, B) = Dot(A - 1, B + 1) And Dot(A, B) = Dot(A - 2, B + 2) And Dot(A, B) = Dot(A - 3, B + 3) And Dot(A, B) = Dot(A + 1, B - 1) Then IfWin = True
  If Dot(A, B) = Dot(A - 2, B) And Dot(A, B) = Dot(A - 1, B) And Dot(A, B) = Dot(A + 1, B) And Dot(A, B) = Dot(A + 2, B) Then IfWin = True
  If Dot(A, B) = Dot(A, B - 2) And Dot(A, B) = Dot(A, B - 1) And Dot(A, B) = Dot(A, B + 1) And Dot(A, B) = Dot(A, B + 2) Then IfWin = True
  If Dot(A, B) = Dot(A - 2, B - 2) And Dot(A, B) = Dot(A - 1, B - 1) And Dot(A, B) = Dot(A + 1, B + 1) And Dot(A, B) = Dot(A + 2, B + 2) Then IfWin = True
  If Dot(A, B) = Dot(A - 2, B + 2) And Dot(A, B) = Dot(A - 1, B + 1) And Dot(A, B) = Dot(A + 1, B - 1) And Dot(A, B) = Dot(A + 2, B - 2) Then IfWin = True
  End Function

'(c)2020 Darth Jesus Yan
