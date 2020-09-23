VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Changer"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   300
      Left            =   4080
      TabIndex        =   8
      Top             =   547
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   4080
      TabIndex        =   7
      Top             =   67
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1020
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "C:\"
      Top             =   555
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "quid erat demonstrandum"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Written by Tako - 2002"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Icon File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EXE File:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim eRR As String
ReplaceIcons Text2.Text, Text1.Text, eRR
MsgBox "Icon Changed", , "Icon Changer"
End Sub
Private Sub Command2_Click()
CD1.Filter = "EXE Files (*.exe)|*.exe"
CD1.ShowOpen
If CD1.FileName <> "" Then Text1.Text = CD1.FileName
End Sub
Private Sub Command3_Click()
CD1.Filter = "Icon Files (*.ico)|*.ico"
CD1.ShowOpen
If CD1.FileName <> "" Then Text2.Text = CD1.FileName
End Sub
