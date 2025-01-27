VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   480
      ScaleHeight     =   4515
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim testl As Long
Picture1.Cls
For testl = 0 To 100
Picture1.PSet (testl * 30, (LSus(testl) * 30)), vbRed
Next testl
End Sub
