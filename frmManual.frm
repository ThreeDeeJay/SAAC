VERSION 5.00
Begin VB.Form frmManual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual Select"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox kList 
      Height          =   3765
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
manBut = -1
Unload Me
End Sub

Private Sub cmdOK_Click()
manBut = kList.ItemData(kList.ListIndex)
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
manBut = -1
Dim iC As Long
kList.Clear
For jLoop = 0 To 307
If Len(GetKeyboardString(jLoop)) > 0 Then
kList.AddItem GetKeyboardString(jLoop), iC
kList.ItemData(iC) = jLoop
iC = iC + 1
End If
Next jLoop

End Sub
