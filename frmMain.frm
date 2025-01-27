VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "San Andreas Advanced Control 1.2"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSec 
      Caption         =   "2nd Player"
      Height          =   255
      Left            =   6480
      TabIndex        =   558
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Timer ETimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   120
   End
   Begin VB.CommandButton bDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   5640
      TabIndex        =   542
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   4920
      TabIndex        =   541
      Top             =   6480
      Width           =   735
   End
   Begin VB.ComboBox preset 
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":0444
      TabIndex        =   435
      ToolTipText     =   "To save a preset type a name here and press save or enter"
      Top             =   6480
      Width           =   4800
   End
   Begin VB.Timer Listen 
      Interval        =   1
      Left            =   6960
      Top             =   120
   End
   Begin VB.Timer PollTimer 
      Interval        =   1
      Left            =   6600
      Top             =   120
   End
   Begin VB.ComboBox jStick 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   326
      Top             =   240
      Width           =   3495
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Axis/Test"
      Height          =   375
      Index           =   3
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Force"
      Height          =   375
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Vehicle"
      Height          =   375
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Foot"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame frmFoot 
      Caption         =   "Foot Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      Begin VB.Frame frmAim 
         Caption         =   "Aim Mode"
         Height          =   1575
         Left            =   7440
         TabIndex        =   548
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton optAim 
            Caption         =   "Auto"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   551
            ToolTipText     =   "Depends on device"
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAim 
            Caption         =   "Classic"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   550
            ToolTipText     =   "Auto targeting PS2 style"
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optAim 
            Caption         =   "Standard"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   549
            ToolTipText     =   "Manual targeting"
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.VScrollBar fScroll 
         Height          =   4815
         Left            =   7080
         Max             =   8
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox sBox 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   600
         ScaleHeight     =   4815
         ScaleWidth      =   7095
         TabIndex        =   4
         Top             =   360
         Width           =   7095
         Begin VB.PictureBox mBox 
            BorderStyle     =   0  'None
            Height          =   10455
            Left            =   -120
            ScaleHeight     =   10455
            ScaleWidth      =   6735
            TabIndex        =   5
            Top             =   0
            Width           =   6735
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   53
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   552
               Top             =   10080
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   53
                  Left            =   4080
                  TabIndex        =   556
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   53
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   555
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   53
                  Left            =   1800
                  TabIndex        =   554
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   53
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   553
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   27
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   170
               Top             =   9720
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   27
                  Left            =   4080
                  TabIndex        =   174
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   27
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   173
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   27
                  Left            =   1800
                  TabIndex        =   172
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   27
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   171
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   26
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   165
               Top             =   9360
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   26
                  Left            =   4080
                  TabIndex        =   169
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   26
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   168
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   26
                  Left            =   1800
                  TabIndex        =   167
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   26
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   166
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   25
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   160
               Top             =   9000
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   25
                  Left            =   4080
                  TabIndex        =   164
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   25
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   163
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   25
                  Left            =   1800
                  TabIndex        =   162
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   25
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   161
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   24
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   155
               Top             =   8640
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   24
                  Left            =   4080
                  TabIndex        =   159
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   24
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   158
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   24
                  Left            =   1800
                  TabIndex        =   157
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   24
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   156
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   23
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   150
               Top             =   8280
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   23
                  Left            =   4080
                  TabIndex        =   154
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   23
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   153
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   23
                  Left            =   1800
                  TabIndex        =   152
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   23
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   151
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   22
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   145
               Top             =   7920
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   22
                  Left            =   4080
                  TabIndex        =   149
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   22
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   148
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   22
                  Left            =   1800
                  TabIndex        =   147
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   22
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   146
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   21
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   140
               Top             =   7560
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   21
                  Left            =   4080
                  TabIndex        =   144
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   21
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   143
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   21
                  Left            =   1800
                  TabIndex        =   142
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   21
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   141
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   20
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   135
               Top             =   7200
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   20
                  Left            =   4080
                  TabIndex        =   139
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   20
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   138
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   20
                  Left            =   1800
                  TabIndex        =   137
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   20
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   136
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   19
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   130
               Top             =   6840
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   19
                  Left            =   4080
                  TabIndex        =   134
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   19
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   133
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   19
                  Left            =   1800
                  TabIndex        =   132
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   19
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   131
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   18
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   125
               Top             =   6480
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   18
                  Left            =   4080
                  TabIndex        =   129
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   18
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   128
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   18
                  Left            =   1800
                  TabIndex        =   127
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   18
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   126
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   17
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   120
               Top             =   6120
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   17
                  Left            =   4080
                  TabIndex        =   124
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   17
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   123
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   17
                  Left            =   1800
                  TabIndex        =   122
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   17
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   121
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   16
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   115
               Top             =   5760
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   16
                  Left            =   4080
                  TabIndex        =   119
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   16
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   118
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   16
                  Left            =   1800
                  TabIndex        =   117
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   16
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   116
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   15
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   110
               Top             =   5400
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   15
                  Left            =   4080
                  TabIndex        =   114
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   15
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   113
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   15
                  Left            =   1800
                  TabIndex        =   112
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   15
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   111
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   14
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   105
               Top             =   5040
               Width           =   4455
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   14
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   109
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   14
                  Left            =   1800
                  TabIndex        =   108
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   14
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   107
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   14
                  Left            =   4080
                  TabIndex        =   106
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   13
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   100
               Top             =   4680
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   13
                  Left            =   4080
                  TabIndex        =   104
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   13
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   103
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   13
                  Left            =   1800
                  TabIndex        =   102
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   13
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   101
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   12
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   95
               Top             =   4320
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   12
                  Left            =   4080
                  TabIndex        =   99
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   12
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   98
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   12
                  Left            =   1800
                  TabIndex        =   97
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   12
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   96
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   11
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   90
               Top             =   3960
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   11
                  Left            =   4080
                  TabIndex        =   94
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   11
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   93
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   11
                  Left            =   1800
                  TabIndex        =   92
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   11
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   91
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   10
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   85
               Top             =   3600
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   10
                  Left            =   4080
                  TabIndex        =   89
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   10
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   88
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   10
                  Left            =   1800
                  TabIndex        =   87
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   10
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   86
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   9
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   80
               Top             =   3240
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   9
                  Left            =   4080
                  TabIndex        =   84
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   9
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   83
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   9
                  Left            =   1800
                  TabIndex        =   82
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   9
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   81
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   8
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   75
               Top             =   2880
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   79
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   8
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   78
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   8
                  Left            =   1800
                  TabIndex        =   77
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   8
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   76
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   7
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   70
               Top             =   2520
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   7
                  Left            =   4080
                  TabIndex        =   74
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   7
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   73
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   7
                  Left            =   1800
                  TabIndex        =   72
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   7
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   71
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   6
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   65
               Top             =   2160
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   6
                  Left            =   4080
                  TabIndex        =   69
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   6
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   68
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   6
                  Left            =   1800
                  TabIndex        =   67
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   6
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   66
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   5
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   60
               Top             =   1800
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   5
                  Left            =   4080
                  TabIndex        =   64
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   5
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   63
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   5
                  Left            =   1800
                  TabIndex        =   62
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   5
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   61
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   4
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   55
               Top             =   1440
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   4
                  Left            =   4080
                  TabIndex        =   59
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   4
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   58
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   4
                  Left            =   1800
                  TabIndex        =   57
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   4
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   3
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   50
               Top             =   1080
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   3
                  Left            =   4080
                  TabIndex        =   54
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   3
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   3
                  Left            =   1800
                  TabIndex        =   52
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   3
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   51
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   2
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   45
               Top             =   720
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   2
                  Left            =   4080
                  TabIndex        =   49
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   2
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   48
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   2
                  Left            =   1800
                  TabIndex        =   47
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   1
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   40
               Top             =   360
               Width           =   4455
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   1
                  Left            =   4080
                  TabIndex        =   44
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   1
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   43
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   42
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   41
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   0
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   35
               Top             =   0
               Width           =   4455
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   39
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   0
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   36
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   0
                  Left            =   4080
                  TabIndex        =   37
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Pause/Menu"
               Height          =   195
               Index           =   61
               Left            =   0
               TabIndex        =   557
               Top             =   10115
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Secondary Fire"
               Height          =   195
               Index           =   27
               Left            =   0
               TabIndex        =   33
               Top             =   9765
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Center Camera"
               Height          =   195
               Index           =   26
               Left            =   0
               TabIndex        =   32
               Top             =   9405
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Up"
               Height          =   195
               Index           =   25
               Left            =   0
               TabIndex        =   31
               Top             =   9045
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Down"
               Height          =   195
               Index           =   24
               Left            =   0
               TabIndex        =   30
               Top             =   8685
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Right"
               Height          =   195
               Index           =   23
               Left            =   0
               TabIndex        =   29
               Top             =   8325
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Left"
               Height          =   195
               Index           =   22
               Left            =   0
               TabIndex        =   28
               Top             =   7965
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Behind"
               Height          =   195
               Index           =   21
               Left            =   0
               TabIndex        =   27
               Top             =   7605
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Walk"
               Height          =   195
               Index           =   20
               Left            =   0
               TabIndex        =   26
               Top             =   7245
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Action"
               Height          =   195
               Index           =   19
               Left            =   0
               TabIndex        =   25
               Top             =   6885
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Crouch"
               Height          =   195
               Index           =   18
               Left            =   0
               TabIndex        =   24
               Top             =   6525
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Aim Weapon"
               Height          =   195
               Index           =   17
               Left            =   0
               TabIndex        =   23
               Top             =   6165
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Sprint"
               Height          =   195
               Index           =   16
               Left            =   0
               TabIndex        =   22
               Top             =   5805
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Jump"
               Height          =   195
               Index           =   15
               Left            =   0
               TabIndex        =   21
               Top             =   5445
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Change Camera"
               Height          =   195
               Index           =   14
               Left            =   0
               TabIndex        =   20
               Top             =   5085
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Enter + Exit"
               Height          =   195
               Index           =   13
               Left            =   0
               TabIndex        =   19
               Top             =   4725
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Zoom Out"
               Height          =   195
               Index           =   12
               Left            =   0
               TabIndex        =   18
               Top             =   4365
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Zoom In"
               Height          =   195
               Index           =   11
               Left            =   0
               TabIndex        =   17
               Top             =   4005
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Right"
               Height          =   195
               Index           =   10
               Left            =   0
               TabIndex        =   16
               Top             =   3645
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Left"
               Height          =   195
               Index           =   9
               Left            =   0
               TabIndex        =   15
               Top             =   3285
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Backwards"
               Height          =   195
               Index           =   8
               Left            =   0
               TabIndex        =   14
               Top             =   2925
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Forward"
               Height          =   195
               Index           =   7
               Left            =   0
               TabIndex        =   13
               Top             =   2565
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Conversation - Yes"
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   12
               Top             =   2205
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Conversation - No"
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   11
               Top             =   1845
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Group Ctrl Back"
               Height          =   195
               Index           =   4
               Left            =   0
               TabIndex        =   10
               Top             =   1485
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Group Ctrl Forward"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   9
               Top             =   1125
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Previous Weapon/Target"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   8
               Top             =   765
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Next Weapon/Target"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   7
               Top             =   405
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Fire"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   45
               Width           =   1920
            End
         End
      End
   End
   Begin VB.Frame frmTest 
      Caption         =   "Axis Settings / Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   325
      Top             =   600
      Width           =   8895
      Begin VB.Frame frmButtons 
         Caption         =   "Buttons"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   402
         Top             =   3840
         Width           =   8620
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "32"
            Height          =   255
            Index           =   31
            Left            =   7680
            TabIndex        =   434
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "31"
            Height          =   255
            Index           =   30
            Left            =   7200
            TabIndex        =   433
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "30"
            Height          =   255
            Index           =   29
            Left            =   6720
            TabIndex        =   432
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "29"
            Height          =   255
            Index           =   28
            Left            =   6240
            TabIndex        =   431
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "28"
            Height          =   255
            Index           =   27
            Left            =   5760
            TabIndex        =   430
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "27"
            Height          =   255
            Index           =   26
            Left            =   5280
            TabIndex        =   429
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "26"
            Height          =   255
            Index           =   25
            Left            =   4800
            TabIndex        =   428
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "25"
            Height          =   255
            Index           =   24
            Left            =   4320
            TabIndex        =   427
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "24"
            Height          =   255
            Index           =   23
            Left            =   3840
            TabIndex        =   426
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "23"
            Height          =   255
            Index           =   22
            Left            =   3360
            TabIndex        =   425
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "22"
            Height          =   255
            Index           =   21
            Left            =   2880
            TabIndex        =   424
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "21"
            Height          =   255
            Index           =   20
            Left            =   2400
            TabIndex        =   423
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "20"
            Height          =   255
            Index           =   19
            Left            =   1920
            TabIndex        =   422
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "19"
            Height          =   255
            Index           =   18
            Left            =   1440
            TabIndex        =   421
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "18"
            Height          =   255
            Index           =   17
            Left            =   960
            TabIndex        =   420
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "17"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   419
            Top             =   600
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "16"
            Height          =   255
            Index           =   15
            Left            =   7680
            TabIndex        =   418
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "15"
            Height          =   255
            Index           =   14
            Left            =   7200
            TabIndex        =   417
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "14"
            Height          =   255
            Index           =   13
            Left            =   6720
            TabIndex        =   416
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "13"
            Height          =   255
            Index           =   12
            Left            =   6240
            TabIndex        =   415
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "12"
            Height          =   255
            Index           =   11
            Left            =   5760
            TabIndex        =   414
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "11"
            Height          =   255
            Index           =   10
            Left            =   5280
            TabIndex        =   413
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "10"
            Height          =   255
            Index           =   9
            Left            =   4800
            TabIndex        =   412
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "9"
            Height          =   255
            Index           =   8
            Left            =   4320
            TabIndex        =   411
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "8"
            Height          =   255
            Index           =   7
            Left            =   3840
            TabIndex        =   410
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "7"
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   409
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   408
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "5"
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   407
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   406
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   405
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   404
            Top             =   360
            Width           =   375
         End
         Begin VB.Label bButton 
            Alignment       =   2  'Center
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   403
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "POV 3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7905
         TabIndex        =   401
         Top             =   2415
         Width           =   825
         Begin VB.Shape Shape6 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   345
            Shape           =   3  'Circle
            Top             =   540
            Width           =   150
         End
         Begin VB.Line POV3 
            BorderColor     =   &H000000FF&
            BorderWidth     =   6
            X1              =   405
            X2              =   410
            Y1              =   600
            Y2              =   605
         End
         Begin VB.Shape Shape3 
            Height          =   675
            Left            =   30
            Shape           =   3  'Circle
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "POV 2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6945
         TabIndex        =   400
         Top             =   2415
         Width           =   825
         Begin VB.Shape Shape5 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   345
            Shape           =   3  'Circle
            Top             =   540
            Width           =   150
         End
         Begin VB.Line POV2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   6
            X1              =   405
            X2              =   410
            Y1              =   600
            Y2              =   605
         End
         Begin VB.Shape Shape2 
            Height          =   675
            Left            =   30
            Shape           =   3  'Circle
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "POV 1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5940
         TabIndex        =   399
         Top             =   2415
         Width           =   825
         Begin VB.Shape Shape4 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   345
            Shape           =   3  'Circle
            Top             =   540
            Width           =   150
         End
         Begin VB.Line POV1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   6
            X1              =   405
            X2              =   410
            Y1              =   600
            Y2              =   605
         End
         Begin VB.Shape Shape1 
            Height          =   675
            Left            =   30
            Shape           =   3  'Circle
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "Slider 1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   6
         Left            =   120
         TabIndex        =   390
         Top             =   2415
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   1725
            TabIndex        =   395
            Top             =   30
            Width           =   975
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   6
            Left            =   225
            Max             =   10000
            TabIndex        =   394
            Top             =   510
            Width           =   2310
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   6
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   393
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.PictureBox PS1 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   391
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PS1 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   392
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   13
            Left            =   75
            TabIndex        =   398
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   12
            Left            =   75
            TabIndex        =   397
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   59
            Left            =   2535
            TabIndex        =   396
            Top             =   225
            Width           =   165
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   6
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "RX-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   2
         Left            =   120
         TabIndex        =   381
         Top             =   1320
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1725
            TabIndex        =   386
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PRX 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   384
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PRX 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   385
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   3
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   383
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   3
            Left            =   225
            Max             =   10000
            TabIndex        =   382
            Top             =   510
            Width           =   2310
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   2
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   58
            Left            =   2535
            TabIndex        =   389
            Top             =   225
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   388
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   387
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "X-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   0
         Left            =   120
         TabIndex        =   372
         Top             =   240
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1725
            TabIndex        =   377
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PX 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   375
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PX 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   376
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   0
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   374
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   0
            Left            =   225
            Max             =   10000
            TabIndex        =   373
            Top             =   510
            Width           =   2310
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   2490
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   0
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   57
            Left            =   2535
            TabIndex        =   380
            Top             =   225
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   379
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   378
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "Slider 2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   7
         Left            =   3030
         TabIndex        =   363
         Top             =   2415
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   1725
            TabIndex        =   368
            Top             =   30
            Width           =   975
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   7
            Left            =   225
            Max             =   10000
            TabIndex        =   367
            Top             =   510
            Width           =   2310
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   7
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   366
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.PictureBox PS2 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   364
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PS2 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   365
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   15
            Left            =   75
            TabIndex        =   371
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   14
            Left            =   75
            TabIndex        =   370
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   56
            Left            =   2535
            TabIndex        =   369
            Top             =   225
            Width           =   165
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   7
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "RY-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   4
         Left            =   3030
         TabIndex        =   354
         Top             =   1320
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1725
            TabIndex        =   359
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PRY 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   357
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PRY 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   358
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   4
            Left            =   225
            Max             =   10000
            TabIndex        =   356
            Top             =   510
            Width           =   2310
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   4
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   355
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   31
            Left            =   2535
            TabIndex        =   362
            Top             =   225
            Width           =   165
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   4
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   8
            Left            =   75
            TabIndex        =   361
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   9
            Left            =   75
            TabIndex        =   360
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "Y-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   1
         Left            =   3030
         TabIndex        =   345
         Top             =   240
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1725
            TabIndex        =   350
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PY 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   348
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PY 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   349
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   1
            Left            =   225
            Max             =   10000
            TabIndex        =   347
            Top             =   510
            Width           =   2310
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   1
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   346
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   30
            Left            =   2535
            TabIndex        =   353
            Top             =   225
            Width           =   165
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   1
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   352
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   351
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "RZ-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   5
         Left            =   5940
         TabIndex        =   336
         Top             =   1320
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   1725
            TabIndex        =   341
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PRZ 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   339
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PRZ 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   340
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   5
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   338
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   5
            Left            =   225
            Max             =   10000
            TabIndex        =   337
            Top             =   510
            Width           =   2310
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   5
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   29
            Left            =   2535
            TabIndex        =   344
            Top             =   225
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   10
            Left            =   75
            TabIndex        =   343
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   11
            Left            =   75
            TabIndex        =   342
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
      Begin VB.Frame fmX 
         Caption         =   "Z-Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   3
         Left            =   5940
         TabIndex        =   327
         Top             =   240
         Width           =   2790
         Begin VB.CheckBox full 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Range"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1725
            TabIndex        =   332
            Top             =   30
            Width           =   975
         End
         Begin VB.PictureBox PZ 
            Height          =   225
            Index           =   0
            Left            =   225
            ScaleHeight     =   165
            ScaleWidth      =   2250
            TabIndex        =   330
            Top             =   285
            Width           =   2310
            Begin VB.PictureBox PZ 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               FillColor       =   &H000000FF&
               ForeColor       =   &H000000C0&
               Height          =   225
               Index           =   1
               Left            =   1170
               ScaleHeight     =   225
               ScaleWidth      =   45
               TabIndex        =   331
               Top             =   -15
               Visible         =   0   'False
               Width           =   45
            End
         End
         Begin VB.HScrollBar dDead 
            Height          =   240
            Index           =   2
            Left            =   225
            Max             =   10000
            TabIndex        =   329
            Top             =   510
            Width           =   2310
         End
         Begin VB.HScrollBar dSaturation 
            Height          =   240
            Index           =   2
            Left            =   225
            Max             =   10000
            Min             =   2000
            TabIndex        =   328
            Top             =   780
            Value           =   10000
            Width           =   2310
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   3
            X1              =   180
            X2              =   75
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   28
            Left            =   2535
            TabIndex        =   335
            Top             =   225
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "D"
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   334
            ToolTipText     =   "Dead-Zone"
            Top             =   525
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "S"
            Height          =   195
            Index           =   5
            Left            =   75
            TabIndex        =   333
            ToolTipText     =   "Saturation"
            Top             =   790
            Width           =   105
         End
      End
   End
   Begin VB.Frame frmForce 
      Caption         =   "Force Feedback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   324
      Top             =   600
      Width           =   8895
      Begin VB.Frame Frame5 
         Caption         =   "Miscellaneous"
         Height          =   2535
         Left            =   6000
         TabIndex        =   536
         Top             =   2880
         Width           =   2655
         Begin VB.CheckBox chkPlay 
            Caption         =   "Play force until finished"
            Height          =   255
            Left            =   240
            TabIndex        =   540
            Top             =   600
            Value           =   2  'Grayed
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.HScrollBar scrOverall 
            Height          =   255
            Left            =   120
            Max             =   10000
            TabIndex        =   538
            Top             =   1560
            Value           =   10000
            Width           =   2415
         End
         Begin VB.CheckBox chkCenter 
            Caption         =   "Enable Return to Center"
            Height          =   255
            Left            =   240
            TabIndex        =   537
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Overall Strength"
            Height          =   195
            Left            =   120
            TabIndex        =   539
            Top             =   1320
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Explosions"
         Height          =   2535
         Index           =   4
         Left            =   3120
         TabIndex        =   513
         Top             =   2880
         Width           =   2655
         Begin VB.CommandButton expTest 
            Caption         =   "*"
            Height          =   315
            Left            =   2280
            TabIndex        =   533
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox dbox 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   4
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   2175
            TabIndex        =   519
            Top             =   1860
            Width           =   2175
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   7
               Left            =   1875
               TabIndex        =   527
               ToolTipText     =   "Force comes from north-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   6
               Left            =   1695
               TabIndex        =   526
               ToolTipText     =   "Force comes from north-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   5
               Left            =   1485
               TabIndex        =   525
               ToolTipText     =   "Force comes from south-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   4
               Left            =   1290
               TabIndex        =   524
               ToolTipText     =   "Force comes from south-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   3
               Left            =   960
               TabIndex        =   523
               ToolTipText     =   "Force comes from north"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   522
               ToolTipText     =   "Force comes from south"
               Top             =   270
               Value           =   -1  'True
               Width           =   195
            End
            Begin VB.OptionButton dExp 
               Height          =   210
               Index           =   0
               Left            =   225
               TabIndex        =   521
               Top             =   285
               Width           =   225
            End
            Begin VB.OptionButton dExp 
               Height          =   240
               Index           =   1
               Left            =   465
               TabIndex        =   520
               ToolTipText     =   "Force comes from west"
               Top             =   270
               Width           =   195
            End
            Begin VB.Line Line4 
               Index           =   4
               X1              =   1245
               X2              =   1245
               Y1              =   0
               Y2              =   495
            End
            Begin VB.Image Image3 
               Height          =   300
               Index           =   4
               Left            =   0
               Picture         =   "frmMain.frx":0446
               Top             =   0
               Width           =   2250
            End
         End
         Begin VB.HScrollBar scrFrequency 
            Height          =   135
            Index           =   3
            Left            =   120
            Max             =   10000
            TabIndex        =   518
            Top             =   1680
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrDuration 
            Height          =   135
            Index           =   3
            Left            =   120
            Max             =   10000
            TabIndex        =   517
            Top             =   1320
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrMagnitude 
            Height          =   135
            Index           =   3
            Left            =   120
            Max             =   10000
            TabIndex        =   516
            Top             =   960
            Value           =   10000
            Width           =   2295
         End
         Begin VB.CheckBox chkForce 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   515
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox fType 
            Height          =   315
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   514
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fequency"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   530
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Duration"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   529
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Magnitude"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   528
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Weapons"
         Height          =   2535
         Index           =   3
         Left            =   240
         TabIndex        =   495
         Top             =   2880
         Width           =   2655
         Begin VB.CommandButton wepTest 
            Caption         =   "*"
            Height          =   315
            Left            =   2280
            TabIndex        =   532
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox dbox 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   3
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   2175
            TabIndex        =   501
            Top             =   1860
            Width           =   2175
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   7
               Left            =   1875
               TabIndex        =   509
               ToolTipText     =   "Force comes from north-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   6
               Left            =   1695
               TabIndex        =   508
               ToolTipText     =   "Force comes from north-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   5
               Left            =   1485
               TabIndex        =   507
               ToolTipText     =   "Force comes from south-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   4
               Left            =   1290
               TabIndex        =   506
               ToolTipText     =   "Force comes from south-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   3
               Left            =   960
               TabIndex        =   505
               ToolTipText     =   "Force comes from north"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   504
               ToolTipText     =   "Force comes from south"
               Top             =   270
               Value           =   -1  'True
               Width           =   195
            End
            Begin VB.OptionButton dWep 
               Height          =   210
               Index           =   0
               Left            =   225
               TabIndex        =   503
               Top             =   285
               Width           =   225
            End
            Begin VB.OptionButton dWep 
               Height          =   240
               Index           =   1
               Left            =   465
               TabIndex        =   502
               ToolTipText     =   "Force comes from west"
               Top             =   270
               Width           =   195
            End
            Begin VB.Line Line4 
               Index           =   3
               X1              =   1245
               X2              =   1245
               Y1              =   0
               Y2              =   495
            End
            Begin VB.Image Image3 
               Height          =   300
               Index           =   3
               Left            =   0
               Picture         =   "frmMain.frx":0639
               Top             =   0
               Width           =   2250
            End
         End
         Begin VB.HScrollBar scrFrequency 
            Height          =   135
            Index           =   2
            Left            =   120
            Max             =   10000
            TabIndex        =   500
            Top             =   1680
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrDuration 
            Height          =   135
            Index           =   2
            Left            =   120
            Max             =   10000
            TabIndex        =   499
            Top             =   1320
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrMagnitude 
            Height          =   135
            Index           =   2
            Left            =   120
            Max             =   10000
            TabIndex        =   498
            Top             =   960
            Value           =   10000
            Width           =   2295
         End
         Begin VB.CheckBox chkForce 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   497
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox fType 
            Height          =   315
            Index           =   2
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   496
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fequency"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   512
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Duration"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   511
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Magnitude"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   510
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Steering"
         Height          =   2535
         Index           =   2
         Left            =   6000
         TabIndex        =   478
         Top             =   240
         Width           =   2655
         Begin VB.HScrollBar scrSuspension 
            Height          =   135
            Left            =   120
            Max             =   100
            TabIndex        =   534
            Top             =   1680
            Value           =   50
            Width           =   2295
         End
         Begin VB.PictureBox dbox 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   2
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   2175
            TabIndex        =   483
            Top             =   1860
            Width           =   2175
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   7
               Left            =   1875
               TabIndex        =   491
               ToolTipText     =   "Force comes from north-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   6
               Left            =   1695
               TabIndex        =   490
               ToolTipText     =   "Force comes from north-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   5
               Left            =   1485
               TabIndex        =   489
               ToolTipText     =   "Force comes from south-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   4
               Left            =   1290
               TabIndex        =   488
               ToolTipText     =   "Force comes from south-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   3
               Left            =   960
               TabIndex        =   487
               ToolTipText     =   "Force comes from north"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   486
               ToolTipText     =   "Force comes from south"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dGrip 
               Height          =   210
               Index           =   0
               Left            =   225
               TabIndex        =   485
               Top             =   285
               Value           =   -1  'True
               Width           =   225
            End
            Begin VB.OptionButton dGrip 
               Height          =   240
               Index           =   1
               Left            =   465
               TabIndex        =   484
               ToolTipText     =   "Force comes from west"
               Top             =   270
               Width           =   195
            End
            Begin VB.Line Line4 
               Index           =   2
               X1              =   1245
               X2              =   1245
               Y1              =   0
               Y2              =   495
            End
            Begin VB.Image Image3 
               Height          =   300
               Index           =   2
               Left            =   0
               Picture         =   "frmMain.frx":082C
               Top             =   0
               Width           =   2250
            End
         End
         Begin VB.HScrollBar scrFriction 
            Height          =   135
            Left            =   120
            Max             =   10000
            TabIndex        =   482
            Top             =   1320
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrSpring 
            Height          =   135
            Left            =   120
            Max             =   300
            TabIndex        =   481
            Top             =   960
            Value           =   100
            Width           =   2295
         End
         Begin VB.HScrollBar scrGrip 
            Height          =   135
            Left            =   120
            Max             =   10000
            TabIndex        =   480
            Top             =   600
            Value           =   5000
            Width           =   2295
         End
         Begin VB.CheckBox chkSteering 
            Caption         =   "Enabled"
            Height          =   195
            Left            =   1320
            TabIndex        =   479
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Suspension Influence"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   535
            Top             =   1440
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Friction"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   494
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Spring Multiplier"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   493
            Top             =   720
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Grip Multiplier (Constant Force)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   492
            Top             =   360
            Width           =   2160
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Suspension"
         Height          =   2535
         Index           =   1
         Left            =   3120
         TabIndex        =   460
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton susTest 
            Caption         =   "*"
            Height          =   315
            Left            =   2280
            TabIndex        =   531
            Top             =   360
            Width           =   255
         End
         Begin VB.ComboBox fType 
            Height          =   315
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   474
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkForce 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   473
            Top             =   0
            Width           =   1095
         End
         Begin VB.HScrollBar scrMagnitude 
            Height          =   135
            Index           =   1
            Left            =   120
            Max             =   10000
            TabIndex        =   472
            Top             =   960
            Value           =   10000
            Width           =   2295
         End
         Begin VB.HScrollBar scrDuration 
            Height          =   135
            Index           =   1
            Left            =   120
            Max             =   10000
            TabIndex        =   471
            Top             =   1320
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrFrequency 
            Height          =   135
            Index           =   1
            Left            =   120
            Max             =   10000
            TabIndex        =   470
            Top             =   1680
            Value           =   2000
            Width           =   2295
         End
         Begin VB.PictureBox dbox 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   1
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   2175
            TabIndex        =   461
            Top             =   1860
            Width           =   2175
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   1
               Left            =   465
               TabIndex        =   469
               ToolTipText     =   "Force comes from west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   210
               Index           =   0
               Left            =   225
               TabIndex        =   468
               Top             =   285
               Width           =   225
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   467
               ToolTipText     =   "Force comes from south"
               Top             =   270
               Value           =   -1  'True
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   3
               Left            =   960
               TabIndex        =   466
               ToolTipText     =   "Force comes from north"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   4
               Left            =   1290
               TabIndex        =   465
               ToolTipText     =   "Force comes from south-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   5
               Left            =   1485
               TabIndex        =   464
               ToolTipText     =   "Force comes from south-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   6
               Left            =   1695
               TabIndex        =   463
               ToolTipText     =   "Force comes from north-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dBump 
               Height          =   240
               Index           =   7
               Left            =   1875
               TabIndex        =   462
               ToolTipText     =   "Force comes from north-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.Image Image3 
               Height          =   300
               Index           =   1
               Left            =   0
               Picture         =   "frmMain.frx":0A1F
               Top             =   0
               Width           =   2250
            End
            Begin VB.Line Line4 
               Index           =   1
               X1              =   1245
               X2              =   1245
               Y1              =   0
               Y2              =   495
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Magnitude"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   477
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Duration"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   476
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fequency"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   475
            Top             =   1440
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Collision"
         Height          =   2535
         Index           =   0
         Left            =   240
         TabIndex        =   442
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton colTest 
            Caption         =   "*"
            Height          =   315
            Left            =   2280
            TabIndex        =   543
            Top             =   360
            Width           =   255
         End
         Begin VB.PictureBox dbox 
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   0
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   2415
            TabIndex        =   451
            Top             =   1860
            Width           =   2415
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   7
               Left            =   1875
               TabIndex        =   459
               ToolTipText     =   "Force comes from north-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   6
               Left            =   1695
               TabIndex        =   458
               ToolTipText     =   "Force comes from north-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   5
               Left            =   1485
               TabIndex        =   457
               ToolTipText     =   "Force comes from south-west"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   4
               Left            =   1290
               TabIndex        =   456
               ToolTipText     =   "Force comes from south-east"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   3
               Left            =   960
               TabIndex        =   455
               ToolTipText     =   "Force comes from north"
               Top             =   270
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   2
               Left            =   720
               TabIndex        =   454
               ToolTipText     =   "Force comes from south"
               Top             =   270
               Value           =   -1  'True
               Width           =   195
            End
            Begin VB.OptionButton dCol 
               Height          =   210
               Index           =   0
               Left            =   225
               TabIndex        =   453
               Top             =   285
               Width           =   225
            End
            Begin VB.OptionButton dCol 
               Height          =   240
               Index           =   1
               Left            =   465
               TabIndex        =   452
               ToolTipText     =   "Force comes from west"
               Top             =   270
               Width           =   195
            End
            Begin VB.Line Line4 
               Index           =   0
               X1              =   1245
               X2              =   1245
               Y1              =   0
               Y2              =   495
            End
            Begin VB.Image Image3 
               Height          =   300
               Index           =   0
               Left            =   0
               Picture         =   "frmMain.frx":0C12
               Top             =   0
               Width           =   2250
            End
         End
         Begin VB.HScrollBar scrFrequency 
            Height          =   135
            Index           =   0
            Left            =   120
            Max             =   10000
            TabIndex        =   447
            Top             =   1680
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrDuration 
            Height          =   135
            Index           =   0
            Left            =   120
            Max             =   10000
            TabIndex        =   446
            Top             =   1320
            Value           =   2000
            Width           =   2295
         End
         Begin VB.HScrollBar scrMagnitude 
            Height          =   135
            Index           =   0
            Left            =   120
            Max             =   10000
            TabIndex        =   445
            Top             =   960
            Value           =   10000
            Width           =   2295
         End
         Begin VB.CheckBox chkForce 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   444
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox fType 
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   443
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fequency"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   450
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Duration"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   449
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Magnitude"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   448
            Top             =   720
            Width           =   750
         End
      End
   End
   Begin VB.Frame frmVehicle 
      Caption         =   "Vehicle Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   176
      Top             =   600
      Width           =   8895
      Begin VB.VScrollBar fScroll2 
         Height          =   4815
         Left            =   7080
         Max             =   8
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox sBox2 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   600
         ScaleHeight     =   4815
         ScaleWidth      =   7095
         TabIndex        =   178
         Top             =   360
         Width           =   7095
         Begin VB.PictureBox mBox2 
            BorderStyle     =   0  'None
            Height          =   9015
            Left            =   -120
            ScaleHeight     =   9015
            ScaleWidth      =   6735
            TabIndex        =   179
            Top             =   0
            Width           =   6735
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   52
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   436
               Top             =   8640
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   52
                  Left            =   1800
                  TabIndex        =   440
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   52
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   439
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   52
                  Left            =   4080
                  TabIndex        =   438
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   52
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   437
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   51
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   319
               Top             =   8280
               Width           =   4455
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   51
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   323
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   51
                  Left            =   4080
                  TabIndex        =   322
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   51
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   321
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   51
                  Left            =   1800
                  TabIndex        =   320
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   50
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   314
               Top             =   7920
               Width           =   4455
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   50
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   318
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   50
                  Left            =   4080
                  TabIndex        =   317
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   50
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   316
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   50
                  Left            =   1800
                  TabIndex        =   315
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   49
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   309
               Top             =   7560
               Width           =   4455
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   49
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   313
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   49
                  Left            =   4080
                  TabIndex        =   312
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   49
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   311
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   49
                  Left            =   1800
                  TabIndex        =   310
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   48
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   304
               Top             =   7200
               Width           =   4455
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   48
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   308
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   48
                  Left            =   4080
                  TabIndex        =   307
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   48
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   306
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   48
                  Left            =   1800
                  TabIndex        =   305
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   47
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   299
               Top             =   6840
               Width           =   4455
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   47
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   303
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   47
                  Left            =   4080
                  TabIndex        =   302
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   47
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   301
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   47
                  Left            =   1800
                  TabIndex        =   300
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   46
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   294
               Top             =   6480
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   46
                  Left            =   1800
                  TabIndex        =   298
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   46
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   297
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   46
                  Left            =   4080
                  TabIndex        =   296
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   46
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   295
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   45
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   289
               Top             =   6120
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   45
                  Left            =   1800
                  TabIndex        =   293
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   45
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   292
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   45
                  Left            =   4080
                  TabIndex        =   291
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   45
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   290
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   44
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   284
               Top             =   5760
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   44
                  Left            =   1800
                  TabIndex        =   288
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   44
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   287
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   44
                  Left            =   4080
                  TabIndex        =   286
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   44
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   285
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   43
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   279
               Top             =   5400
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   43
                  Left            =   1800
                  TabIndex        =   283
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   43
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   282
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   43
                  Left            =   4080
                  TabIndex        =   281
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   43
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   280
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   42
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   274
               Top             =   5040
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   42
                  Left            =   1800
                  TabIndex        =   278
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   42
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   277
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   42
                  Left            =   4080
                  TabIndex        =   276
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   42
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   275
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   41
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   269
               Top             =   4680
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   41
                  Left            =   1800
                  TabIndex        =   273
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   41
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   272
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   41
                  Left            =   4080
                  TabIndex        =   271
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   41
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   270
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   40
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   264
               Top             =   4320
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   40
                  Left            =   1800
                  TabIndex        =   268
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   40
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   267
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   40
                  Left            =   4080
                  TabIndex        =   266
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   40
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   265
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   39
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   259
               Top             =   3960
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   39
                  Left            =   1800
                  TabIndex        =   263
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   39
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   262
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   39
                  Left            =   4080
                  TabIndex        =   261
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   39
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   260
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   38
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   254
               Top             =   3600
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   38
                  Left            =   1800
                  TabIndex        =   258
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   38
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   257
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   38
                  Left            =   4080
                  TabIndex        =   256
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   38
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   255
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   37
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   249
               Top             =   3240
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   37
                  Left            =   1800
                  TabIndex        =   253
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   37
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   252
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   37
                  Left            =   4080
                  TabIndex        =   251
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   37
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   250
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   36
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   244
               Top             =   2880
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   36
                  Left            =   1800
                  TabIndex        =   248
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   36
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   247
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   36
                  Left            =   4080
                  TabIndex        =   246
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   36
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   245
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   35
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   239
               Top             =   2520
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   35
                  Left            =   1800
                  TabIndex        =   243
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   35
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   242
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   35
                  Left            =   4080
                  TabIndex        =   241
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   35
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   240
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   34
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   234
               Top             =   2160
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   34
                  Left            =   1800
                  TabIndex        =   238
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   34
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   237
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   34
                  Left            =   4080
                  TabIndex        =   236
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   34
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   235
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   33
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   229
               Top             =   1800
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   33
                  Left            =   1800
                  TabIndex        =   233
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   33
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   232
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   33
                  Left            =   4080
                  TabIndex        =   231
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   33
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   230
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   32
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   224
               Top             =   1440
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   32
                  Left            =   1800
                  TabIndex        =   228
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   32
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   227
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   32
                  Left            =   4080
                  TabIndex        =   226
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   32
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   225
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   31
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   219
               Top             =   1080
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   31
                  Left            =   1800
                  TabIndex        =   223
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   31
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   222
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   31
                  Left            =   4080
                  TabIndex        =   221
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   31
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   220
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   30
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   214
               Top             =   720
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   30
                  Left            =   1800
                  TabIndex        =   218
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   30
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   217
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   30
                  Left            =   4080
                  TabIndex        =   216
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   30
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   215
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   29
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   209
               Top             =   360
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   29
                  Left            =   1800
                  TabIndex        =   213
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   29
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   212
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   29
                  Left            =   4080
                  TabIndex        =   211
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   29
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   210
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.PictureBox ctlBox 
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   28
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   4455
               TabIndex        =   180
               Top             =   0
               Width           =   4455
               Begin VB.CommandButton cmdPrimary 
                  Height          =   285
                  Index           =   28
                  Left            =   1800
                  TabIndex        =   182
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtPrimary 
                  Height          =   285
                  Index           =   28
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   181
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSecondary 
                  Height          =   285
                  Index           =   28
                  Left            =   4080
                  TabIndex        =   184
                  Top             =   0
                  Width           =   255
               End
               Begin VB.TextBox txtSecondary 
                  Height          =   285
                  Index           =   28
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   183
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Special Ctrl Down"
               Height          =   195
               Index           =   60
               Left            =   0
               TabIndex        =   441
               Top             =   8640
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Fire"
               Height          =   195
               Index           =   55
               Left            =   0
               TabIndex        =   208
               Top             =   45
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Secondary Fire"
               Height          =   195
               Index           =   54
               Left            =   0
               TabIndex        =   207
               Top             =   405
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Accelerate"
               Height          =   195
               Index           =   53
               Left            =   0
               TabIndex        =   206
               Top             =   765
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Brake/Reverse"
               Height          =   195
               Index           =   52
               Left            =   0
               TabIndex        =   205
               Top             =   1125
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Left"
               Height          =   195
               Index           =   51
               Left            =   0
               TabIndex        =   204
               Top             =   1485
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Right"
               Height          =   195
               Index           =   50
               Left            =   0
               TabIndex        =   203
               Top             =   1845
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Steer Forward/Down"
               Height          =   195
               Index           =   49
               Left            =   0
               TabIndex        =   202
               Top             =   2205
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Steer Back/Up"
               Height          =   195
               Index           =   48
               Left            =   0
               TabIndex        =   201
               Top             =   2565
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Enter + Exit"
               Height          =   195
               Index           =   47
               Left            =   0
               TabIndex        =   200
               Top             =   2925
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Trip Skip"
               Height          =   195
               Index           =   46
               Left            =   0
               TabIndex        =   199
               Top             =   3285
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Next Radio Station"
               Height          =   195
               Index           =   45
               Left            =   0
               TabIndex        =   198
               Top             =   3645
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Previous Radio Station"
               Height          =   195
               Index           =   44
               Left            =   0
               TabIndex        =   197
               Top             =   4005
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "User Track Skip"
               Height          =   195
               Index           =   43
               Left            =   0
               TabIndex        =   196
               Top             =   4365
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Horn"
               Height          =   195
               Index           =   42
               Left            =   0
               TabIndex        =   195
               Top             =   4725
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Sub-mission"
               Height          =   195
               Index           =   41
               Left            =   0
               TabIndex        =   194
               Top             =   5085
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Change Camera"
               Height          =   195
               Index           =   40
               Left            =   0
               TabIndex        =   193
               Top             =   5445
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Handbrake"
               Height          =   195
               Index           =   39
               Left            =   0
               TabIndex        =   192
               Top             =   5805
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Behind"
               Height          =   195
               Index           =   38
               Left            =   0
               TabIndex        =   191
               Top             =   6165
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Mouse Look"
               Height          =   195
               Index           =   37
               Left            =   0
               TabIndex        =   190
               Top             =   6525
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Left"
               Height          =   195
               Index           =   36
               Left            =   0
               TabIndex        =   189
               Top             =   6885
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Look Right"
               Height          =   195
               Index           =   35
               Left            =   0
               TabIndex        =   188
               Top             =   7245
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Special Ctrl Left"
               Height          =   195
               Index           =   34
               Left            =   0
               TabIndex        =   187
               Top             =   7605
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Special Ctrl Right"
               Height          =   195
               Index           =   33
               Left            =   0
               TabIndex        =   186
               Top             =   7965
               Width           =   1920
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Special Ctrl Up"
               Height          =   195
               Index           =   32
               Left            =   0
               TabIndex        =   185
               Top             =   8325
               Width           =   1920
            End
         End
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Racer_S"
      Height          =   195
      Left            =   7800
      TabIndex        =   547
      Top             =   6360
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ToCAEDIT.COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7800
      TabIndex        =   546
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Presets"
      Height          =   195
      Left            =   120
      TabIndex        =   545
      Top             =   6240
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Device"
      Height          =   195
      Left            =   5520
      TabIndex        =   544
      Top             =   0
      Width           =   510
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Unbind"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Manual Select"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements DirectXEvent

Private Sub bDelete_Click()
On Error Resume Next
dSettings preset.List(preset.ListIndex)
End Sub

Private Sub bSave_Click()
On Error Resume Next
sSettings preset.Text
End Sub

Private Sub chkCenter_Click()
On Error Resume Next
SetRTCGAIN CBool(chkCenter.value), CLng(scrOverall.value)
End Sub



Private Sub chkSteering_Click()
On Error Resume Next
If chkSteering.value = Unchecked Then StopSteer
End Sub

Private Sub cmdPrimary_Click(Index As Integer)
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
bListen = False

ctlIndex = Index
ctlSecond = False
frmMain.PopupMenu mnuPop
End Sub

Private Sub cmdSecondary_Click(Index As Integer)
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
bListen = False


ctlIndex = Index
ctlSecond = True
frmMain.PopupMenu mnuPop

End Sub





Private Sub cTest_Click()
On Error Resume Next

Call StartEffect(fType(0).ListIndex, scrMagnitude(0).value, scrDuration(0).value, scrFrequency(0).value, dCol)

End Sub


Private Sub colTest_Click()
On Error Resume Next
Call StartEffect(fType(0).ListIndex, scrMagnitude(0).value, scrDuration(0).value, scrFrequency(0).value, dCol)

End Sub

Private Sub ctlBox_GotFocus(Index As Integer)
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
mBox.SetFocus
bListen = False
End Sub

Private Sub dDead_Change(Index As Integer)
On Error Resume Next
    With DiProp_Dead
        .lData = dDead(Index).value
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Dead)
        Select Case Index
        Case 0
        .lObj = DIJOFS_X
        
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        
        Case 1
        .lObj = DIJOFS_Y
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 2
        .lObj = DIJOFS_Z
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 3
        .lObj = DIJOFS_RX
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 4
        .lObj = DIJOFS_RY
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 5
        .lObj = DIJOFS_RZ
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 6
        .lObj = DIJOFS_SLIDER0
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        Case 7
        .lObj = DIJOFS_SLIDER1
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        End Select
    End With

End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
On Error Resume Next



'fill the OLD one
If Not oinCar = inCar Then
For jLoop = 0 To 307
olrawInputs(jLoop) = rawInputs(jLoop) - 1
ozrawInputs(jLoop) = zrawInputs(jLoop) - 1
Next jLoop

Else
For jLoop = 0 To 307
olrawInputs(jLoop) = rawInputs(jLoop)
ozrawInputs(jLoop) = zrawInputs(jLoop)
Next jLoop

End If
oinCar = inCar
If Not objDIDevC Is Nothing Then
'interpret axes
If full(0) = Unchecked Then rawInputs(270) = js.X - 255 Else rawInputs(270) = 255 - (js.X / 2)       'x
If full(0) = Unchecked Then rawInputs(271) = 255 - js.X Else rawInputs(271) = (js.X / 2)             'x
If AxisPresent(1) = False Then rawInputs(270) = 0: rawInputs(271) = 0
If full(1) = Unchecked Then rawInputs(272) = js.Y - 255 Else rawInputs(272) = 255 - (js.Y / 2)       'y
If full(1) = Unchecked Then rawInputs(273) = 255 - js.Y Else rawInputs(273) = (js.Y / 2)             'y
If AxisPresent(2) = False Then rawInputs(272) = 0: rawInputs(273) = 0
If full(2) = Unchecked Then rawInputs(274) = js.z - 255 Else rawInputs(274) = 255 - (js.z / 2)       'z
If full(2) = Unchecked Then rawInputs(275) = 255 - js.z Else rawInputs(275) = (js.z / 2)             'z
If AxisPresent(3) = False Then rawInputs(274) = 0: rawInputs(275) = 0

'interpret R axes
If full(3) = Unchecked Then rawInputs(276) = js.rx - 255 Else rawInputs(276) = 255 - (js.rx / 2)     'rx
If full(3) = Unchecked Then rawInputs(277) = 255 - js.rx Else rawInputs(277) = (js.rx / 2)           'rx
If AxisPresent(4) = False Then rawInputs(276) = 0: rawInputs(277) = 0
If full(4) = Unchecked Then rawInputs(278) = js.ry - 255 Else rawInputs(278) = 255 - (js.ry / 2)     'ry
If full(4) = Unchecked Then rawInputs(279) = 255 - js.ry Else rawInputs(279) = (js.ry / 2)           'ry
If AxisPresent(5) = False Then rawInputs(278) = 0: rawInputs(279) = 0
If full(5) = Unchecked Then rawInputs(280) = js.rz - 255 Else rawInputs(280) = 255 - (js.rz / 2)   'rz
If full(5) = Unchecked Then rawInputs(281) = 255 - js.rz Else rawInputs(281) = (js.rz / 2)         'rz
If AxisPresent(6) = False Then rawInputs(280) = 0: rawInputs(281) = 0

'interpret sliders
If full(6) = Unchecked Then rawInputs(282) = js.slider(0) - 255 Else rawInputs(282) = 255 - (js.slider(0) / 2)  'slider1
If full(6) = Unchecked Then rawInputs(283) = 255 - js.slider(0) Else rawInputs(283) = (js.slider(0) / 2)        'slider1
If AxisPresent(7) = False Then rawInputs(282) = 0: rawInputs(283) = 0
If full(7) = Unchecked Then rawInputs(284) = js.slider(1) - 255 Else rawInputs(284) = 255 - (js.slider(1) / 2)  'slider2
If full(7) = Unchecked Then rawInputs(285) = 255 - js.slider(1) Else rawInputs(285) = (js.slider(1) / 2)        'slider2
If AxisPresent(8) = False Then rawInputs(284) = 0: rawInputs(285) = 0

'interpret POV
InterpretPOV js.POV(0), rawInputs(286), rawInputs(287), rawInputs(288), rawInputs(289)
InterpretPOV js.POV(1), rawInputs(290), rawInputs(291), rawInputs(292), rawInputs(293)
InterpretPOV js.POV(2), rawInputs(294), rawInputs(295), rawInputs(296), rawInputs(297)

'interpret buttons

For jLoop = 0 To 31
If js.buttons(jLoop) = 128 Then rawInputs(jLoop + 238) = 255 Else rawInputs(jLoop + 238) = 0
Next jLoop



If testing Then

'X
If AxisPresent(1) And full(0) = Unchecked Then
SetPos PX, js.X
ElseIf AxisPresent(1) And full(0) = Checked Then
SetPos PX, js.X, 0, 1
End If

'Y
If AxisPresent(2) And full(1) = Unchecked Then
SetPos PY, js.Y
ElseIf AxisPresent(2) And full(1) = Checked Then
SetPos PY, js.Y, 0, 1
End If

'Z
If AxisPresent(3) And full(2) = Unchecked Then
SetPos PZ, js.z
ElseIf AxisPresent(3) And full(2) = Checked Then
SetPos PZ, js.z, 0, 1
End If

'RX
If AxisPresent(4) And full(3) = Unchecked Then
SetPos PRX, js.rx
ElseIf AxisPresent(4) And full(3) = Checked Then
SetPos PRX, js.rx, 0, 1
End If
'RY
If AxisPresent(5) And full(4) = Unchecked Then
SetPos PRY, js.ry
ElseIf AxisPresent(5) And full(4) = Checked Then
SetPos PRY, js.ry, 0, 1
End If

'RZ
If AxisPresent(6) And full(5) = Unchecked Then
SetPos PRZ, js.rz
ElseIf AxisPresent(6) And full(5) = Checked Then
SetPos PRZ, js.rz, 0, 1
End If

If AxisPresent(7) And full(6) = Unchecked Then
SetPos PS1, js.slider(0)
ElseIf AxisPresent(7) And full(6) = Checked Then
SetPos PS1, js.slider(0), 0, 1
End If

If AxisPresent(8) And full(7) = Unchecked Then
SetPos PS2, 510 - js.slider(1)
ElseIf AxisPresent(8) And full(7) = Checked Then
SetPos PS2, 510 - js.slider(1), 0, 1
End If

SetPOV POV1, js.POV(0)
SetPOV POV2, js.POV(1)
SetPOV POV3, js.POV(2)

For jLoop = 0 To 31
If js.buttons(jLoop) = 128 Then bButton(jLoop).BorderStyle = 1 Else bButton(jLoop).BorderStyle = 0
Next jLoop
End If
End If

'interpret keyboard
objDIDevK.GetDeviceStateKeyboard diState
For jLoop = 1 To 237
If diState.Key(jLoop) = 128 Then rawInputs(jLoop) = 255 Else rawInputs(jLoop) = 0
Next jLoop

'interpret mouse
        'rawInputs(307) = 0
       'rawInputs(306) = 0
'NumItems = objDIDev.GetDeviceData(diDeviceData, 0)


'        If diDeviceData(DIMOFS_BUTTON0).lData > 0 Then
 '       rawInputs(298) = 255
  '      ElseIf diDeviceData(DIMOFS_BUTTON0).lData = 0 Then
   '     rawInputs(298) = 0
    '    End If

objDIDev.GetDeviceStateMouse mState
'For jLoop = 1 To 7
If (mState.X <> 0 Or mState.Y <> 0) Then
If chkSec.value = Unchecked Then SetByte &HB6EC2E, 1
End If


        If mState.z > 0 Then
        rawInputs(306) = 255
        rawInputs(307) = 0
        ElseIf mState.z < 0 Then
        rawInputs(307) = 255
        rawInputs(306) = 0
        Else
        rawInputs(307) = 0
        rawInputs(306) = 0
        End If

        If mState.buttons(0) > 0 Then
        rawInputs(298) = 255
        ElseIf mState.buttons(0) = 0 Then
        rawInputs(298) = 0
        End If

        If mState.buttons(1) > 0 Then
        rawInputs(299) = 255
        ElseIf mState.buttons(1) = 0 Then
        rawInputs(299) = 0
        End If

        If mState.buttons(2) > 0 Then
        rawInputs(300) = 255
        ElseIf mState.buttons(2) = 0 Then
        rawInputs(300) = 0
        End If

        If mState.buttons(3) > 0 Then
        rawInputs(301) = 255
        ElseIf mState.buttons(3) = 0 Then
        rawInputs(301) = 0
        End If

 ' Next jLoop
  
For jLoop = 0 To 307
zrawInputs(jLoop) = rawInputs(jLoop)
Next jLoop
  

cPoint = &HB73458
If chkSec.value = Unchecked Then
cPoint2 = &H53F6E0
Else
cPoint2 = &H53F6A0
End If
'cPoint = &HB73604
WriteControls False

'cPoint = &HB734D0
'WriteControls False
'cPoint = &HB73488
'WriteControls True
'Pause 100




End Sub

Private Sub dSaturation_Change(Index As Integer)
On Error Resume Next
    With DiProp_Saturation
        .lData = dSaturation(Index).value
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Saturation)
        Select Case Index
        Case 0
        .lObj = DIJOFS_X
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 1
        .lObj = DIJOFS_Y
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 2
        .lObj = DIJOFS_Z
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 3
        .lObj = DIJOFS_RX
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 4
        .lObj = DIJOFS_RY
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 5
        .lObj = DIJOFS_RZ
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 6
        .lObj = DIJOFS_SLIDER0
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        Case 7
        .lObj = DIJOFS_SLIDER1
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        End Select
    End With

End Sub

Private Sub ETimer_Timer()
If KeyDown(&H37) Then
PollTimer.Enabled = True
ElseIf KeyDown(&HB5) Then
PollTimer.Enabled = False
didInit = False
DeInitGame
End If


End Sub

Private Sub expTest_Click()
On Error Resume Next
Call StartEffect(fType(3).ListIndex, scrMagnitude(3).value, scrDuration(3).value, scrFrequency(3).value, dExp)
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If

bListen = False
End Sub

Public Function lPresets()
On Error Resume Next
If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\presets.ini"
Else
strMySystemFile = App.Path & "presets.ini"
End If

Dim t As Long
Dim pNum As Long
Dim pName As String


pNum = ReadFromFile("Main", "pNum")
For t = 1 To pNum
pName = ReadFromFile("Main", "pName" & t)
preset.AddItem pName
Next t
End Function
Public Function sSettings(sName As String)
On Error Resume Next
If Len(sName) = 0 Then MsgBox "You must enter a name.", vbOKOnly, "SAAC": Exit Function

If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\presets.ini"
Else
strMySystemFile = App.Path & "presets.ini"
End If

Dim doesE As Boolean
Dim iloop As Long
For iloop = 0 To preset.ListCount
If preset.List(iloop) = sName Then doesE = True
Next iloop

If doesE = True Then GoTo jSave
preset.AddItem sName
Dim tNum As Long
tNum = ReadFromFile("Main", "pNum")
WriteToFile "Main", "pNum", (tNum + 1)
WriteToFile "Main", "pName" & (tNum + 1), sName

jSave:
For iloop = 0 To full.UBound
WriteToFile sName, "fA" & iloop, full(iloop).value
Next iloop

For iloop = 0 To dDead.UBound
WriteToFile sName, "dD" & iloop, dDead(iloop).value
WriteToFile sName, "dS" & iloop, dSaturation(iloop).value
Next iloop

For iloop = 0 To txtPrimary.UBound
WriteToFile sName, "aP" & iloop, CStr(ctlKeys(0, iloop))
WriteToFile sName, "aS" & iloop, CStr(ctlKeys(1, iloop))
Next iloop

For iloop = 0 To chkForce.UBound
WriteToFile sName, "fE" & iloop, CStr(chkForce(iloop).value)
WriteToFile sName, "fT" & iloop, CStr(fType(iloop).ListIndex)
WriteToFile sName, "fM" & iloop, CStr(scrMagnitude(iloop).value)
WriteToFile sName, "fD" & iloop, CStr(scrDuration(iloop).value)
WriteToFile sName, "fF" & iloop, CStr(scrFrequency(iloop).value)
Debug.Print findDirection(dCol)
Select Case iloop
Case 0: WriteToFile sName, "fDIR" & iloop, CStr(findDirection(dCol))
Case 1: WriteToFile sName, "fDIR" & iloop, CStr(findDirection(dBump))
Case 2: WriteToFile sName, "fDIR" & iloop, CStr(findDirection(dWep))
Case 3: WriteToFile sName, "fDIR" & iloop, CStr(findDirection(dExp))

End Select

Next iloop

WriteToFile sName, "fO", CStr(scrOverall.value)
WriteToFile sName, "fC", CStr(chkCenter.value)

WriteToFile sName, "fSE", CStr(chkSteering.value)
WriteToFile sName, "fSG", CStr(scrGrip.value)
WriteToFile sName, "fSS", CStr(scrSpring.value)
WriteToFile sName, "fSF", CStr(scrFriction.value)
WriteToFile sName, "fSI", CStr(scrSuspension.value)
WriteToFile sName, "fSDIR", CStr(findDirection(dGrip))
WriteToFile sName, "aM", CStr(aMode)


End Function
Public Function lSettings(sName As String)
On Error Resume Next

If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\presets.ini"
Else
strMySystemFile = App.Path & "presets.ini"
End If

Dim iloop As Long
For iloop = 0 To full.UBound
full(iloop).value = ReadFromFile(sName, "fA" & iloop)
Next iloop
For iloop = 0 To dDead.UBound
dDead(iloop).value = ReadFromFile(sName, "dD" & iloop)
dSaturation(iloop).value = ReadFromFile(sName, "dS" & iloop)
Next iloop
For iloop = 0 To txtPrimary.UBound
ctlKeys(0, iloop) = ReadFromFile(sName, "aP" & iloop)
ctlKeys(1, iloop) = ReadFromFile(sName, "aS" & iloop)
Next iloop


For iloop = 0 To chkForce.UBound
chkForce(iloop).value = ReadFromFile(sName, "fE" & iloop)
fType(iloop).ListIndex = ReadFromFile(sName, "fT" & iloop)
scrMagnitude(iloop).value = ReadFromFile(sName, "fM" & iloop)
scrDuration(iloop).value = ReadFromFile(sName, "fD" & iloop)
scrFrequency(iloop).value = ReadFromFile(sName, "fF" & iloop)

Select Case iloop
Case 0: dCol(ReadFromFile(sName, "fDIR" & iloop)).value = True
Case 1: dBump(ReadFromFile(sName, "fDIR" & iloop)).value = True
Case 2: dWep(ReadFromFile(sName, "fDIR" & iloop)).value = True
Case 3: dExp(ReadFromFile(sName, "fDIR" & iloop)).value = True
End Select
Next iloop

scrOverall.value = ReadFromFile(sName, "fO")
chkCenter.value = ReadFromFile(sName, "fC")

chkSteering.value = ReadFromFile(sName, "fSE")
scrGrip.value = ReadFromFile(sName, "fSG")
scrSpring.value = ReadFromFile(sName, "fSS")
scrFriction.value = ReadFromFile(sName, "fSF")
scrSuspension.value = ReadFromFile(sName, "fSI")
dGrip(ReadFromFile(sName, "fSDIR")).value = True
aMode = ReadFromFile(sName, "aM")
optAim(aMode).value = True
SetAim aMode

lText

End Function
Public Function dSettings(sName As String)
On Error Resume Next
Dim pNum As Long
Dim iloop As Long
Dim iNum As Long
Dim dName() As String
Dim sE As Boolean

If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\presets.ini"
Else
strMySystemFile = App.Path & "presets.ini"
End If


For iloop = 0 To preset.ListCount
If preset.List(iloop) = sName Then sE = True
Next iloop
If sE = False Then GoTo nothere

pNum = ReadFromFile("Main", "pNum")
iNum = pNum
WriteToFile "Main", "pNum", (pNum - 1)
ReDim dName(1 To pNum)
For iloop = 1 To pNum
If ReadFromFile("Main", "pName" & iloop) = sName Then iNum = iloop
dName(iloop) = ReadFromFile("Main", "pName" & iloop)
Next iloop

For iloop = 1 To pNum
DeleteFromFile "Main", "pName" & iloop
Next iloop

For iloop = 0 To full.UBound
DeleteFromFile sName, "fA" & iloop
Next iloop

For iloop = 0 To dDead.UBound
DeleteFromFile sName, "dD" & iloop
DeleteFromFile sName, "dS" & iloop
Next iloop

For iloop = 0 To txtPrimary.UBound
DeleteFromFile sName, "aP" & iloop
DeleteFromFile sName, "aS" & iloop
Next iloop

For iloop = 0 To chkForce.UBound
DeleteFromFile sName, "fE" & iloop
DeleteFromFile sName, "fT" & iloop
DeleteFromFile sName, "fM" & iloop
DeleteFromFile sName, "fD" & iloop
DeleteFromFile sName, "fF" & iloop
DeleteFromFile sName, "fDIR" & iloop
Next iloop

DeleteFromFile sName, "fO"
DeleteFromFile sName, "fC"
DeleteFromFile sName, "fSE"
DeleteFromFile sName, "fSG"
DeleteFromFile sName, "fSS"
DeleteFromFile sName, "fSF"
DeleteFromFile sName, "fSI"
DeleteFromFile sName, "fSDIR"
DeleteFromFile sName, "aM"


If pNum - 1 = 0 Then GoTo sre
For iloop = 1 To (pNum - 1)
If iloop >= iNum Then
WriteToFile "Main", "pName" & iloop, dName(iloop + 1)
Else
WriteToFile "Main", "pName" & iloop, dName(iloop)
End If
Next iloop
sre:
Dim myD As String
Dim free
free = FreeFile
myD = String(FileLen(strMySystemFile), 0)
Open strMySystemFile For Binary As #free
Get #free, 1, myD
Close #free
myD = Replace(myD, "[" & sName & "]", "")
Open strMySystemFile For Output As #free
Print #free, myD
Close #free
nothere:
preset.RemoveItem preset.ListIndex

End Function


Private Sub Form_Load()
On Error Resume Next
'MsgBox "set colors"
SetToDefaultSysColors frmMain, ""

'MsgBox "load presets"
lPresets

'MsgBox "init"
InitDirectInput

'MsgBox "add sticks"
Dim i As Long
For i = 1 To JoyCount
jStick.AddItem Joysticks(i)
Next i

'MsgBox "set stick"

If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\SAAC.ini"
Else
strMySystemFile = App.Path & "SAAC.ini"
End If

joyIndex = ReadFromFile("Main", "LastJoy")
preIndex = ReadFromFile("Main", "LastPre")

jStick.ListIndex = (joyIndex - 1)
preset.ListIndex = (preIndex - 1)
'MsgBox "is testing"
lText
testing = True
'MsgBox "show form"
'frmDebug.Show







End Sub

Private Sub Form_Resize()
On Error Resume Next
fScroll.Max = ((mBox.Height - sBox.Height) / sBox.Height) * 15
fScroll2.Max = ((mBox2.Height - sBox2.Height) / sBox2.Height) * 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
DeInit


If Len(App.Path) > 3 Then
strMySystemFile = App.Path & "\SAAC.ini"
Else
strMySystemFile = App.Path & "SAAC.ini"
End If

Call WriteToFile("Main", "LastJoy", CStr(jStick.ListIndex + 1))
Call WriteToFile("Main", "LastPre", CStr(preset.ListIndex + 1))


End Sub

Private Sub frmFoot_Click()
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
mBox.SetFocus
bListen = False
End Sub

Private Sub fScroll_Change()
On Error Resume Next
mBox.Top = -(((mBox.Height - sBox.Height) / fScroll.Max) * fScroll.value)
If bListen = True Then
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
'mBox.SetFocus
bListen = False
End If
End Sub

Private Sub fScroll_Scroll()
On Error Resume Next
mBox.Top = -(((mBox.Height - sBox.Height) / fScroll.Max) * fScroll.value)
If bListen = True Then
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
'mBox.SetFocus
bListen = False
End If
End Sub

Private Sub fScroll2_Change()
On Error Resume Next
mBox2.Top = -(((mBox2.Height - sBox2.Height) / fScroll2.Max) * fScroll2.value)

End Sub

Private Sub fScroll2_Scroll()
On Error Resume Next
mBox2.Top = -(((mBox2.Height - sBox2.Height) / fScroll2.Max) * fScroll2.value)
End Sub

Private Sub full_Click(Index As Integer)
On Error Resume Next
'DirectXEvent_DXCallback 0
End Sub

Private Sub grpTest_Click()
On Error Resume Next
Call StartEffect(fType(2).ListIndex, scrMagnitude(2).value, scrDuration(2).value, scrFrequency(2).value, dGrip)
End Sub

Private Sub jStick_Click()
On Error Resume Next
'joyIndex = jStick.ListIndex + 1
AcquireJoystick jStick.ListIndex + 1




End Sub

Private Sub jStick_GotFocus()
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If

bListen = False
End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
mBox.SetFocus
bListen = False
End Sub

Private Sub Label7_Click()
StartDoc "http://tocaedit.com"
End Sub

Private Sub Listen_Timer()
On Error Resume Next
If bListen Then
Dim kLoop As Long

For kLoop = 1 To 297
If KeyDown(kLoop) Then
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(kLoop)
ctlKeys(0, ctlIndex) = kLoop
txtPrimary(ctlIndex).SetFocus
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(kLoop)
ctlKeys(1, ctlIndex) = kLoop
txtSecondary(ctlIndex).SetFocus
End If
'mBox.SetFocus
bListen = False
End If
Next kLoop

End If
End Sub

Private Sub mBox_GotFocus()
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
mBox.SetFocus
bListen = False
End Sub

Private Sub mnu1_Click()
If ctlSecond = False Then
ctlKeys(0, ctlIndex) = 0
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
ctlKeys(1, ctlIndex) = 0
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
End Sub

Private Sub mnu2_Click()
frmManual.Show 1, frmMain
If Not manBut = -1 Then
If ctlSecond = False Then
ctlKeys(0, ctlIndex) = manBut
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
ctlKeys(1, ctlIndex) = manBut
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If

End If
End Sub

Private Sub optAim_Click(Index As Integer)
aMode = Index
SetAim aMode
End Sub

Private Sub optMode_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0:
frmForce.Visible = False
frmTest.Visible = False
frmVehicle.Visible = False
frmFoot.Visible = True
Case 1:
frmForce.Visible = False
frmTest.Visible = False
frmVehicle.Visible = True
frmFoot.Visible = False
Case 2:
frmForce.Visible = True
frmTest.Visible = False
frmVehicle.Visible = False
frmFoot.Visible = False
Case 3:
frmForce.Visible = False
frmTest.Visible = True
frmVehicle.Visible = False
frmFoot.Visible = False
End Select

End Sub

Private Sub optMode_GotFocus(Index As Integer)
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If

bListen = False
End Sub

Private Sub PollTimer_Timer()
On Error Resume Next

GetID "Grand theft auto San Andreas"
If hWin > 0 Then
    If didInit = False Then
        InitGame
        didInit = True
    End If
Else
    didInit = False
End If


JoyPoll
DirectXEvent_DXCallback 0

If didInit = False Then Exit Sub



Dim carpoint As Long, curWeapon As Long, curStatus As Long, curTime As Long, curTime2 As Long
Dim crashMag As Single, spinZ As Single, Speed As Single, sInfluence As Single
carpoint = GetLong(&HB6F3B8)
'If carpoint = 0 Then Exit Sub
If GetByte(&HB7CB49) = 1 Then Exit Sub






'Suspension stuff
Old1 = rr
Old2 = fr
Old3 = rl
Old4 = fl

If isBike(GetInteger(carpoint + &H22)) Then
fl = CLng(GetFloat(carpoint + &H720) * 100)
rl = CLng(GetFloat(carpoint + &H724) * 100)
fr = CLng(GetFloat(carpoint + &H728) * 100)
rr = CLng(GetFloat(carpoint + &H72C) * 100)
Else
fl = CLng(GetFloat(carpoint + &H7E4) * 100)
rl = CLng(GetFloat(carpoint + &H7E8) * 100)
fr = CLng(GetFloat(carpoint + &H7EC) * 100)
rr = CLng(GetFloat(carpoint + &H7F0) * 100)
End If

If Not ffinCar = inCar And inCar = True Then
ffWait = GetTickCount + 1000
End If
ffinCar = inCar
rr2 = Abs(Old1 - rr)
fr2 = Abs(Old2 - fr)
rl2 = Abs(Old3 - rl)
fl2 = Abs(Old4 - fl)
sAll = rr2 Or fr2 Or rl2 Or fl2
'Debug.Print sAll


If chkForce(1).value = Checked And inCar = True Then
    If sAll > 20 Then
        If GetTickCount > ffWait Then
            Call StartEffect(fType(1).ListIndex, ((sAll * 100) * 0.0001) * scrMagnitude(1).value, scrDuration(1).value, scrFrequency(1).value, dBump)
        End If
    End If
End If


'Collision
Static crashtest
If chkForce(0).value = Checked Then
crashMag = GetFloat(carpoint + &HD8)
crashtest = crashtest + crashMag + 30
crashMag = crashtest
'If crashMag > 0 Then
Call StartEffect(fType(0).ListIndex, (crashMag * 0.01) * scrMagnitude(0).value, scrDuration(0).value, scrFrequency(0).value, dCol)
crashtest = crashtest - 50
If crashtest < 0 Then crashtest = 0
'End If
End If

'Weapons
If chkForce(2).value = Checked Then
curWeapon = GetLong(&HB7CDBC)
curStatus = GetLong(((carpoint + &H5A0) + (curWeapon * &H1C)) + &H4)
If curStatus = 1 Then
Call StartEffect(fType(2).ListIndex, scrMagnitude(2).value, scrDuration(2).value, scrFrequency(2).value, dWep)
Call SetByte(((carpoint + &H5A0) + (curWeapon * &H1C)) + &H4, 0)
didLoad = False
End If

If curStatus = 2 Then
If didLoad = False Then
Call StartEffect(fType(2).ListIndex, scrMagnitude(2).value, scrDuration(2).value, scrFrequency(2).value, dWep)
didLoad = True
Else
'didLoad = False
End If
End If


End If

'If curStatus > 0 Then
'Debug.Print curStatus
'End If




'Explosions
If chkForce(3).value = Checked Then
curTime = GetLong(&HB6F084)
If oldTime = 0 Then oldTime = curTime
If Not oldTime = curTime Then
Call StartEffect(fType(3).ListIndex, GetFloat(&HB6F154) * 100000, scrDuration(3).value, scrFrequency(3).value, dExp)
oldTime = curTime
End If
End If

'Grip
If chkSteering.value = Checked And inCar = True Then
spinZ = -(GetFloat(carpoint + &H58) * 100)
Speed = Sqr((GetFloat(carpoint + &H44) ^ 2) + (GetFloat(carpoint + &H48) ^ 2) + (GetFloat(carpoint + &H4C) ^ 2)) * 50
sInfluence = ((LSus(fr) + LSus(fl)) * scrSuspension.value)

Debug.Print sInfluence; " "; fr
Call StartSteer(10000 - sInfluence, scrGrip.value, spinZ, ((scrSpring.value * Speed) - sInfluence), scrFriction - sInfluence, dGrip)
ElseIf chkSteering.value = Checked Then
StopSteer
End If




End Sub


Private Sub preset_Click()
On Error Resume Next
lSettings preset.List(preset.ListIndex)

End Sub

Private Sub preset_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
bSave_Click
End If
End Sub

Private Sub sBox_GotFocus()
On Error Resume Next
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
mBox.SetFocus
bListen = False
End Sub

Private Sub scrOverall_Change()
On Error Resume Next
SetRTCGAIN CBool(chkCenter.value), CLng(scrOverall.value)
End Sub

Private Sub susTest_Click()
On Error Resume Next
Call StartEffect(fType(1).ListIndex, scrMagnitude(1).value, scrDuration(1).value, scrFrequency(1).value, dBump)
End Sub

Private Sub txtPrimary_Click(Index As Integer)
On Error Resume Next
If bListen = True Then
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
bListen = False
End If

For jLoop = 0 To 307
orawInputs(jLoop) = rawInputs(jLoop)
Next jLoop

txtPrimary(Index).Text = "Press a key..."
ctlIndex = Index
ctlSecond = False
bListen = True
End Sub

Private Sub txtPrimary_LostFocus(Index As Integer)
'blisten = false
End Sub

Private Sub txtSecondary_Click(Index As Integer)
On Error Resume Next
If bListen = True Then
If ctlSecond = False Then
txtPrimary(ctlIndex).Text = GetKeyboardString(ctlKeys(0, ctlIndex))
Else
txtSecondary(ctlIndex).Text = GetKeyboardString(ctlKeys(1, ctlIndex))
End If
bListen = False
End If

For jLoop = 0 To 307
orawInputs(jLoop) = rawInputs(jLoop)
Next jLoop

txtSecondary(Index).Text = "Press a key..."
ctlIndex = Index
ctlSecond = True
bListen = True
End Sub


Public Function lText()
On Error Resume Next
For jLoop = 0 To txtPrimary.UBound
txtPrimary(jLoop) = GetKeyboardString(ctlKeys(0, jLoop))
txtSecondary(jLoop) = GetKeyboardString(ctlKeys(1, jLoop))
Next jLoop
End Function

Private Sub wepTest_Click()
On Error Resume Next
Call StartEffect(fType(2).ListIndex, scrMagnitude(2).value, scrDuration(2).value, scrFrequency(2).value, dWep)
End Sub

Public Function SetOld()
                           
 'set stuff
    Dim sLoop As Integer
    For sLoop = 0 To dDead.Count
        dDead_Change sLoop
        dSaturation_Change sLoop
    Next sLoop
            chkCenter_Click
        scrOverall_Change
End Function
