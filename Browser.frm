VERSION 5.00
Begin VB.Form Browser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows System Colors Browser"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scroll Bars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   107
      Left            =   960
      TabIndex        =   112
      Top             =   6510
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000001A"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   106
      Left            =   4620
      TabIndex        =   111
      Top             =   6510
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbScrollBars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   105
      Left            =   2700
      TabIndex        =   110
      Top             =   6510
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   104
      Left            =   120
      TabIndex        =   109
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scroll Bars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   103
      Left            =   960
      TabIndex        =   108
      Top             =   6285
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000019"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   102
      Left            =   4620
      TabIndex        =   107
      Top             =   6285
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbScrollBars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   101
      Left            =   2700
      TabIndex        =   106
      Top             =   6285
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   100
      Left            =   120
      TabIndex        =   105
      Top             =   6285
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ToolTip"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   99
      Left            =   960
      TabIndex        =   104
      Top             =   6060
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000018"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   98
      Left            =   4620
      TabIndex        =   103
      Top             =   6060
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbInfoBackground"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   97
      Left            =   2700
      TabIndex        =   102
      Top             =   6060
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   96
      Left            =   120
      TabIndex        =   101
      Top             =   6060
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ToolTip Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   95
      Left            =   960
      TabIndex        =   100
      Top             =   5835
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000017"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   94
      Left            =   4620
      TabIndex        =   99
      Top             =   5835
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbInfoText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   93
      Left            =   2700
      TabIndex        =   98
      Top             =   5835
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000017&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   92
      Left            =   120
      TabIndex        =   97
      Top             =   5835
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Light Shadow"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   91
      Left            =   960
      TabIndex        =   96
      Top             =   5610
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000016"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   90
      Left            =   4620
      TabIndex        =   95
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vb3DLight"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   89
      Left            =   2700
      TabIndex        =   94
      Top             =   5610
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   88
      Left            =   120
      TabIndex        =   93
      Top             =   5610
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Dark Shadow"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   87
      Left            =   960
      TabIndex        =   92
      Top             =   5385
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000015"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   86
      Left            =   4620
      TabIndex        =   91
      Top             =   5385
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vb3DDKShadow"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   85
      Left            =   2700
      TabIndex        =   90
      Top             =   5385
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   84
      Left            =   120
      TabIndex        =   89
      Top             =   5385
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Highlight"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   83
      Left            =   960
      TabIndex        =   88
      Top             =   5160
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000014"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   82
      Left            =   4620
      TabIndex        =   87
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vb3DHighlight"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   81
      Left            =   2700
      TabIndex        =   86
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   80
      Left            =   120
      TabIndex        =   85
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inactive Title Bar Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   79
      Left            =   960
      TabIndex        =   84
      Top             =   4935
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000013"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   78
      Left            =   4620
      TabIndex        =   83
      Top             =   4935
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbInactiveTitleBarText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   77
      Left            =   2700
      TabIndex        =   82
      Top             =   4935
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   76
      Left            =   120
      TabIndex        =   81
      Top             =   4935
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   75
      Left            =   960
      TabIndex        =   80
      Top             =   4710
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000012"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   74
      Left            =   4620
      TabIndex        =   79
      Top             =   4710
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbButtonText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   73
      Left            =   2700
      TabIndex        =   78
      Top             =   4710
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   72
      Left            =   120
      TabIndex        =   77
      Top             =   4710
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disabled Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   71
      Left            =   960
      TabIndex        =   76
      Top             =   4485
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000011"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   70
      Left            =   4620
      TabIndex        =   75
      Top             =   4485
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbGrayText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   69
      Left            =   2700
      TabIndex        =   74
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   68
      Left            =   120
      TabIndex        =   73
      Top             =   4485
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Shadow"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   67
      Left            =   960
      TabIndex        =   72
      Top             =   4260
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000010"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   66
      Left            =   4620
      TabIndex        =   71
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbButtonShadow"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   65
      Left            =   2700
      TabIndex        =   70
      Top             =   4260
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   64
      Left            =   120
      TabIndex        =   69
      Top             =   4260
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button Face"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   63
      Left            =   960
      TabIndex        =   68
      Top             =   4035
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000F"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   62
      Left            =   4620
      TabIndex        =   67
      Top             =   4035
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbButtonFace"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   61
      Left            =   2700
      TabIndex        =   66
      Top             =   4035
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   60
      Left            =   120
      TabIndex        =   65
      Top             =   4035
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   59
      Left            =   960
      TabIndex        =   64
      Top             =   3810
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000E"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   58
      Left            =   4620
      TabIndex        =   63
      Top             =   3810
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbHighlightText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   57
      Left            =   2700
      TabIndex        =   62
      Top             =   3810
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   56
      Left            =   120
      TabIndex        =   61
      Top             =   3810
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   55
      Left            =   960
      TabIndex        =   60
      Top             =   3585
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000D"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   54
      Left            =   4620
      TabIndex        =   59
      Top             =   3585
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbHighlight"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   53
      Left            =   2700
      TabIndex        =   58
      Top             =   3585
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   52
      Left            =   120
      TabIndex        =   57
      Top             =   3585
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Application Workspace"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   51
      Left            =   960
      TabIndex        =   56
      Top             =   3360
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000C"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   50
      Left            =   4620
      TabIndex        =   55
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbApplicationWorkspace"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   49
      Left            =   2700
      TabIndex        =   54
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   48
      Left            =   120
      TabIndex        =   53
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inactive Border"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   47
      Left            =   960
      TabIndex        =   52
      Top             =   3135
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000B"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   46
      Left            =   4620
      TabIndex        =   51
      Top             =   3135
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbInactiveBorder"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   45
      Left            =   2700
      TabIndex        =   50
      Top             =   3135
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   44
      Left            =   120
      TabIndex        =   49
      Top             =   3135
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active Border"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   43
      Left            =   960
      TabIndex        =   48
      Top             =   2910
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H8000000A"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   42
      Left            =   4620
      TabIndex        =   47
      Top             =   2910
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbActiveBorder"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   41
      Left            =   2700
      TabIndex        =   46
      Top             =   2910
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   40
      Left            =   120
      TabIndex        =   45
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active Title Bar Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   39
      Left            =   960
      TabIndex        =   44
      Top             =   2685
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000009"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   38
      Left            =   4620
      TabIndex        =   43
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbActiveTitleBarText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   37
      Left            =   2700
      TabIndex        =   42
      Top             =   2685
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000009&
      Height          =   240
      Index           =   36
      Left            =   120
      TabIndex        =   41
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Window Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   35
      Left            =   960
      TabIndex        =   40
      Top             =   2460
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000008"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   34
      Left            =   4620
      TabIndex        =   39
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbWindowText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   33
      Left            =   2700
      TabIndex        =   38
      Top             =   2460
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   32
      Left            =   120
      TabIndex        =   37
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu Text"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   31
      Left            =   960
      TabIndex        =   36
      Top             =   2235
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000007"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   30
      Left            =   4620
      TabIndex        =   35
      Top             =   2235
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbMenuText"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   29
      Left            =   2700
      TabIndex        =   34
      Top             =   2235
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   28
      Left            =   120
      TabIndex        =   33
      Top             =   2235
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Window Frame"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   27
      Left            =   960
      TabIndex        =   32
      Top             =   2010
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000006"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   26
      Left            =   4620
      TabIndex        =   31
      Top             =   2010
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbWindowFrame"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   25
      Left            =   2700
      TabIndex        =   30
      Top             =   2010
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   24
      Left            =   120
      TabIndex        =   29
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Window Background"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   23
      Left            =   960
      TabIndex        =   28
      Top             =   1785
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000005"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   22
      Left            =   4620
      TabIndex        =   27
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbWindowBackground"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   21
      Left            =   2700
      TabIndex        =   26
      Top             =   1785
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   20
      Left            =   120
      TabIndex        =   25
      Top             =   1785
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu Bar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   19
      Left            =   960
      TabIndex        =   24
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000004"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   18
      Left            =   4620
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbMenuBar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   17
      Left            =   2700
      TabIndex        =   22
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   16
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inactive Title Bar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   15
      Left            =   960
      TabIndex        =   20
      Top             =   1335
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000003"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   14
      Left            =   4620
      TabIndex        =   19
      Top             =   1335
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbInactiveTitleBar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   13
      Left            =   2700
      TabIndex        =   18
      Top             =   1335
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   17
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active Title Bar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   11
      Left            =   960
      TabIndex        =   16
      Top             =   1110
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000002"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   10
      Left            =   4620
      TabIndex        =   15
      Top             =   1110
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbActiveTitleBar"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   2700
      TabIndex        =   14
      Top             =   1110
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desktop"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   960
      TabIndex        =   12
      Top             =   885
      Width           =   1755
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000001"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   4620
      TabIndex        =   11
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbDesktop"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   2700
      TabIndex        =   10
      Top             =   885
      Width           =   1935
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "IDE Calls It"
      Height          =   195
      Left            =   1020
      TabIndex        =   8
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scroll Bars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label Label5 
      Caption         =   "Hex Code"
      Height          =   195
      Left            =   4680
      TabIndex        =   6
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&&H80000000"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   4620
      TabIndex        =   5
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "VB Constant"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   420
      Width           =   1875
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vbScrollBars"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Sample"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Click on a cell to copy its value to the clipboard"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   855
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' Windows System Color Browser
' Created 18 May 2000 by James Vincent Carnicelli
'
' Notes:
' This trivial application is designed to help you use the colors
' that Windows allows users to customize instead of hard-coded
' RGB values that'll make your app. stick out like a sore thumb.
' If your memory is as crummy as mine, you won't remember the
' VB constants for these colors, let alone their hex values,
' when you're coding.  Here's your cheat sheet.
'
' Interestingly enough, the colors here will automatically be
' updated if you change your preferences without your having
' to restart the program.
'****************************************************************

Option Explicit

Private OriginalBackColor As Long

Private Sub Cell_Click(Index As Integer)
    If Cell(Index).Caption <> "" Then
        Clipboard.Clear
        If Left(Cell(Index).Caption, 2) = "&&" Then
            Clipboard.SetText Mid(Cell(Index).Caption, 2)
        Else
            Clipboard.SetText Cell(Index).Caption
        End If
    End If
End Sub

Private Sub Cell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Cell(Index).Caption <> "" Then
        OriginalBackColor = Cell(Index).BackColor
        Cell(Index).BackColor = vbHighlight
        Cell(Index).ForeColor = vbHighlightText
    End If
End Sub

Private Sub Cell_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Cell(Index).Caption <> "" Then
        Cell(Index).BackColor = OriginalBackColor
        Cell(Index).ForeColor = vbWindowText
    End If
End Sub
