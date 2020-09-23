VERSION 5.00
Begin VB.Form frmDisplay 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9435
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      ToolTipText     =   "Click here to Exit"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   164
      Left            =   8640
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   163
      Left            =   8640
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   162
      Left            =   8640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   161
      Left            =   8640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   160
      Left            =   8640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   159
      Left            =   8640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   158
      Left            =   8640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   157
      Left            =   8640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   156
      Left            =   8640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   155
      Left            =   8640
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   154
      Left            =   8640
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   153
      Left            =   8040
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   152
      Left            =   8040
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   151
      Left            =   8040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   150
      Left            =   8040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   149
      Left            =   8040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   148
      Left            =   8040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   147
      Left            =   8040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   146
      Left            =   8040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   145
      Left            =   8040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   144
      Left            =   8040
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   143
      Left            =   8040
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   142
      Left            =   7440
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   141
      Left            =   7440
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   140
      Left            =   7440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   139
      Left            =   7440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   138
      Left            =   7440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   137
      Left            =   7440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   136
      Left            =   7440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   135
      Left            =   7440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   134
      Left            =   7440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   133
      Left            =   7440
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   132
      Left            =   7440
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   131
      Left            =   6840
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   130
      Left            =   6840
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   129
      Left            =   6840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   128
      Left            =   6840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   127
      Left            =   6840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   126
      Left            =   6840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   125
      Left            =   6840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   124
      Left            =   6840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   123
      Left            =   6840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   122
      Left            =   6840
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   121
      Left            =   6840
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   120
      Left            =   6240
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   119
      Left            =   6240
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   118
      Left            =   6240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   117
      Left            =   6240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   116
      Left            =   6240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   115
      Left            =   6240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   114
      Left            =   6240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   113
      Left            =   6240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   112
      Left            =   6240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   111
      Left            =   6240
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   110
      Left            =   6240
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   109
      Left            =   5640
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   108
      Left            =   5640
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   107
      Left            =   5640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   106
      Left            =   5640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   105
      Left            =   5640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   104
      Left            =   5640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   103
      Left            =   5640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   102
      Left            =   5640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   101
      Left            =   5640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   100
      Left            =   5640
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   99
      Left            =   5640
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   98
      Left            =   5040
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   97
      Left            =   5040
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   96
      Left            =   5040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   95
      Left            =   5040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   94
      Left            =   5040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   93
      Left            =   5040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   92
      Left            =   5040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   91
      Left            =   5040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   90
      Left            =   5040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   89
      Left            =   5040
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   88
      Left            =   5040
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   87
      Left            =   4440
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   86
      Left            =   4440
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   85
      Left            =   4440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   84
      Left            =   4440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   83
      Left            =   4440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   82
      Left            =   4440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   81
      Left            =   4440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   80
      Left            =   4440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   79
      Left            =   4440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   78
      Left            =   4440
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   77
      Left            =   4440
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   76
      Left            =   3840
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   75
      Left            =   3840
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   74
      Left            =   3840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   73
      Left            =   3840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   72
      Left            =   3840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   71
      Left            =   3840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   70
      Left            =   3840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   69
      Left            =   3840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   68
      Left            =   3840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   67
      Left            =   3840
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   66
      Left            =   3840
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   65
      Left            =   3240
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   64
      Left            =   3240
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   63
      Left            =   3240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   62
      Left            =   3240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   61
      Left            =   3240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   60
      Left            =   3240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   59
      Left            =   3240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   58
      Left            =   3240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   57
      Left            =   3240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   56
      Left            =   3240
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   55
      Left            =   3240
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   54
      Left            =   2640
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   53
      Left            =   2640
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   52
      Left            =   2640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   51
      Left            =   2640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   50
      Left            =   2640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   49
      Left            =   2640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   48
      Left            =   2640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   47
      Left            =   2640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   46
      Left            =   2640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   45
      Left            =   2640
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   44
      Left            =   2640
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   43
      Left            =   2040
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   42
      Left            =   2040
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   41
      Left            =   2040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   40
      Left            =   2040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   39
      Left            =   2040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   38
      Left            =   2040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   37
      Left            =   2040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   36
      Left            =   2040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   35
      Left            =   2040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   34
      Left            =   2040
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   33
      Left            =   2040
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   32
      Left            =   1440
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   31
      Left            =   1440
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   30
      Left            =   1440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   29
      Left            =   1440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   28
      Left            =   1440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   27
      Left            =   1440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   26
      Left            =   1440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   25
      Left            =   1440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   24
      Left            =   1440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   23
      Left            =   1440
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   22
      Left            =   1440
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   21
      Left            =   840
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   20
      Left            =   840
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   19
      Left            =   840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   18
      Left            =   840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   17
      Left            =   840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   16
      Left            =   840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   15
      Left            =   840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   14
      Left            =   840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   13
      Left            =   840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   12
      Left            =   840
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   11
      Left            =   840
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   10
      Left            =   240
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   9
      Left            =   240
      Top             =   5520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   8
      Left            =   240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   7
      Left            =   240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   6
      Left            =   240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   5
      Left            =   240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   4
      Left            =   240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   3
      Left            =   240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   2
      Left            =   240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   720
      Width           =   615
   End
   Begin VB.Image imgLayer1 
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'create general arrays for the 165 images ( 0 - 164 )
Dim texture(0 To 164) As Integer     'this array hold the textures number
Dim Walkable(0 To 164) As Integer     'this array hold the walk value
Dim MapName(0 To 164) As String


Private Sub cmdExit_Click()
'as the user of they would like to exit
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbYes Then
    End
End If

End Sub

Private Sub Form_Load()
'show the texture form
frmMain.Show

'initialize all the arrays
For x = 0 To 164
    texture(x) = 1000   '1000 = blank image
    Walkable(x) = 0     'set the walk values to 0 ( 0 = can't, 1 = can )
    MapName(x) = ""
    imgLayer1(x).Visible = True     'show layer 1
Next x

End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbYes Then
    End
End If
End Sub

Private Sub imgLayer1_Click(Index As Integer)

Dim Current As Integer

'set the image value to the currently selected image on the textures form
imgLayer1(Index).Picture = frmMain.picTextures(frmMain.txtCurrent.Text).Picture
'set the corrosponding texture variable to the texture number
texture(Index) = frmMain.txtCurrent.Text

If texture(Index) = 13 Then
    a = InputBox("Where does this door point to?", "Map Editor")
    MapName(Index) = a
    Walkable(Index) = 3
Else
    Walkable(Index) = frmMain.txtWalkable.Text
End If

End Sub


Public Function SaveMap()
'get the file name and path with an inputbox
'note :common dialog could have been used, but would be too time consuming
FileName = InputBox("Enter the path and file name of the file. The file extention should be *.MAP", "Game Editor")

'open the file to be written to
Open FileName For Output As #1

'write each variable
For x = 0 To 164
    Write #1, texture(x), Walkable(x), MapName(x) ', Walkable(X), Door(X)
Next x

'close the file
Close

End Function

Public Function LoadMap()

Dim ext As String

'get the file name and path with an inputbox
'note: common dialog could have been used, but would be too time consuming
FileName = InputBox("Enter the path of the file.", "Game Editor")

If Dir(FileName) <> "" And FileName <> "" Then
    'get the last three letters of the filename (extention)
    ext = UCase$(Right$(FileName, 3))
    'If ext = "MAP" Then     'if its a MAP file
        
        Open FileName For Input As #1

        For x = 0 To 164
            'read the data into the 4 variables
            Input #1, texture(x), Walkable(x)
            frmDisplay.imgLayer1(x).Picture = frmMain.picTextures(texture(x)).Picture
        Next x
        'close the file
        Close
    'Else        'if its not a MAP file
    '    'tell the user to only open MAP files
    '    MsgBox "Invalid file extention. Map editor can only read files saved in .MAP format", vbExclamation + vbOKOnly, "Map Editor"
    '    Exit Function
    'End If
        
Else

    MsgBox "Error: Path or file not found.", vbExclamation + vbOKOnly + vbDefaultButton1, "Map Editor."

End If

End Function

Public Function NewMap()
'clear all the images and variables
For x = 0 To 164
    texture(x) = 1000
    Walkable(x) = 0
    Door(x) = ""
    imgLayer1(x).Picture = frmMain.imgTextures(1000).Picture
    imgLayer1(x).Visible = True
Next x
End Function
