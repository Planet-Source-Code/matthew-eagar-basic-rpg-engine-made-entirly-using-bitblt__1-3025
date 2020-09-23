VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Map Editor"
   ClientHeight    =   5910
   ClientLeft      =   9570
   ClientTop       =   735
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7920
   Visible         =   0   'False
   Begin VB.Frame fraTextures 
      Caption         =   "Trees"
      Height          =   3255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   36
         Left            =   1560
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   24
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   37
         Left            =   840
         Picture         =   "frmMain.frx":160C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   23
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   38
         Left            =   1560
         Picture         =   "frmMain.frx":290E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   22
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   39
         Left            =   1560
         Picture         =   "frmMain.frx":3C10
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   21
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   40
         Left            =   840
         Picture         =   "frmMain.frx":4F12
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   20
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   41
         Left            =   840
         Picture         =   "frmMain.frx":6214
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   19
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   42
         Left            =   120
         Picture         =   "frmMain.frx":7516
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   18
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   43
         Left            =   120
         Picture         =   "frmMain.frx":8818
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   17
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   44
         Left            =   120
         Picture         =   "frmMain.frx":9B1A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   16
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   45
         Left            =   1560
         Picture         =   "frmMain.frx":AE1C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   46
         Left            =   840
         Picture         =   "frmMain.frx":C11E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   14
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   47
         Left            =   120
         Picture         =   "frmMain.frx":D420
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   13
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   57
         Left            =   5880
         Picture         =   "frmMain.frx":E722
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   12
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   61
         Left            =   2280
         Picture         =   "frmMain.frx":FA24
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   11
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   66
         Left            =   2280
         Picture         =   "frmMain.frx":10D26
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   67
         Left            =   2280
         Picture         =   "frmMain.frx":12028
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   9
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   68
         Left            =   2280
         Picture         =   "frmMain.frx":1332A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   8
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   70
         Left            =   4440
         Picture         =   "frmMain.frx":1462C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   660
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Water"
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   55
         Left            =   3000
         Picture         =   "frmMain.frx":1592E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   74
         Top             =   360
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   54
         Left            =   2280
         Picture         =   "frmMain.frx":16C30
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   73
         Top             =   1800
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":17F32
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   35
         Top             =   1800
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   2
         Left            =   2280
         Picture         =   "frmMain.frx":19234
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   34
         Top             =   360
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   8
         Left            =   840
         Picture         =   "frmMain.frx":1A536
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   33
         Top             =   1080
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   9
         Left            =   840
         Picture         =   "frmMain.frx":1B838
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   32
         Top             =   1800
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   11
         Left            =   840
         Picture         =   "frmMain.frx":1CB3A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   31
         Top             =   360
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   12
         Left            =   2280
         Picture         =   "frmMain.frx":1DE3C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   30
         Top             =   1080
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   15
         Left            =   1560
         Picture         =   "frmMain.frx":1F13E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   29
         Top             =   1080
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   16
         Left            =   1560
         Picture         =   "frmMain.frx":20440
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   28
         Top             =   1800
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   17
         Left            =   1560
         Picture         =   "frmMain.frx":21742
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   27
         Top             =   360
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   19
         Left            =   120
         Picture         =   "frmMain.frx":22A44
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   26
         Top             =   1080
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   22
         Left            =   120
         Picture         =   "frmMain.frx":23D46
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   25
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Paths"
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   1
         Left            =   2280
         Picture         =   "frmMain.frx":25048
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   54
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   3
         Left            =   2280
         Picture         =   "frmMain.frx":2634A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   53
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   4
         Left            =   2280
         Picture         =   "frmMain.frx":2764C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   52
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   5
         Left            =   2280
         Picture         =   "frmMain.frx":2894E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   51
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   6
         Left            =   3000
         Picture         =   "frmMain.frx":29C50
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   50
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   7
         Left            =   3000
         Picture         =   "frmMain.frx":2AF52
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   49
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   10
         Left            =   3000
         Picture         =   "frmMain.frx":2C254
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   48
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   24
         Left            =   120
         Picture         =   "frmMain.frx":2D556
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   47
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   25
         Left            =   840
         Picture         =   "frmMain.frx":2E858
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   46
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   26
         Left            =   1560
         Picture         =   "frmMain.frx":2FB5A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   45
         Top             =   2400
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   27
         Left            =   120
         Picture         =   "frmMain.frx":30E5C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   44
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   28
         Left            =   120
         Picture         =   "frmMain.frx":3215E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   43
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   29
         Left            =   120
         Picture         =   "frmMain.frx":33460
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   42
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   30
         Left            =   840
         Picture         =   "frmMain.frx":34762
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   41
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   31
         Left            =   840
         Picture         =   "frmMain.frx":35A64
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   40
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   32
         Left            =   1560
         Picture         =   "frmMain.frx":36D66
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   39
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   33
         Left            =   1560
         Picture         =   "frmMain.frx":38068
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   38
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   34
         Left            =   840
         Picture         =   "frmMain.frx":3936A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   37
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   35
         Left            =   1560
         Picture         =   "frmMain.frx":3A66C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   36
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "House"
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   13
         Left            =   1560
         Picture         =   "frmMain.frx":3B96E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   67
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   14
         Left            =   2280
         Picture         =   "frmMain.frx":3CC70
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   66
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   18
         Left            =   3000
         Picture         =   "frmMain.frx":3DF72
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   65
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   20
         Left            =   120
         Picture         =   "frmMain.frx":3F274
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   64
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   21
         Left            =   120
         Picture         =   "frmMain.frx":40576
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   63
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   23
         Left            =   120
         Picture         =   "frmMain.frx":41878
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   62
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   48
         Left            =   1560
         Picture         =   "frmMain.frx":42B7A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   61
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   49
         Left            =   840
         Picture         =   "frmMain.frx":43E7C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   60
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   50
         Left            =   840
         Picture         =   "frmMain.frx":4517E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   59
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   51
         Left            =   1560
         Picture         =   "frmMain.frx":46480
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   58
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   52
         Left            =   840
         Picture         =   "frmMain.frx":47782
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   57
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   53
         Left            =   2280
         Picture         =   "frmMain.frx":48A84
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   56
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   58
         Left            =   2280
         Picture         =   "frmMain.frx":49D86
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   55
         Top             =   1680
         Width           =   660
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Grass"
      Height          =   2415
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   60
         Left            =   120
         Picture         =   "frmMain.frx":4B088
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   72
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   63
         Left            =   840
         Picture         =   "frmMain.frx":4C38A
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   71
         Top             =   240
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   64
         Left            =   120
         Picture         =   "frmMain.frx":4D68C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   70
         Top             =   960
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   65
         Left            =   120
         Picture         =   "frmMain.frx":4E98E
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   69
         Top             =   1680
         Width           =   660
      End
      Begin VB.PictureBox picTextures 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   71
         Left            =   840
         Picture         =   "frmMain.frx":4FC90
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   68
         Top             =   960
         Width           =   660
      End
   End
   Begin VB.TextBox txtWalkable 
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCurrent 
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgTextures 
      Height          =   600
      Index           =   61
      Left            =   6000
      Picture         =   "frmMain.frx":50F92
      Top             =   840
      Width           =   600
   End
   Begin VB.Image imgTextures 
      Height          =   600
      Index           =   1000
      Left            =   7200
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPictures 
      Caption         =   "&Pictures"
      Begin VB.Menu mnuPics 
         Caption         =   "&Water"
         Index           =   1
      End
      Begin VB.Menu mnuPics 
         Caption         =   "&Paths"
         Index           =   2
      End
      Begin VB.Menu mnuPics 
         Caption         =   "&Houses"
         Index           =   3
      End
      Begin VB.Menu mnuPics 
         Caption         =   "&Trees"
         Index           =   4
      End
      Begin VB.Menu mnuPics 
         Caption         =   "&Grass"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declare the user32 DLL for use with the help file, at the entry point OSWinHelp%
'NOTE: although I do have an understanding of how to use DLLs, I did not write this
'function, because I didn't know which DLL handeled help files.
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub Form_Load()
'set the current texture to the blank texture, 1000
txtCurrent.Text = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)

'ask the user if they would like to exit
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbNo Then
    Cancel = True 'if NO then cancel the form unload
Else
    End
End If

End Sub

Private Sub picTextures_click(Index As Integer)
'set the current texture number to the textbox
txtCurrent.Text = Index

End Sub

Private Sub mnuFileExit_Click()
'prompt the user if they would like to exit
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbYes Then
    End
End If

End Sub

Private Sub mnuFileNew_Click()
temp = MsgBox("Save before starting new map?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Map Editor")

If temp = vbYes Then
    Call frmDisplay.SaveMap
ElseIf temp = vbNo Then
    Call frmDisplay.NewMap
ElseIf temp = vbCancel Then
    Exit Sub
End If

End Sub

Private Sub mnuFileOpen_Click()
'call the LoadMap sub in the frmDisplay
frmDisplay.LoadMap
End Sub

Private Sub mnuFileSave_Click()
'call the SaveMap sub in the frmDisplay
frmDisplay.SaveMap
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show

End Sub


Private Sub mnuPics_Click(Index As Integer)
'NOTE: a combo box could have been used, but it seemed easier to use an array of menus
'hide all the frames, and uncheck all the menu items
For x = 1 To mnuPics.Count
    mnuPics(x).Checked = False
    fraTextures(x).Visible = False
Next x
'determine the value of the walkable text box

If Index = 1 Then
    txtWalkable.Text = 1
ElseIf Index = 2 Then
    txtWalkable.Text = 0
ElseIf Index = 3 Then
    txtWalkable.Text = 1
ElseIf Index = 4 Then
    txtWalkable.Text = 1
ElseIf Index = 5 Then
    txtWalkable.Text = 0
End If


mnuPics(Index).Checked = True

'determine what frame to show based on which menu item was clicked (mnuPics is an array)
fraTextures(Index).Visible = True

End Sub
