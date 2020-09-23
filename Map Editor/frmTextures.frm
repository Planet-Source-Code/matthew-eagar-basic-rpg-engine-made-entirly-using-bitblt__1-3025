VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Map Editor"
   ClientHeight    =   3570
   ClientLeft      =   9570
   ClientTop       =   735
   ClientWidth     =   4830
   Icon            =   "frmTextures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4830
   Visible         =   0   'False
   Begin VB.Frame fraTextures 
      Caption         =   "Special Items"
      Height          =   2535
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtWalkable 
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Grass"
      Height          =   2415
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   61
         Left            =   120
         Picture         =   "frmTextures.frx":030A
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   62
         Left            =   120
         Picture         =   "frmTextures.frx":160C
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   63
         Left            =   120
         Picture         =   "frmTextures.frx":290E
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   64
         Left            =   840
         Picture         =   "frmTextures.frx":3C10
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   65
         Left            =   840
         Picture         =   "frmTextures.frx":4F12
         Top             =   960
         Width           =   600
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "House"
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   60
         Left            =   1560
         Picture         =   "frmTextures.frx":6214
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   59
         Left            =   1560
         Picture         =   "frmTextures.frx":7516
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   58
         Left            =   1560
         Picture         =   "frmTextures.frx":8818
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   57
         Left            =   2280
         Picture         =   "frmTextures.frx":9B1A
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   56
         Left            =   2280
         Picture         =   "frmTextures.frx":AE1C
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   55
         Left            =   2280
         Picture         =   "frmTextures.frx":C11E
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   54
         Left            =   840
         Picture         =   "frmTextures.frx":D420
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   53
         Left            =   840
         Picture         =   "frmTextures.frx":E722
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   52
         Left            =   840
         Picture         =   "frmTextures.frx":FA24
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   51
         Left            =   3000
         Picture         =   "frmTextures.frx":10D26
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   50
         Left            =   120
         Picture         =   "frmTextures.frx":12028
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   49
         Left            =   120
         Picture         =   "frmTextures.frx":1332A
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   48
         Left            =   120
         Picture         =   "frmTextures.frx":1462C
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Paths"
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   47
         Left            =   3000
         Picture         =   "frmTextures.frx":1592E
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   46
         Left            =   3000
         Picture         =   "frmTextures.frx":16C30
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   45
         Left            =   3000
         Picture         =   "frmTextures.frx":17F32
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   44
         Left            =   2280
         Picture         =   "frmTextures.frx":19234
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   43
         Left            =   2280
         Picture         =   "frmTextures.frx":1A536
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   42
         Left            =   2280
         Picture         =   "frmTextures.frx":1B838
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   41
         Left            =   2280
         Picture         =   "frmTextures.frx":1CB3A
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   40
         Left            =   1560
         Picture         =   "frmTextures.frx":1DE3C
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   39
         Left            =   1560
         Picture         =   "frmTextures.frx":1F13E
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   38
         Left            =   1560
         Picture         =   "frmTextures.frx":20440
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   37
         Left            =   1560
         Picture         =   "frmTextures.frx":21742
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   36
         Left            =   840
         Picture         =   "frmTextures.frx":22A44
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   35
         Left            =   840
         Picture         =   "frmTextures.frx":23D46
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   34
         Left            =   840
         Picture         =   "frmTextures.frx":25048
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   33
         Left            =   840
         Picture         =   "frmTextures.frx":2634A
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   32
         Left            =   120
         Picture         =   "frmTextures.frx":2764C
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   31
         Left            =   120
         Picture         =   "frmTextures.frx":2894E
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   30
         Left            =   120
         Picture         =   "frmTextures.frx":29C50
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   29
         Left            =   120
         Picture         =   "frmTextures.frx":2AF52
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Water"
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   16
         Left            =   120
         Picture         =   "frmTextures.frx":2C254
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   17
         Left            =   120
         Picture         =   "frmTextures.frx":2D556
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   18
         Left            =   120
         Picture         =   "frmTextures.frx":2E858
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   19
         Left            =   120
         Picture         =   "frmTextures.frx":2FB5A
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   20
         Left            =   840
         Picture         =   "frmTextures.frx":30E5C
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   21
         Left            =   840
         Picture         =   "frmTextures.frx":3215E
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   22
         Left            =   840
         Picture         =   "frmTextures.frx":33460
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   23
         Left            =   840
         Picture         =   "frmTextures.frx":34762
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   24
         Left            =   1560
         Picture         =   "frmTextures.frx":35A64
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   25
         Left            =   1560
         Picture         =   "frmTextures.frx":36D66
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   26
         Left            =   1560
         Picture         =   "frmTextures.frx":38068
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   27
         Left            =   1560
         Picture         =   "frmTextures.frx":3936A
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   28
         Left            =   2280
         Picture         =   "frmTextures.frx":3A66C
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Trees"
      Height          =   3135
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   15
         Left            =   2280
         Picture         =   "frmTextures.frx":3B96E
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   14
         Left            =   2280
         Picture         =   "frmTextures.frx":3CC70
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   13
         Left            =   2280
         Picture         =   "frmTextures.frx":3DF72
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   12
         Left            =   2280
         Picture         =   "frmTextures.frx":3F274
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   11
         Left            =   1560
         Picture         =   "frmTextures.frx":40576
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   10
         Left            =   1560
         Picture         =   "frmTextures.frx":41878
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   9
         Left            =   1560
         Picture         =   "frmTextures.frx":42B7A
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   8
         Left            =   1560
         Picture         =   "frmTextures.frx":43E7C
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   7
         Left            =   840
         Picture         =   "frmTextures.frx":4517E
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   6
         Left            =   840
         Picture         =   "frmTextures.frx":46480
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   5
         Left            =   840
         Picture         =   "frmTextures.frx":47782
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   4
         Left            =   840
         Picture         =   "frmTextures.frx":48A84
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   3
         Left            =   120
         Picture         =   "frmTextures.frx":49D86
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   2
         Left            =   120
         Picture         =   "frmTextures.frx":4B088
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   1
         Left            =   120
         Picture         =   "frmTextures.frx":4C38A
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgTextures 
         Height          =   600
         Index           =   0
         Left            =   120
         Picture         =   "frmTextures.frx":4D68C
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.TextBox txtCurrent 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgTextures 
      Height          =   600
      Index           =   1000
      Left            =   4080
      Top             =   2640
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
      Begin VB.Menu mnuPics 
         Caption         =   "&Special Items"
         Index           =   6
      End
   End
   Begin VB.Menu mnuFileView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewLayer1 
         Caption         =   "Layer&1 (Terrain)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLayer2 
         Caption         =   "Layer&2 (Items)"
      End
      Begin VB.Menu mnuViewMap 
         Caption         =   "&View Map"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Index"
      End
      Begin VB.Menu mnuHelpSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
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

'call the load registry sub
LoadReg

'set the current texture to the blank texture, 1000
txtCurrent.Text = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)

'ask the user if they would like to exit
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbNo Then
    Cancel = True 'if NO then cancel the form unload
Else
    SaveReg
    End
End If

End Sub

Private Sub imgTextures_Click(Index As Integer)
'set the current texture number to the textbox
txtCurrent.Text = Index

'set all the texture's borderstyle to 0
For x = 0 To 65
    imgTextures(x).BorderStyle = 0
Next x
'NOTE: the special reserved items array numbers are 900 - 999, this is so the
'computer can distinguish between which textures are items and which are textures
'also: the array number 1000 is a blank texture
'set the texture that was click on border style to 1
imgTextures(Index).BorderStyle = 1

End Sub

Private Sub mnuFileExit_Click()
'prompt the user if they would like to exit
If MsgBox("Are you sure you would like to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Map Editor") = vbYes Then
    'call teh save registry sub
    SaveReg
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


Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        'call the user32 dll to show the HELP file
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuHelpIndex_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        'call the user32 dll to show the HELP file index
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuPics_Click(Index As Integer)
'NOTE: a combo box could have been used, but it seemed easier to use an array of menus
'hide all the frames
For x = 1 To fraTextures.Count
    fraTextures(x).Visible = False
Next x

'determine what frame to show based on which menu item was clicked (mnuPics is an array)
fraTextures(Index).Visible = True

'set the walkable index based on which menu item was selected
txtWalkable.Text = Index

End Sub

Private Sub mnuViewLayer2_Click()
'check the layer 2 menu item, uncheck the layer 1 menu item
mnuViewLayer1.Checked = False
mnuViewLayer2.Checked = True

'show the layer 2 image array, hide the layer 1 image array
For x = 0 To 164
    frmDisplay.imgLayer1(x).Visible = False
    frmDisplay.imgLayer2(x).Visible = True
Next x
'set the current label caption on frmDisplay
frmDisplay.lblCurrent.Caption = "Viewing Layer 2"
End Sub

Private Sub mnuViewLayer1_Click()
'check the layer 1 menu item, uncheck the layer 2 menu item
mnuViewLayer1.Checked = True
mnuViewLayer2.Checked = False

'show the layer 1 image array, hide the layer 2 image array
For x = 0 To 164
    frmDisplay.imgLayer1(x).Visible = True
    frmDisplay.imgLayer2(x).Visible = False
Next x

'set the current label caption on frmDisplay
frmDisplay.lblCurrent.Caption = "Viewing Layer 1"

End Sub

Private Sub mnuViewMap_Click()
'shows both layers, so that the user can see what the final map will look like
For x = 0 To 164
    frmDisplay.imgLayer1(x).Visible = True
    frmDisplay.imgLayer2(x).Visible = True
Next x

'set the current label caption on frmDisplay
frmDisplay.lblCurrent.Caption = "Viewing Whole Map"
End Sub
