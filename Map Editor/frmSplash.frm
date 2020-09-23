VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":030A
   ScaleHeight     =   4755
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   4320
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ======================================================
'   ======================================================
'   =================== MAP EDITOR FOR ===================
'   ================ MATTHEW EAGAR'S GAME ================
'   ======================================================
'   ======================================================
'
'This program has been created by Matthew Eagar, Graphics also by Matthew Eagar.
'Some code for the actual game (not the Map Editor) was obtained at
'www.vb-world.net which is a web site devoted to helping VB programmers.
'no un-origional code was used in the creation of this editor.
'
'Date Started: March 09, 1999
'Date Finished: ??????
'
'NOTE: Because this program was written in VB6, some code may differ from that
'of VB3.  If you have any questions about code, contact the Aurthor, Matthew Eagar
'at meagar@home.com or eagarm@usa.net
'
'The purpose of this program is not for the user to create their own maps, but
'to help me create the game.  It is easier to make and then use this program
'then to write each map file out in notepad and save it as a .map file

Dim interval As Integer

Private Sub Form_Load()

'check to see if we are already running
If App.PrevInstance = True Then
    MsgBox "You may only run one copy of Map Editor at a time!", vbExclamation + vbOKOnly + vbDefaultButton1, "Map Editor"
    End
Else
    tmrLoad.Enabled = True
End If
End Sub



Private Sub tmrLoad_Timer()
 
    'load frmDisplay and frmMain
    Load frmDisplay
    Load frmMain
    frmDisplay.Visible = True
    frmMain.Visible = True
    
    Me.Hide
    
    tmrLoad.Enabled = False

End Sub
