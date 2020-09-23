VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10365
   ControlBox      =   0   'False
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDisplay.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   8595
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFile 
      BackColor       =   &H00000000&
      Caption         =   " File Options "
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return To Game"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9570
      Left            =   -9360
      ScaleHeight     =   9510
      ScaleWidth      =   11355
      TabIndex        =   2
      Top             =   -8760
      Visible         =   0   'False
      Width           =   11415
   End
   Begin VB.PictureBox picBuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   480
      ScaleHeight     =   6495
      ScaleWidth      =   8850
      TabIndex        =   1
      Top             =   1440
      Width           =   8855
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   7440
   End
   Begin VB.PictureBox picSpr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6195
      Left            =   -4680
      Picture         =   "frmDisplay.frx":0A1C
      ScaleHeight     =   6135
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   -5760
      Visible         =   0   'False
      Width           =   6195
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================================================='
'========================================================='
'================= Dragon Lore Version ==================='
'============== Written by Matthew Eagar ================='
'============ Compiled in Visual Basic 6.0 ==============='
'========================================================='
'========================================================='
'
'   This program is an example of a VERY simple RPG game made
'   with VB 6.0.  I drew all the graphics in MS Paint,
'   and all coding is origional.
'
'   This isn't ment to be a full game, just an example.
'   there is no actual objective.  I havn't yet got doors
'   working, because that would require me to draw some more
'   textures for the insides of houses, which takes FOREVER!
'   Also, the textures could REALLY use some work.
'
'   This program may not run well on some computers.
'   The method used, bitblt, isn't very fast for animation.
'   It runs well on a P233, and a little slow on a P75 or less.
'
'   I'm still working on this, so look for me to post newer versions
'   of it.  It'll remain as free source code, and it's really ment for
'   educational purposes.
'
'   Please contact me with ANY questions, comments, suggestions, or problems,
'   ANY input is welcome:
'
'   email:  meagar@home.com
'   ICQ:    45058462
'
'   Also, I havn't tested this on any computer running anything less then VB6.
'   You will need the VB6 runtime files the use this.

Dim animX As Integer    'holds the current x location of the animation frame
Dim animY As Integer    'holds the current y location of the animation frame

Dim direction As Integer    'the direction the characters facing
Dim charX As Integer       'holds the character's x coords
Dim charY As Integer       'holds the character's y coords
Dim lastX As Integer    'holds the character's last y coords
Dim lastY As Integer    'holds the character's last x coords
Dim BackBuilt As Integer 'determines if the back ground needs to be built
Dim Speed As Integer    'holds the current speed, set by pressing the + or - keys
Dim mapx As Integer     'holds the current map x number
Dim mapy As Integer     'holds the current map y number
Dim MapName As String   'holds the name of the map

'symbolic constants
'directions
Const dLEFT As Integer = 1    'left direction
Const dUP As Integer = 2      'up direction
Const dRIGHT As Integer = 3   'right direction
Const dDOWN  As Integer = 4   'down direction

'animation frames
Const aLEFT As Integer = 2    'left animation
Const aUP As Integer = 104    'up animation
Const aRIGHT As Integer = 206 'right animation
Const aDOWN As Integer = 308  'down animation

'the Quit button
Private Sub cmdQuit_Click()
If MsgBox("Are you sure you would like to exit?", vbYesNo + vbDefaultButton2 + vbQuestion, "Dragon Lore") = vbYes Then End
End Sub

'the Return to game button
Private Sub cmdReturn_Click()
fraFile.Visible = False
End Sub


'when the user presses a key
Private Sub picBuf_KeyDown(KeyCode As Integer, Shift As Integer)

'copy the current location of the character into the lastx and lasty variables.
lastX = charX
lastY = charY

'determine how to act, based on which key the user presses.
Select Case KeyCode

Case Is = 37    'left arrow key
    animX = aLEFT   'set the animation frame to the proper direction
    direction = dLEFT 'set the direction

Case Is = 38    'up arrow key
    animX = aUP 'set the animation frame to the proper direction
    direction = dUP

Case Is = 39    'right arrow key
    animX = aRIGHT
    direction = dRIGHT

Case Is = 40    'down arrow key
    animX = aDOWN
    direction = dDOWN

Case Is = 27    'escape key

    fraFile.Visible = True
    tmrMove.Enabled = False
    
Case Is = 109   'minus key - decreases game speed
    Speed = Speed - 1
    If Speed <= 1 Then Speed = 1
    
Case Is = 107   'plus key - increases game speed
    Speed = Speed + 1
    If Speed >= 20 Then Speed = 20

End Select

'see if the movement timer should be enabled
If KeyCode >= 37 And KeyCode <= 40 Then
     
    If touching() <> 1 Then    'if the character isn't touching any nonwalkable tiles
        
        'move the character in the proper direction
        If direction = dLEFT Then
            charX = charX - Speed
        ElseIf direction = dUP Then
            charY = charY - Speed
        ElseIf direction = dRIGHT Then
            charX = charX + Speed
        ElseIf direction = dDOWN Then
            charY = charY + Speed
        End If
        
        'enable the movement timer, which animates the character
        tmrMove.Enabled = True
        
        'if movement, turn the mouse cursor into the invisible icon.
        'simply making a mouse cursor that was invisible is easier
        'then using API calls.
        frmDisplay.MouseIcon = frmTextures.picInvisible.Picture
    
    End If

End If

End Sub

Private Sub picBuf_KeyUp(KeyCode As Integer, Shift As Integer)

'disable the movement timer
tmrMove.Enabled = False

End Sub

Private Sub Form_Load()

'initialize the variables
animX = 2
animY = 1
charX = 200
charY = 200
BackBuilt = False
mapx = 0
mapy = 0
Speed = 6

picBuf.left = (Screen.Width - picBuf.Width) / 2
picBuf.Top = (Screen.Height - picBuf.Height) / 2
fraFile.left = (Screen.Width - fraFile.Width) / 2
fraFile.Top = (Screen.Height - fraFile.Width) / 2

Call BuildBack
Call redrawPic

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, animX As Single, animY As Single)
'turn the mouse icon into the visible icon
frmDisplay.MouseIcon = frmTextures.picVisible.Picture
End Sub



Private Sub tmrMove_Timer()

'see if the back ground has been built
If BackBuilt = False Then
    'build the background
    Call BuildBack
    BackBuilt = True
End If

Call redrawPic 'redraws the form

animY = animY + 51 'advance the frame

'there are 8 frames in the character's animation: this sees if the last frame has
'been shown. if it has, it resets it to the first.
If animY >= (51 * 8) Then
    animY = 1  'goes to first frame
End If

'see if the character has left the screen, by checking if the character's
'x or y position is greater then the total amount of tiles

If charX >= (40 * 13 + 50) Then 'if the character has left the right side of the screen
    mapx = mapx + 1 'set the current map name to the next map name
    charX = 5   'set the character's position back to the left side of the screen
    Call BuildBack  'redraw the back ground

ElseIf (charX + 40) <= 0 Then   'see if the character has left the left side of the screen
    mapx = mapx - 1 'set the current map name to the next map name
    charX = (40 * 13)   'set the character position to the right side of the screen
    Call BuildBack  'redraw the back ground

ElseIf charY <= 0 Then  'see if the character has left the top of the screen
    mapy = mapy + 1 'set the current map name to the next map name
    charY = (40 * 11)   'set the characters position to the bottom of the screen
    Call BuildBack  'redraw the back ground

ElseIf charY >= (40 * 11) Then  'see if the character has left the bottom of the screen
    mapy = mapy - 1 'set the current map name to the next map name
    charY = 5   'move the character to the top of the screen
    Call BuildBack  'redraw the back ground
End If

redrawPic   'refresh the picture

End Sub

'assembles the back ground
Sub BuildBack()

'this sub builds the back ground.  It is called only once per map,
'as the map is built in a hidden pic box, and kept untill the next map is needed.

Dim g As Integer    'counting variable
Dim a As Integer    'temp variable
Dim x As Integer    'holds x coords of tile
Dim y As Integer    'holds y coords of tile
On Error GoTo errHandler

x = 1
y = 1

'set the name of the map
MapName = App.Path & "\" & mapx & "a" & mapy & ".map"

'read the textures and the walkable values from the map file
Open MapName For Input As #1
    For g = 0 To 164
        Input #1, Texture(g), Walkable(g)
    Next g
Close

'clear the picture box which will hold the back ground
picBack.Cls

x = 0
y = 0

'loop through each tile, getting it with bitblt from frmTextures, and putting it into
'the picBack pic box.
For g = 0 To 164
    tileLeft(g) = y
    tileTop(g) = x
    a = BitBlt(picBack.hDC, y, x, 40, 40, frmTextures.picTextures(Texture(g)).hDC, 0, 0, SRCCOPY)
    x = x + 40
    
    'if a column has been finished, goto the next one
    If x >= (40 * 11) Then
        x = 0
        y = y + 40
    End If
Next g

'by-pass error handler
GoTo endSub

'for errors
errHandler:
    
MsgBox "Error number " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Dragon Lore"
MsgBox MapName & " was not found or was corrupted.  Please re-install this program."
End

endSub:
End Sub

Sub redrawPic()
Dim a As Integer

'this function draws the picture to the screen.

'Copy the back ground to the buffer pic box
a = BitBlt(picBuf.hDC, 0, 0, 590, 445, picBack.hDC, 0, 0, SRCCOPY)
'Copy the first layer of the sprite to the buffer
a = BitBlt(picBuf.hDC, charX, charY, 50, 50, picSpr.hDC, animX + 50, animY, SRCAND)
'Copy the second layer of the sprite to the buffer, for transparent effect.
a = BitBlt(picBuf.hDC, charX, charY, 50, 50, picSpr.hDC, animX, animY, SRCINVERT)
'refresh the picture
picBuf.Refresh

End Sub

Private Function touching() As Integer
Dim g As Integer ' counting variable

'this is the main part of the program
'it looks at the direction the character is moving, and sees if the next step
'will put the character onto a tile which has a walkable value of 1, which is
'either water trees or a building.  If it is, it returns 1. if not, it returns 0.

'loop for each of the tiles

If direction = dLEFT Then     'if the character is facing left
    For g = 0 To 164    'check each tile
        'check the left side of the character
        If (charX - Speed) >= tileLeft(g) And (charX - Speed) <= (tileLeft(g) + 40) And Walkable(g) = 1 Then
            
            'check the top left
            If charY >= tileTop(g) And charY <= (tileTop(g) + 40) Then
                
                touching = 1    'return that it is touching
                GoTo endFunct
            'check the bottom left
            ElseIf (charY + 50) >= tileTop(g) And (charY + 50) <= (tileTop(g) + 40) Then
                
                touching = 1    'return that it is touching
                GoTo endFunct
            End If
        
        End If
        
    Next g
        
ElseIf direction = dUP Then     'if the character is facing up
    
    For g = 0 To 164    'check each tile
        'check to top side of the character
        If (charY - Speed) >= tileTop(g) And (charY - Speed) <= (tileTop(g) + 40) And Walkable(g) = 1 Then
            
            'check the top left
            If charX >= tileLeft(g) And charX <= (tileLeft(g) + 40) Then
                
                touching = 1    'return that it is touching
                GoTo endFunct
            'check the top right
            ElseIf (charX + 50) >= tileLeft(g) And (charX + 50) <= (tileLeft(g) + 40) Then
                
                touching = 1    'return that it is touching
                GoTo endFunct
            End If
       
       End If
    
    Next g
        
ElseIf direction = dRIGHT Then     'if the character is facing right
    
    For g = 0 To 164    'check each tile
        'check the right side of the character
        If (charX + 50 + Speed) >= tileLeft(g) And (charX + 50 + Speed) <= (tileLeft(g) + 40) And Walkable(g) = 1 Then
                
            'check the top right
            If charY >= tileTop(g) And charY <= (tileTop(g) + 40) Then
                    
                touching = 1    'return that it is touching
                GoTo endFunct
            'check the botton right
            ElseIf (charY + 50) >= tileTop(g) And (charY + 50) <= (tileTop(g) + 40) Then
                    
                touching = 1    'return that it is touching
                GoTo endFunct
            End If
            
        End If
    Next g
    
ElseIf direction = dDOWN Then     'if the character is facing down

    For g = 0 To 164    'check each tile
        'check the bottom side of the character
        If (charY + 50 + Speed) >= tileTop(g) And (charY + 50 + Speed) <= (tileTop(g) + 40) And Walkable(g) = 1 Then
                    
            'check the bottom right
            If charX >= tileLeft(g) And charX <= (tileLeft(g) + 40) Then
                        
                touching = 1    'return that it is touching
                GoTo endFunct
            'check the bottom left
            ElseIf (charX + 50) >= tileLeft(g) And (charX + 50) <= (tileLeft(g) + 40) Then
                        
                touching = 1    'return that it is touching
                GoTo endFunct
            End If
            
       End If
    Next g
End If

endFunct:
End Function
