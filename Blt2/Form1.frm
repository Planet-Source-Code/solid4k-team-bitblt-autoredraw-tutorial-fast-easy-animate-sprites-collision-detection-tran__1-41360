VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox uChr2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   4200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox uChr1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   4200
      Picture         =   "Form1.frx":4842
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox rChr2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   2160
      Picture         =   "Form1.frx":9084
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox rChr1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   120
      Picture         =   "Form1.frx":D8C6
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox lChr2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   2160
      Picture         =   "Form1.frx":12108
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox lChr1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   120
      Picture         =   "Form1.frx":1694A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox dChr2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   2160
      Picture         =   "Form1.frx":1B18C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox dChr1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   120
      Picture         =   "Form1.frx":1F9CE
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3795
      Left            =   120
      Picture         =   "Form1.frx":24210
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   120
      Width           =   6810
      Begin VB.Timer Timer3 
         Interval        =   200
         Left            =   1080
         Top             =   120
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ///////////////////////////////////////////////////////
' // BitBlt + AutoRedraw
' // And EASY & FAST way to make Games
' ///////////////////////////////////////////////////////
'
' BitBlt, as most know, stands for "Bit-Block Transfer."
' It is used for transferring a picture from one spot
' to another, and faster than doing the ol' GetPixel/SetPixel.
'
' This tutorial will show you how to use BitBlt to:
' 1. Transfer an Image to the Main Frame (Picture1)
' 2. Make a picture's background transparent
' 3. Animate a Sprite.
' 4. Keep track of the character Sprites' position.
' 5. Using the Position, make cheap, but working, collision detection.

' First off, let me explain some of the properties
' you will need before you can BitBlt a picture.
' Every picture you are going to BitBlt, you will
' need to set the "AutoRedraw" property ON.
' By the way, this eliminates Image Boxes, since they
' do not have a hDC nor AutoRedraw.

' Next, it is advised that you set the Picture Boxes' ScaleMode
' to "3 - Pixel".  This will help in keeping track of your Positions.
' Note:  You can also set your picture boxes' "Visible" property
'           to False, if you don't like it showing on your form.  You
'           can also just size your form to hide it.

' Now, after your properties are set, and you have all your pictures
' on your form,  you will then start your coding.  Let me explain
' the Declare Function BitBlt now:
' 1. BitBlt Lib "gdi32"
' >> The Declare is part of the gdi32 DLL (Dynamic Link Library).
' 2. hDestDC
' >> The Destination HDC of where your picture will be drawn to.
' >> In this program, it is Picture1.
' 3. X
' >> The X (horizontal) position to put the image on the Main Frame.
' >> This is PosX in this project.
' 4. Y
' >> The Y (vertical) position to put the image on the Main Frame.
' >> This is PosY in this project.
' 5. nWidth
' >> How wide is the picture?  In this project, the character is 32 pixels wide.
' 6. nHeight
' >> How tall is the picture?  In this project, the character is 48 pixels tall.
' 7. hSrcDC
' >> This is the picture that will be BitBlt'd to the Main Frame.  No matter
' >> how big it is, it will be cut to the size you like, specified in the nWidth
' >> and nHeight options.
' 8. xSrc
' >> On the picture, this option specifies where to cut from.  For a quick
' >> example, let's say you have a picture of : "||||||||||" (ten "|"s)
' >> and it is 10 pixels wide.  You would make xSrc 8, and it would make
' >> the picture start to be BitBlt'd for only 2 pixels, making the end result
' >> only "||".
' 9. ySrc
' >> The same as above, but the Y (vertical) plane, instead of X (horizontal).
' 10. dwRop
' >> Okay, this is the confusing one.  This specifies the drawing mode to use.
' >> The most common one is SRCPAINT, which is a const in our project.
' >> Oh, by the way, you MUST const the dwRops you are going to use, or
' >> it will not recognize them.

' Alright!  Done with that part--and on to the fun part: the Coding...
' Err...  Wait, we still need to define some variables...  Sorry!

' Right.  Uhh, first is the Declare GetKeyState.  Another API that is used to
' determine if a given key is down, or up.  I made a VERY useful sub
' for it : GetKey.  Quickly, here's how to use it:
' If GetKey(vbKey[key]) = True Then
' Example:
' If GetKey(vbKeyEscape) = True Then
'     End
' End If

' Then, the position keeper Variables; PosX; and PosY.  They are both Integers,
' since they hold pixel positions, and don't use decimal points.  These variables
' help us with collision detection.

' Next, we have the two dwRops that I use.  The most common; SRCPAINT; and
' another common one made for transparency; SRCAND.

' Finally, we have the commoners.  I will explain in haste:
' 1. Frames : It holds the Frames Per Second (FPS)
' 2. Col : The COLumn.  It holds what animation number we're on.
' 3. Tickz : This is the actual number width of what animation to show next.
' >> For example, the character is 32 pixels wide, and there's 4 pictures on each
' >> sheet.  That would make Tickz be either 0, 32, 64, or 96 (depending on
' >> the variable 'Col's value.  Phew.  That wasn't short at all!
' 4. CMove : Holds what position the character is facing.  Helps change
' >> his animation sheet, and in collision detection.
' 5. LastStand : Err, this holds what position he was last in when he
' >> was walking, so we'll know what position it is he'll be standing, of
' >> the player stops walking.

' Alright.  Now I'll just comment after each line of code, so you don't have to
' scroll back up and back down all the time.

' Note : As you might have noticed, the character I used is Sabin, from
' Final Fantasy 6 (3 in US).  I got his sprites from:
' http://www.rpgicons.com/frames.html
' But I edited him a bit, to my liking =]

' Need help?  Contact me at Solid4k1@cs.com
' Vote for me at PlanetSourceCode.com if you like it!

' Oh, a quick tip.  Transparency requires two pictures in seperate
' BMP files.  One with a white background, and your picture, and
' the other with a black background, and your picture is completely
' white.

' The best way to do this is make your background green, and then
' white out your main picture, and then fill your background black.

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim PosX As Integer
Dim PosY As Integer
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086
Dim Frames As Integer
Dim Col As Integer
Dim Tickz As Integer
Dim CMove As String
Dim LastStand As String
Public Function GetKey(lngKey As Long) As Boolean
If GetKeyState(lngKey&) < 0 Then
    GetKey = True
Else
    GetKey = False
End If
End Function

Private Sub Form_Load()
' The character starts facing down, and standing...
LastStand$ = "down"
' The character starts standing.
CMove$ = "stand"
' Just to set the right size, if it isn't in design time.
Me.Height = 4420
' The starting position is in the middle (sort of)
PosX = Picture1.ScaleWidth / 2 - 14
' Just the starting PosY
PosY = 40
' I had to change it a lot, so instead of loading the new picture all the time, I
' just added in the code.  Blah!  Heheh.
Picture1.Picture = LoadPicture(App.Path$ & "\bg.jpg")
End Sub

Private Sub Timer1_Timer()
' Timer1 is the refresh timer.  This timer is required for your computer
' to BitBlt the picture.
If GetKey(vbKeyLeft) Then
    ' If the left key is being pressed, move left and set variables for left.
    PosX = PosX - 1
    CMove$ = "left"
    LastStand$ = "left"
End If
If GetKey(vbKeyRight) Then
    ' If the right key is being pressed, move right and set variables for right.
    PosX = PosX + 1
    CMove$ = "right"
    LastStand$ = "right"
End If
If GetKey(vbKeyUp) Then
    ' The same as above, just up!
    PosY = PosY - 1
    CMove$ = "up"
    LastStand$ = "up"
End If
If GetKey(vbKeyDown) Then
    ' Blah!
    PosY = PosY + 1
    CMove$ = "down"
    LastStand$ = "down"
End If
If GetKey(vbKeyLeft) = False And GetKey(vbKeyRight) = False And GetKey(vbKeyUp) = False And GetKey(vbKeyDown) = False Then
    ' If none of the keys are being held down, set it in stand mode.
    CMove$ = "stand"
End If
' Make sure the character can't go offscreen.
If PosX < 5 Then
    PosX = 5
End If
If PosX > 415 Then
    PosX = 415
End If
If PosY < 5 Then
    PosY = 5
End If
If PosY > 195 Then
    PosY = 195
End If
' Collision detection on the Crates.
' If your position is inside the Crate's zone.
If PosX > 220 And PosX < 310 Then
    If PosY < 20 Then
        ' Checks which way you're moving, and pushes it back the opposite way.
        If CMove$ = "up" Then
            PosY = PosY + 1
        End If
        If CMove$ = "right" Then
            PosX = PosX - 1
        End If
        If CMove$ = "left" Then
            PosX = PosX + 1
        End If
    End If
End If
' Same as above.
If PosX > 35 And PosX < 130 Then
    If PosY > 55 And PosY < 110 Then
        If CMove$ = "right" Then
            PosX = PosX - 1
        End If
        If CMove$ = "left" Then
            PosX = PosX + 1
        End If
        If CMove$ = "up" Then
            PosY = PosY + 1
        End If
        If CMove$ = "down" Then
            PosY = PosY - 1
        End If
    End If
End If
' Same as above.
If PosX > 345 And PosX < 408 Then
    If PosY > 48 And PosY < 110 Then
        If CMove$ = "right" Then
            PosX = PosX - 1
        End If
        If CMove$ = "left" Then
            PosX = PosX + 1
        End If
        If CMove$ = "up" Then
            PosY = PosY + 1
        End If
        If CMove$ = "down" Then
            PosY = PosY - 1
        End If
    End If
End If
' Same as above.
If PosX > 290 And PosX < 352 Then
    If PosY > 145 And PosY < 185 Then
        If CMove$ = "right" Then
            PosX = PosX - 1
        End If
        If CMove$ = "left" Then
            PosX = PosX + 1
        End If
        If CMove$ = "up" Then
            PosY = PosY + 1
        End If
        If CMove$ = "down" Then
            PosY = PosY - 1
        End If
    End If
End If
' Clears the picture box.  Also is required.
Picture1.Cls
' Paints the pictures, if in stand mode!!!
' Note:  Transparancy requires two BitBlts.
If CMove$ = "stand" Then
    If LastStand$ = "down" Then
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, dChr2.hDC, 0, 0, SRCPAINT
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, dChr1.hDC, 0, 0, SRCAND
    End If
    If LastStand$ = "left" Then
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, lChr2.hDC, 0, 0, SRCPAINT
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, lChr1.hDC, 0, 0, SRCAND
    End If
    If LastStand$ = "right" Then
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, rChr2.hDC, 0, 0, SRCPAINT
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, rChr1.hDC, 0, 0, SRCAND
    End If
    If LastStand$ = "up" Then
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, uChr2.hDC, 0, 0, SRCPAINT
        BitBlt Picture1.hDC, PosX, PosY, 32, 48, uChr1.hDC, 0, 0, SRCAND
    End If
End If
' Paints pictures!!!
If CMove$ = "down" Then
    ' If the character is moving down, paint the down character, and the
    ' sprite that it's currently on.
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, dChr2.hDC, Tickz, 0, SRCPAINT
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, dChr1.hDC, Tickz, 0, SRCAND
End If
If CMove$ = "left" Then
    ' Same as above, but left.
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, lChr2.hDC, Tickz, 0, SRCPAINT
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, lChr1.hDC, Tickz, 0, SRCAND
End If
If CMove$ = "right" Then
    ' Same as above, but right.
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, rChr2.hDC, Tickz, 0, SRCPAINT
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, rChr1.hDC, Tickz, 0, SRCAND
End If
If CMove$ = "up" Then
    ' Same as above, but up.
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, uChr2.hDC, Tickz, 0, SRCPAINT
    BitBlt Picture1.hDC, PosX, PosY, 32, 48, uChr1.hDC, Tickz, 0, SRCAND
End If
' The Frames Per Second (FPS).
Frames = Frames + 1
End Sub

Private Sub Timer2_Timer()
' Print the Frame Rate and stuff.
Me.Caption = "Blt + AutoRedraw Tutorial - Frame Rate : " & Frames
' Reset the Frame Rate.
Frames = 0
End Sub

Private Sub Timer3_Timer()
' Change the collumn number.
Col = Col + 1
If Col = 4 Then
    ' 3 is the max, so if it's 4, reset it back to 0.
    Col = 0
End If
If Col = 0 Then
    ' The start point for Col 0 is pixel 0.
    Tickz = 0
End If
If Col = 1 Then
    ' The start point for Col 1 is pixel 32.
    Tickz = 32
End If
If Col = 2 Then
    ' The start point for Col 2 is pixel 64.
    Tickz = 64
End If
If Col = 3 Then
    ' The start point for Col 3 is pixel 96.
    Tickz = 96
End If
End Sub
