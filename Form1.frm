VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "E.L. Diablo's 2D RT Image Rotation..."
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Demonstration"
      Height          =   330
      Left            =   375
      TabIndex        =   3
      Top             =   2460
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   1425
   End
   Begin VB.PictureBox picBck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   3900
      Left            =   2190
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   1
      Top             =   165
      Width           =   3900
   End
   Begin VB.PictureBox picCol 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   135
      Picture         =   "Form1.frx":30044
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   165
      Width           =   1980
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Danny@slave-studios.co.uk"
      ForeColor       =   &H8000000C&
      Height          =   240
      Left            =   2190
      TabIndex        =   5
      Top             =   4125
      Width           =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5340
      TabIndex        =   4
      Top             =   4035
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FRAME As Integer

Private Sub DanRotate(ByRef picDestHdc As Long, xPos As Long, yPos As Long, _
                       ByVal Angle As Long, _
                       ByRef picSrcHdc As Long, srcXoffset As Long, srcYoffset As Long, _
                       ByVal srcWidth As Long, ByVal srcHeight As Long)

'## DanRotate - Rotates an image.
'##
'## PicDestHdc      = the hDc of the target picturebox (ie. Picture2.hdc )
'## xPos            = the target coordinates (note that the image will be centered around these
'## yPos              coordinates).
'## Angle           = Rotate Angle (0-360)
'## PicSrcHdc       = The source image to rotate (ie. Picture1.hdc )
'## srcXoffset      = The offset coordinates within the Source Image to grab.
'## srcYoffset
'## srcWidth        = The width/height of the source image to grab.
'## srcHeight
'##
'## Returns: Nothing.

'## Please note this function doesn't check or returns anything. It's up to you to make sure all parameters
'## are valid, checked, etc.
'##
'## Use this code as you like. Credits appreciated.
'##
'## Danny van der Ark (danny@slave-studios.co.uk)
'## Aug 2Oo2



Dim Points(3) As POINTS2D
Dim DefPoints(3) As POINTS2D
Dim ThetS As Single, ThetC As Single
Dim ret As Long
    
    'SET LOCAL AXIS / ALIGNMENT
    Points(0).x = -srcWidth * 0.5
    Points(0).y = -srcHeight * 0.5
    
    Points(1).x = Points(0).x + srcWidth
    Points(1).y = Points(0).y
    
    Points(2).x = Points(0).x
    Points(2).y = Points(0).y + srcHeight
    
    'ROTATE AROUND Z-AXIS
    ThetS = Sin(Angle * NotPI)
    ThetC = Cos(Angle * NotPI)
    DefPoints(0).x = (Points(0).x * ThetC - Points(0).y * ThetS) + xPos
    DefPoints(0).y = (Points(0).x * ThetS + Points(0).y * ThetC) + yPos
    
    DefPoints(1).x = (Points(1).x * ThetC - Points(1).y * ThetS) + xPos
    DefPoints(1).y = (Points(1).x * ThetS + Points(1).y * ThetC) + yPos
    
    DefPoints(2).x = (Points(2).x * ThetC - Points(2).y * ThetS) + xPos
    DefPoints(2).y = (Points(2).x * ThetS + Points(2).y * ThetC) + yPos
    
    PlgBlt picDestHdc, DefPoints(0), picSrcHdc, srcXoffset, srcYoffset, srcWidth, srcHeight, 0, 0, 0

End Sub

Private Sub Command1_Click()

 Unload Me
 End
 
End Sub

Private Sub Command2_Click()
Dim tel As Integer

    'Reset & start FPS counter
    FRAME = 0
    Timer1.Enabled = True

    For tel = 0 To 360 Step 1
        picBck.Cls
        DanRotate picBck.hDC, 128, 128, tel, picCol.hDC, 0, 0, 128, 128
        picBck.Refresh
        FRAME = FRAME + 1
        DoEvents
    Next tel

    'Stop FPS timer
    Timer1.Enabled = False
    
End Sub

Private Sub Timer1_Timer()

 Label1.Caption = FRAME & " fps"
 FRAME = 0
 
End Sub
