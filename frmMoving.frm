VERSION 5.00
Begin VB.Form frmMoving 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMask 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2040
      Picture         =   "frmMoving.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   3840
      Width           =   1020
   End
   Begin VB.PictureBox picSprite 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3120
      Picture         =   "frmMoving.frx":3042
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   3840
      Width           =   1020
   End
   Begin VB.Timer TimerMove 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1440
      Top             =   4200
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmMoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Moving Sprites
'
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim X As Long, Y As Long

Dim SpriteWidth As Long
Dim SpriteHeight As Long

Private Sub cmdStart_Click()

TimerMove.Enabled = True

End Sub

Private Sub Form_Load()

'Assign the width and height of the two picture boxes (they are identical)
SpriteWidth = picSprite.ScaleWidth
SpriteHeight = picSprite.ScaleHeight

End Sub

Private Sub TimerMove_Timer()
Static X As Long, Y As Long


X = X + 1
Y = Y + 1

'Keep the ball of the egde
If X > Me.ScaleWidth Then
    X = 0
End If

If Y > Me.ScaleHeight Then
    Y = 0
End If

'Clears the form
'uncomment
'Me.Cls

BitBlt Me.hDC, X, Y, SpriteWidth, SpriteHeight, picMask.hDC, 0, 0, vbSrcAnd
BitBlt Me.hDC, X, Y, SpriteWidth, SpriteHeight, picSprite.hDC, 0, 0, vbSrcPaint

'Force the form to update
'uncomment
'Me.Refresh

End Sub
