VERSION 5.00
Begin VB.Form F_OSDCTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "F_OSDCTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oBB As New OffscreenDC
Private oBG As New OffscreenDC
Private oBall As New OffscreenDC
Private oMask As New OffscreenDC


Private Sub Form_Load()
Dim iX As Long, iY As Long
Dim iXD As Long, iYD As Long
Dim iSpd As Integer

  Me.Show
  Me.ScaleMode = vbPixels
  
  ' set the speed to move the ball
  iSpd = 2
  
  ' resize the offscreen DCs, this is always the first stepo in _
    using this class
 oBall.Resize 30, 30
 oMask.Resize 30, 30
 
 ' load the background image
 oBG.LoadStdPicture LoadPicture(IIf(Right(App.Path, 1) = "\", App.Path & "bg.jpg", App.Path & "\bg.jpg")), True
 
 ' set the form and backbuffer dc to the size of the background image
 Me.Height = Me.ScaleY(oBG.Height, vbPixels, vbTwips)
 Me.Width = Me.ScaleX(oBG.Width, vbPixels, vbTwips)
 oBB.Resize Me.ScaleHeight, Me.ScaleWidth
 
 ' set the background color of the Ball to white _
   for transparent blitting
 oBall.BackColor = vbWhite
 
 ' set the forecolor (text) and the fillcolor (shapes) of the balls DC
 oBall.ForeColor = vbBlack
 oBall.FillColor = vbRed
 ' clear the DC
 oBall.Clear
 ' draw the ball
 oBall.DrawCircle 15, 15, 15
 'draw some text
 oBall.DrawAlignedTxt "Ball", HCenter, VCenter, True
 ' create a mask from the ball
 oBall.MaskTo oMask.hdc, vbWhite
 
 iXD = iSpd
 iYD = iSpd
 
 ' randomize the starting position
 Randomize
 iX = Rnd * Me.ScaleHeight
 iY = Rnd * Me.ScaleWidth
 
 ' run animation loop until form is closed
 Do Until DoEvents = 0
 
    iX = iX + iXD
    iY = iY + iYD
    
    ' set the direction of travel on the x axis
    If iX + oBall.Width >= Me.ScaleWidth Then
        
        iXD = -iSpd
        
    ElseIf iX <= 0 Then
    
        iXD = iSpd
        
    End If
    
    ' set the direction of travel for the y axis
    If iY + oBall.Height >= Me.ScaleHeight Then
    
        iYD = -iSpd
        
    ElseIf iY <= 0 Then
    
        iYD = iSpd
        
    End If
    
    ' clear the backbuffer
    oBB.Clear
    
    ' blt the background image to the backbuffer
    oBG.BltTo oBB.hdc
    
    ' blt the ball forcing all white areas to be transparent
    oMask.BltTo oBB.hdc, iX, iY, , , , , vbSrcPaint
    oBall.BltTo oBB.hdc, iX, iY, , , , , vbSrcAnd
    
    'display the backbuffer
    oBB.BltTo Me.hdc
    
 Loop
End Sub

Private Sub Form_Terminate()
 ' free memory
 Set oBB = Nothing
 Set oBG = Nothing
 Set oBall = Nothing
 Set oMask = Nothing
End Sub

