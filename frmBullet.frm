VERSION 5.00
Begin VB.Form frmBullet 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   ClientHeight    =   180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   12
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton opt2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Value           =   -1  'True
      Width           =   270
   End
   Begin VB.OptionButton opt 
      Height          =   195
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "frmBullet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ChangeBkColor(pX As Single, pY As Single)
    On Error Resume Next
    'We don't want the bullet to look like a square
    
    opt.BackColor = GetPixel(frmGuns.picDesktop.hdc, pX, pY)
End Sub

Public Function Position(pintLeft As Integer, pintTop As Integer) As Integer
    
    Me.Left = pintLeft
    Me.Top = pintTop

    Position = Hit
    
End Function

Public Sub HideBullet()
    Me.Left = -200
    Me.Top = -200
End Sub

Private Function Hit() As Integer
    
    Dim i As Integer
    
    For i = 0 To gintDTIconsCount - 1
            
        If (Me.Left + Me.Width) / SCREENTOFORM > gudtPointArray(i).x And _
            Me.Left / SCREENTOFORM < gudtPointArray(i).x + gintIconWidth And _
            (Me.Top + Me.Height) / SCREENTOFORM > gudtPointArray(i).y And _
            Me.Top / SCREENTOFORM < gudtPointArray(i).y + gintIconHeight Then
            'You hit something!
            
            'But didn't you hit something at the bottom?
            If gudtPointArray(i).y <> CLng(gintScreenHeight * gsngBottom) Then
                gintNrOfHits = gintNrOfHits + 1
            End If
            
            opt.Value = True
                        
            If Not gblnSoundOff Then MessageBeep 0
            
            Hit = i
            Exit Function
        End If
        
    Next
    
    'Nothing's hit
    Hit = -1
    
End Function

