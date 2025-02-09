VERSION 5.00
Begin VB.Form frmGun 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   40
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   39
      X2              =   39
      Y1              =   28
      Y2              =   40
   End
   Begin VB.Line linGun2 
      BorderColor     =   &H80000005&
      X1              =   14
      X2              =   35
      Y1              =   14
      Y2              =   35
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   405
      Top             =   420
      Width           =   195
   End
   Begin VB.Line linGun 
      BorderWidth     =   4
      X1              =   35
      X2              =   14
      Y1              =   35
      Y2              =   14
   End
End
Attribute VB_Name = "frmGun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()

    Me.Left = frmGuns.Left
    Me.Top = frmGuns.Top - Me.Height
    
    linGun.X1 = 35
    linGun.X2 = 5
    linGun.Y1 = 35
    linGun.Y2 = 35
            
    linGun2.X1 = 35
    linGun2.X2 = 5
    linGun2.Y1 = 35
    linGun2.Y2 = 35
    
    Form_Paint
End Sub

Public Sub MoveGun(pintAngle As Integer)
'Between horizontal and vertical the sum of X2 and Y2 drops and rises again
'Max. sum value is 40 and the min. is 26 at 45 degrees. So we have to compensate.
'The min. value for X and Y is 5

Dim intCorr As Integer

    '40 - 26 = 14 namelijk
    intCorr = 14 - (Abs(45 - pintAngle) * (14 / 45))

    linGun.X2 = 5 + (pintAngle / 3) - (intCorr / 2)
    linGun2.X2 = linGun.X2
    linGun.Y2 = 40 - linGun.X2 - intCorr
    linGun2.Y2 = linGun.Y2

End Sub

Private Sub Form_Paint()

    Me.Show
    Me.AutoRedraw = False
    
    PaintDesktop Me.hdc
        
    frmGuns.SetFocus
End Sub


