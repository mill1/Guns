VERSION 5.00
Begin VB.Form frmDesktop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()

    'I need this form to repaint the picture box
    Me.WindowState = vbNormal
    Me.Left = 0
    Me.Top = 0
    Me.Width = gintScreenWidht * SCREENTOFORM * gsngBattlefieldWidth
    Me.Height = gintScreenHeight * SCREENTOFORM
    Me.AutoRedraw = False
    Me.Show
    PaintDesktop Me.hdc
   
End Sub

