VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000010&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guns Help"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Label lblMessage 
      BackColor       =   &H80000010&
      Caption         =   "lblMessage"
      ForeColor       =   &H80000014&
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
   Begin VB.Menu mnuBtnContextMenu 
      Caption         =   "Invisible"
      Visible         =   0   'False
      Begin VB.Menu mnuBtnWhatsThis 
         Caption         =   "What's This?"
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ThisControl As Control


Private Sub Form_Load()
    lblMessage.Caption = "Guns" & vbCrLf & vbCrLf & _
    "The purpose of this game is to drop the desktop icons as quickly " & _
    "as possible." & vbCrLf & _
    "High scores exist per number of icons present on the desktop." & vbCrLf & _
    "You can choose to express the highscore either in time or in the " & _
    "number of shots needed. To do so just toggle the 'High score in shots fired' " & _
    "option in the menu you just accessed." & vbCrLf & vbCrLf & _
    "Start the game by clicking the 'Play' button; to end the " & _
    "application click 'Exit'." & vbCrLf & _
    "To get additional information on the game controls right-click " & _
    "the control in question and select 'What's This?'. " & _
    "You'll have to right-click text controls twice." & vbCrLf & _
    "The game is over when the fire button stays disabled." & vbCrLf & vbCrLf & _
    "Enjoy!"

End Sub

Private Sub mnuBtnWhatsThis_Click()
    
    Me.Height = 1300
    
    Select Case ThisControl.Name
    
        Case "cmdFire"
            lblMessage.Caption = _
            "This button fires a shot. The button is disabled as long " & _
            "as the fired bullet is still moving." & vbCrLf & _
            "The hot key for this control is ALT+I."
            
        Case "txtScore"
            lblMessage.Caption = _
            "This control keeps track of the game, either by displaying " & _
            "the elapsed time in seconds or the number of shots fired."
            
        Case "txtAngle"
            lblMessage.Caption = _
            "Enter in this textbox a value between 0 and 89 degrees." & vbCrLf & _
            "The hot key for this control is 'A'."
            
        Case "cmdIncreaseAngle"
        
            Me.Height = 1100

            lblMessage.Caption = _
            "This button increases the angle by one." & vbCrLf & _
            "The hot key for this control is 'O'."
        
        Case "cmdDecreaseAngle"
        
            Me.Height = 1100
            
            lblMessage.Caption = _
            "This button decreases the angle by one." & vbCrLf & _
            "The hot key for this control is 'K'."
            
        Case "txtVelocity"
            lblMessage.Caption = _
            "Enter in this textbox a velocity value between 0 and " & gintMaxVelocity & "." & vbCrLf & _
            "The hot key for this control is 'V'."
            
        Case "cmdIncreaseVelocity"
        
            Me.Height = 1100
        
            lblMessage.Caption = _
            "This button increases the velocity by five." & vbCrLf & _
            "The hot key for this control is 'P'."
        
        Case "cmdDecreaseVelocity"
        
            Me.Height = 1100
            
            lblMessage.Caption = _
            "This button decreases the velocity by five." & vbCrLf & _
            "The hot key for this control is 'L'."
            
        Case "cmdPlay"
        
            Me.Height = 1500
            
            lblMessage.Caption = _
            "Initiate a game by pressing this button." & vbCrLf & _
            "The game does not start until the icons are in postion." & vbCrLf & _
            "The hot key for this control is ALT+P."
            
        Case "cmdExit"
        
            Me.Height = 1500
        
            lblMessage.Caption = _
            "Selecting this button will end the application." & vbCrLf & _
            "The icons are placed back in their original positions." & vbCrLf & _
            "The hot key for this control is ALT+E."
        
        Case "picDesktop"
            lblMessage.Caption = _
            "This control shows the battlefield in which the icons " & _
            "are situated. I also need it for other stuff that is " & _
            "of no concern to you..."
    End Select
    
    Me.Show
End Sub
