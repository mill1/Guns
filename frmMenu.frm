VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGuns 
      Caption         =   "&Guns"
      Begin VB.Menu mnuHighScoreInShots 
         Caption         =   "&High score in shots fired"
      End
      Begin VB.Menu mnuSoundOff 
         Caption         =   "&Sound off"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Guns"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'To show a context menu you need an extra form with the menu to avoid a
'line around the main form (thanx Bill)

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelp_Click()
    If Not frmHelp Is Nothing Then
        Unload frmHelp
    End If
        
    frmHelp.Show
End Sub

Private Sub mnuHighScoreInShots_Click()

    mnuHighScoreInShots.Checked = Not mnuHighScoreInShots.Checked
    
    If mnuHighScoreInShots.Checked Then
        SaveSetting REG_APP, REG_SETTINGS, "HighScoreInShots", 1
        gblnShots = True
        frmGuns.lblScore.Caption = "Shots fired:"
    Else
        SaveSetting REG_APP, REG_SETTINGS, "HighScoreInShots", 0
        gblnShots = False
        frmGuns.lblScore.Caption = "Elapsed time:"
    End If
    
    GetHighScore
    frmGuns.lblHighScore.Caption = gstrHighScore
    frmGuns.txtScore.Text = 0
    
    GetSettings
End Sub


Private Sub mnuSoundOff_Click()
    
    mnuSoundOff.Checked = Not mnuSoundOff.Checked
    
    If mnuSoundOff.Checked Then
        SaveSetting REG_APP, REG_SETTINGS, "Sound Off", 1
    Else
        SaveSetting REG_APP, REG_SETTINGS, "Sound Off", 0
    End If
    
    GetSettings
End Sub
