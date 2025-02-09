VERSION 5.00
Begin VB.Form frmGuns 
   BorderStyle     =   0  'None
   ClientHeight    =   5895
   ClientLeft      =   -75
   ClientTop       =   -360
   ClientWidth     =   3090
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtScore 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   315
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2610
      Top             =   2430
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2610
      Top             =   1965
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "F&ire"
      Height          =   375
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2610
      Top             =   1500
   End
   Begin VB.TextBox txtVelocity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "150"
      ToolTipText     =   "Insert a velocity value"
      Top             =   885
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton cmdIncreaseVelocity 
      Caption         =   "+"
      Height          =   225
      Left            =   1950
      TabIndex        =   5
      Top             =   885
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdDecreaseVelocity 
      Caption         =   "-"
      Height          =   225
      Left            =   1950
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdDecreaseAngle 
      Caption         =   "-"
      Height          =   225
      Left            =   750
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdIncreaseAngle 
      Caption         =   "+"
      Height          =   225
      Left            =   750
      TabIndex        =   2
      Top             =   885
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtAngle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   225
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Insert an angle value"
      Top             =   885
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2610
      Top             =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1290
      TabIndex        =   8
      Top             =   1905
      Width           =   1020
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   210
      TabIndex        =   7
      Top             =   1905
      Width           =   1020
   End
   Begin VB.PictureBox picDesktop 
      Height          =   2880
      Left            =   150
      ScaleHeight     =   188
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   9
      Top             =   2925
      Width           =   2820
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      Caption         =   "Elapsed time:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1230
      TabIndex        =   14
      Top             =   345
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblVelocity 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      Caption         =   "&Velocity"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1305
      TabIndex        =   12
      Top             =   1500
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      Caption         =   "&Angle"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   1500
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblHighScore 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      Caption         =   "Highscore: 8888 seconds"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   420
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmGuns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private mintHitIconIndex As Integer
Private mblnExit As Boolean
Private mintPaintCount As Integer
Private mblnRepaint As Boolean


Public Sub Init()

    Dim strDisplayMessage As String
    
    
    SetControls
    mintHitIconIndex = -1
    Form_Paint
    
    mblnRepaint = True
    
    strDisplayMessage = GetSetting(REG_APP, REG_SYSTEM, "Display Message", "?")
    
    If strDisplayMessage = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Display Message", 1
        strDisplayMessage = "1"
    End If

    If CBool(strDisplayMessage) Then frmActiveDeskTop.Show
    
End Sub

Public Sub SetControls()

'Take the resolution into account

    picDesktop.Width = gintScreenWidht / 5.2
    picDesktop.Height = gintScreenHeight / 4.1
    
    picDesktop.ScaleMode = vbTwips

    Me.Width = picDesktop.ScaleWidth + 200
    Me.Height = picDesktop.ScaleHeight * 2.5
    Me.Left = gintScreenWidht * SCREENTOFORM * gsngBattlefieldWidth + 15
    Me.Top = (gintScreenHeight * SCREENTOFORM) * gsngGunPosition

    picDesktop.ScaleMode = vbPixels

    picDesktop.Left = Me.ScaleHeight / 100
    picDesktop.Top = Me.ScaleHeight / 2
    
    'The rest of the controls:
    cmdFire.Top = 10
    cmdFire.Left = picDesktop.Left
    
    txtScore.Top = cmdFire.Top
    txtScore.Left = (picDesktop.Left + picDesktop.Width) - txtScore.Width
    
    lblScore.Top = cmdFire.Top + 2
    lblScore.Left = txtScore.Left - (lblScore.Width + 3)
    
    txtAngle.Top = (picDesktop.Top - (cmdFire.Top + cmdFire.Height)) * 0.1
    txtAngle.Top = txtAngle.Top + cmdFire.Top + cmdFire.Height
    txtAngle.Left = picDesktop.Left
    
    cmdIncreaseAngle.Top = txtAngle.Top
    cmdIncreaseAngle.Left = txtAngle.Left + txtAngle.Width + 2
    
    cmdDecreaseAngle.Top = cmdIncreaseAngle.Top + cmdIncreaseAngle.Height + 2
    cmdDecreaseAngle.Left = cmdIncreaseAngle.Left
    
    lblAngle.Top = txtAngle.Top + txtAngle.Height + 3
    lblAngle.Left = picDesktop.Left
    
    cmdPlay.Top = (picDesktop.Top - (cmdFire.Top + cmdFire.Height)) * 0.6
    cmdPlay.Top = cmdPlay.Top + cmdFire.Top + cmdFire.Height
    cmdPlay.Left = picDesktop.Left
    
    cmdExit.Top = cmdPlay.Top
    cmdExit.Left = cmdPlay.Left + cmdPlay.Width + 3
    
    txtVelocity.Top = txtAngle.Top
    txtVelocity.Left = cmdExit.Left
    
    cmdIncreaseVelocity.Top = txtVelocity.Top
    cmdIncreaseVelocity.Left = txtVelocity.Left + txtVelocity.Width + 2
    
    cmdDecreaseVelocity.Top = cmdIncreaseVelocity.Top + cmdIncreaseVelocity.Height + 2
    cmdDecreaseVelocity.Left = cmdIncreaseVelocity.Left
    
    lblVelocity.Top = txtVelocity.Top + txtVelocity.Height + 3
    lblVelocity.Left = txtVelocity.Left
    
    lblHighScore.Top = picDesktop.Top - lblHighScore.Height - 5
    lblHighScore.Left = picDesktop.Left + (picDesktop.Width / 2) - (lblHighScore.Width / 2)
    
    'Initially hide the bullet
    frmBullet.HideBullet
    frmBullet.Show
    
End Sub


Private Sub cmdFire_Click()
        
    If InvalidAngle Then Exit Sub
    If InvalidVelocity Then Exit Sub
    
    cmdFire.Enabled = False

    GetTrajectory CInt(txtAngle.Text), CInt(txtVelocity.Text)
    
    If gblnShots Then txtScore.Text = CInt(txtScore.Text) + 1
    
    Timer2.Interval = gintTimer2Interval
    Timer2.Enabled = True
    
End Sub
Private Sub cmdExit_Click()
    
On Error GoTo HandleError
    
    mblnExit = True
    
    If UBound(gudtPointArrayOld) < 0 Then
        'raise an error to check if user has played at all
    Else
        gudtPointArrayOld = gudtPointArray
        gudtPointArray = gudtOrigPointArray
        Timer1.Enabled = True
    End If
    
    Exit Sub
    
HandleError:
    Bye
End Sub

Private Sub cmdPlay_Click()

    Static blnReplay As Boolean
    
    cmdPlay.Enabled = False
    
    If blnReplay Then
        gudtPointArrayOld = gudtPointArray
    Else
        gudtPointArrayOld = gudtOrigPointArray
        blnReplay = True
        cmdPlay.Caption = "&Play Again"
    End If
    
    GetRandomizedArray
    
    cmdPlay.Enabled = True
    
    Timer1.Enabled = True
    
End Sub

Private Sub cmdIncreaseAngle_Click()
    
    'Because of hot key
    If Not IsNumeric(txtAngle.Text) Then
        InvalidAngle
        Exit Sub
    End If
            
    If CInt(txtAngle.Text) <= 88 Then
        txtAngle.Text = CInt(txtAngle.Text) + 1
        frmGun.MoveGun CInt(txtAngle.Text)
    End If
    
End Sub

Private Sub cmdDecreaseAngle_Click()

    'Because of hot key
    If Not IsNumeric(txtAngle.Text) Then
        InvalidAngle
        Exit Sub
    End If

    If CInt(txtAngle.Text) >= 1 Then
        txtAngle.Text = CInt(txtAngle.Text) - 1
        frmGun.MoveGun CInt(txtAngle.Text)
    End If

End Sub

Private Sub cmdIncreaseVelocity_Click()

    'Because of hot key
    If Not IsNumeric(txtVelocity.Text) Then
        InvalidVelocity
        Exit Sub
    End If

    If CInt(txtVelocity.Text) <= gintMaxVelocity - 5 Then
        txtVelocity.Text = CInt(txtVelocity.Text) + 5
    End If

End Sub

Private Sub cmdDecreaseVelocity_Click()

    'Because of hot key
    If Not IsNumeric(txtVelocity.Text) Then
        InvalidVelocity
        Exit Sub
    End If
    
    If CInt(txtVelocity.Text) >= 5 Then
        txtVelocity.Text = CInt(txtVelocity.Text) - 5
    End If
    
End Sub

Private Sub Form_Paint()
    
'I wouldn't touch this bit:

    mintPaintCount = mintPaintCount + 1
    
    If Not mblnRepaint Then
        If mintPaintCount > 2 Then
        
            Unload frmDesktop
            
            MsgBox "This game encountered a problem." & vbCrLf & _
                "Please increase your color resolution to more than" & vbCrLf & _
                "256 colors and try again." & vbCrLf & _
                "(right-click the desktop ; Properties ; tab Settings)", vbInformation

            End 'Stop program
        End If
    End If

    frmDesktop.Init
    
    Me.AutoRedraw = True
    
    'Before showing AutoRedraw has to be set to True
    Me.Show
    
    Me.AutoRedraw = False
    

    PaintDesktop Me.hdc
        
    picDesktop.AutoRedraw = False
    
    StretchBlt picDesktop.hdc, 0, 0, gintScreenWidht / 4, gintScreenHeight / 4, _
                frmDesktop.hdc, 0, 0, _
                gintScreenWidht, _
                gintScreenHeight, SCRCOPY
    DoEvents
    
    LockWindowUpdate Me.hwnd
    
    'This is the only way to (keep) show(ing) characters on the form
    If gblnShots Then
        lblScore.Caption = "Shots fired:"
    Else
        lblScore.Caption = "Elapsed time:"
    End If
    
    lblAngle.Caption = "&Angle"
    lblVelocity.Caption = "&Velocity"
    lblHighScore.Caption = gstrHighScore
    
    'Unlock the windowupdate...
    LockWindowUpdate False
    
    Unload frmDesktop
    
End Sub

Private Sub lblVelocity_DblClick()
    frmEegg.Show
End Sub

Private Sub Timer1_Timer()

'Distribute the icons over the battlefield

    Static i As Integer
    Static intStep As Integer
    Static blnInitGun As Boolean
    Static blnGunMoved As Boolean
    
    intStep = intStep + 1
    
    If intStep < gintIconSteps + 1 Then
             
        For i = 0 To gintDTIconsCount - 1
            MoveIcon i, NewLeft(intStep, i), NewTop(intStep, i)
        Next
    
    Else
        If mblnExit Then
            Bye
        Else
            'move the gun
            If Not (blnInitGun) Then
                frmGun.Init
                i = 1
                blnInitGun = True
            End If
            
            If blnGunMoved Then
                'All initializations take place here
                'also when pressing (Re)play
                intStep = 0
                Timer1.Enabled = False
                cmdFire.Enabled = True
                cmdFire.SetFocus
                frmMenu.mnuHighScoreInShots.Enabled = False
                gintNrOfHits = 0
                txtScore.Text = 0
                
                If Not gblnShots Then Timer4.Enabled = False
                If Not gblnShots Then Timer4.Enabled = True
            Else
                frmGun.MoveGun i
            
                i = i + 2
                
                If i > 45 Then
                    blnGunMoved = True
                    
                    lblScore.Visible = True
                    
                    If gblnShots Then
                        lblScore.Caption = "Shots fired:"
                    Else
                        lblScore.Caption = "Elapsed time:"
                    End If
                    
                    txtScore.Visible = True
                    txtAngle.Text = 45
                    txtAngle.Visible = True
                    lblAngle.Visible = True
                    lblAngle.Caption = "Angle"
                    lblVelocity.Visible = True
                    lblVelocity.Caption = "Velocity"
                    lblHighScore.Visible = True
                    lblHighScore.Caption = "HighScore"
                    txtVelocity.Visible = True
                    cmdIncreaseAngle.Visible = True
                    cmdDecreaseAngle.Visible = True
                    cmdIncreaseVelocity.Visible = True
                    cmdDecreaseVelocity.Visible = True
                    cmdFire.Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Sub Timer2_Timer()

'Move the bullet over the desktop

'   'i = time
    Static i As Integer

    picDesktop_MouseDown 0, 1, CSng(gudtTrajectoryArray(i).x), CSng(gudtTrajectoryArray(i).y)

    If mintHitIconIndex <> -1 Then
        'Something was hit
        GoTo EndShot
    Else
        i = i + 1
    
        If i = UBound(gudtTrajectoryArray) Then
            'Nothing was hit
            frmBullet.HideBullet
            cmdFire.Enabled = True
            cmdFire.SetFocus
            GoTo EndShot
        End If
    End If
        
    Exit Sub
    
EndShot:
    Timer2.Enabled = False
    i = 0
    
End Sub

Private Sub Timer3_Timer()
    
'Drop the hit icon
    
    Static i As Integer
    Static blnBounced As Boolean
    Static intLowestY As Integer
    Static blnStep As Boolean
    Static intStep As Integer
    
    If i = 1 Then
        frmBullet.HideBullet
        frmBullet.opt.Value = False
    End If
    
    If Not blnBounced Then
        gudtPointArray(mintHitIconIndex).y = _
            gudtPointArray(mintHitIconIndex).y - CLng(GetY(0, i, 45, 1))
        
        MoveIcon mintHitIconIndex, _
                gudtPointArray(mintHitIconIndex).x, gudtPointArray(mintHitIconIndex).y
            
        i = i + 1
     
        If gudtPointArray(mintHitIconIndex).y > gintScreenHeight * gsngBottom Then
            blnBounced = True
            intLowestY = gudtPointArray(mintHitIconIndex).y
            i = 0
        End If
    Else
                    
        i = i + 1
        
        If mintHitIconIndex = -1 Then
            Exit Sub
        Else
            If Not blnStep Then
                intStep = IIf(CInt(intLowestY - gintScreenHeight * gsngBottom) >= 5, _
                            CInt(intLowestY - gintScreenHeight * gsngBottom) / 5, 2)
                
                blnStep = True
            End If
            
            MoveIcon mintHitIconIndex, gudtPointArray(mintHitIconIndex).x, _
                intLowestY - (intLowestY - (gintScreenHeight * gsngBottom)) * _
                i / intStep
        End If
            
        If i = intStep Then
            
            i = 0
            blnBounced = False
            blnStep = False
            gudtPointArray(mintHitIconIndex).y = gintScreenHeight * gsngBottom
            Timer3.Enabled = False
            
            If gintNrOfHits <> gintDTIconsCount Then
            
                cmdFire.Enabled = True
                cmdFire.SetFocus
                            
            Else
            'Game over
                frmMenu.mnuHighScoreInShots.Enabled = True

                If gblnShots Then
                    
                    If CInt(txtScore.Text) < CInt(GetSetting(REG_APP, REG_SHOTS, gintDTIconsCount)) Then
                        MsgBox "Highscore!", vbExclamation
                        
                        SaveSetting REG_APP, REG_SHOTS, gintDTIconsCount, txtScore.Text
                        gstrHighScore = "Highscore: " & txtScore.Text & " shots"
                        lblHighScore.Caption = gstrHighScore
                        
                    End If
                Else
                    Timer4.Enabled = False
                    
                    If CInt(txtScore.Text) < CInt(GetSetting(REG_APP, REG_TIME, gintDTIconsCount)) Then
                        MsgBox "Highscore!", vbExclamation
                        
                        SaveSetting REG_APP, REG_TIME, gintDTIconsCount, txtScore.Text
                        gstrHighScore = "Highscore: " & txtScore.Text & " seconds"
                        lblHighScore.Caption = gstrHighScore
                        
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Handle the hot keys
    Select Case LCase(Chr(KeyAscii))
        Case "o"
            If cmdIncreaseAngle.Enabled Then
                cmdIncreaseAngle.SetFocus
                cmdIncreaseAngle_Click
            End If
        Case "k"
            If cmdDecreaseAngle.Enabled Then
                cmdDecreaseAngle.SetFocus
                cmdDecreaseAngle_Click
            End If
        Case "p"
            If cmdIncreaseVelocity.Enabled Then
                cmdIncreaseVelocity.SetFocus
                cmdIncreaseVelocity_Click
            End If
        Case "l"
            If cmdDecreaseVelocity.Enabled Then
                cmdDecreaseVelocity.SetFocus
                cmdDecreaseVelocity_Click
            End If
        Case "a"
            txtAngle.SetFocus
        Case "v"
            txtVelocity.SetFocus
    End Select
    
End Sub

Private Sub Timer4_Timer()
'duh
    txtScore.Text = CInt(txtScore.Text) + 1
End Sub

Private Sub DropIcon()
    'Drop Icon
    '0 is "My Computer"
    '1 is "My Documents"
    
    Timer3.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then   ' Check if right mouse button was clicked.
      PopupMenu frmMenu.mnuGuns  ' Display the File menu as a pop-up menu.
   End If
End Sub

Private Sub txtAngle_Change()
    
    If Not IsNumeric(txtAngle.Text) Then Exit Sub
    
    If CInt(txtAngle.Text) > 88 Then
        cmdIncreaseAngle.Enabled = False
        txtAngle.SetFocus
    Else
        cmdIncreaseAngle.Enabled = True
    End If
    
    If CInt(txtAngle.Text) < 1 Then
        cmdDecreaseAngle.Enabled = False
        txtAngle.SetFocus
    Else
        cmdDecreaseAngle.Enabled = True
    End If
    
End Sub

Private Sub txtVelocity_Change()
    
    If Not IsNumeric(txtVelocity.Text) Then Exit Sub
    
    If CInt(txtVelocity.Text) > gintMaxVelocity - 5 Then
        cmdIncreaseVelocity.Enabled = False
        txtVelocity.SetFocus
    Else
        cmdIncreaseVelocity.Enabled = True
    End If
    
    If CInt(txtVelocity.Text) < 5 Then
        cmdDecreaseVelocity.Enabled = False
        txtVelocity.SetFocus
    Else
        cmdDecreaseVelocity.Enabled = True
    End If
    
End Sub

'Help functionality

Private Sub cmdFire_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdFire
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
      
End Sub

Private Sub txtScore_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = txtScore
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
      
End Sub

Private Sub txtAngle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = txtAngle
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdIncreaseAngle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdIncreaseAngle
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdDecreaseAngle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdDecreaseAngle
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub txtVelocity_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = txtVelocity
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdIncreaseVelocity_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdIncreaseVelocity
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdDecreaseVelocity_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdDecreaseVelocity
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdPlay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdPlay
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       Set frmHelp.ThisControl = cmdExit
       PopupMenu frmHelp.mnuBtnContextMenu
       Set frmHelp.ThisControl = Nothing
    End If
End Sub

Private Sub picDesktop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If Shift = 0 Then
        If Button = vbRightButton Then
           Set frmHelp.ThisControl = picDesktop
           PopupMenu frmHelp.mnuBtnContextMenu
           Set frmHelp.ThisControl = Nothing
        End If
    ElseIf Button = 0 Then
        If y < 0 Then
            frmBullet.HideBullet
        Else
            frmBullet.ChangeBkColor x, y
            mintHitIconIndex = frmBullet.Position(CInt(x) * PICTOSCREEN, CInt(y) * PICTOSCREEN)
            
            If mintHitIconIndex <> -1 Then
                DropIcon
                Exit Sub
            End If
        End If
    End If
    
End Sub

'Validations

Private Sub txtAngle_Validate(Cancel As Boolean)

    Cancel = InvalidAngle
    
    If Not Cancel Then frmGun.MoveGun CInt(txtAngle.Text)
    
End Sub

Private Sub txtVelocity_Validate(Cancel As Boolean)
    
    Cancel = InvalidVelocity
        
End Sub

Private Function InvalidAngle() As Boolean
    
    If Not IsNumeric(txtAngle.Text) Then
        MsgBox "Enter a numeric value between 0 and 89", _
            vbInformation, "Invalid input"
            InvalidAngle = True
    Else
        If CInt(txtAngle.Text) < 0 Then
            MsgBox "Minimum angle is 0", _
            vbInformation, "Invalid input"
            InvalidAngle = True
        ElseIf txtAngle.Text > 89 Then
            MsgBox "Maximum angle is 89", _
            vbInformation, "Invalid input"
            InvalidAngle = True
        End If
    End If
    
    If InvalidAngle Then txtAngle.SetFocus
    
End Function

Private Function InvalidVelocity() As Boolean
    
    If Not IsNumeric(txtVelocity.Text) Then
        MsgBox "Enter a numeric value between 0 and " & gintMaxVelocity, _
            vbInformation, "Invalid input"
            InvalidVelocity = True
    Else
        If CInt(txtVelocity.Text) < 0 Then
            MsgBox "Minimum value is 0", _
            vbInformation, "Invalid input"
            InvalidVelocity = True
        ElseIf txtVelocity.Text > gintMaxVelocity Then
            MsgBox "Maximum value is " & gintMaxVelocity, _
            vbInformation, "Invalid input"
            InvalidVelocity = True
        End If
    End If
    
    If InvalidVelocity Then txtVelocity.SetFocus
    
End Function
