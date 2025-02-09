VERSION 5.00
Begin VB.Form frmEegg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmEegg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "&Defaults"
      Height          =   390
      Left            =   195
      TabIndex        =   6
      Top             =   3645
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2460
      TabIndex        =   8
      Top             =   3645
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1320
      TabIndex        =   7
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3210
      Left            =   195
      TabIndex        =   9
      Top             =   165
      Width           =   3345
      Begin VB.TextBox txtBattlefieldWidth 
         Height          =   315
         Left            =   2475
         TabIndex        =   16
         Top             =   2655
         Width           =   600
      End
      Begin VB.TextBox txtIconMoveSteps 
         Height          =   315
         Left            =   2475
         TabIndex        =   2
         Top             =   1110
         Width           =   600
      End
      Begin VB.TextBox txtVelocityFactor 
         Height          =   315
         Left            =   2475
         TabIndex        =   1
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGravity 
         Height          =   315
         Left            =   2475
         TabIndex        =   0
         Top             =   330
         Width           =   600
      End
      Begin VB.TextBox txtTimer2Interval 
         Height          =   315
         Left            =   2475
         TabIndex        =   3
         Top             =   1500
         Width           =   600
      End
      Begin VB.TextBox txtBottom 
         Height          =   315
         Left            =   2475
         TabIndex        =   4
         Top             =   1890
         Width           =   600
      End
      Begin VB.TextBox txtGunPosition 
         Height          =   315
         Left            =   2475
         TabIndex        =   5
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Battlefield width (restart game)"
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   2700
         Width           =   2130
      End
      Begin VB.Label Label6 
         Caption         =   "Icon move steps"
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1125
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Velocity factor"
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Gravity"
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Bullet events in milliseconds"
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   1515
         Width           =   1980
      End
      Begin VB.Label Label3 
         Caption         =   "Bottom position"
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Gun position (restart game)"
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   2325
         Width           =   1950
      End
   End
End
Attribute VB_Name = "frmEegg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Go away

Private Sub cmdDefaults_Click()
    txtGravity.Text = 9.81
    txtVelocityFactor.Text = 8
    txtIconMoveSteps.Text = 30
    txtTimer2Interval.Text = 50
    txtBottom.Text = 0.8
    txtGunPosition.Text = 0.3
    txtBattlefieldWidth.Text = 0.75
End Sub

Private Sub Form_Load()
    txtGravity.Text = gsngGravity
    txtVelocityFactor.Text = gintVelocityFactor
    txtIconMoveSteps.Text = gintIconSteps
    txtTimer2Interval.Text = gintTimer2Interval
    txtBottom.Text = gsngBottom
    txtGunPosition.Text = gsngGunPosition
    txtBattlefieldWidth.Text = gsngBattlefieldWidth
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    On Error Resume Next
        
    Me.ValidateControls
    
    'Type mismatch
    If Err.Number = 380 Then
        Exit Sub
    End If
    
    On Error GoTo 0
        
    SaveSetting REG_APP, REG_SYSTEM, "Gravity", txtGravity.Text
    gsngGravity = GetValue(txtGravity.Text)
    
    SaveSetting REG_APP, REG_SYSTEM, "Velocity Factor", txtVelocityFactor.Text
    gintVelocityFactor = CInt(txtVelocityFactor.Text)
    
    gintMaxVelocity = gintScreenWidht / (1.7 + ((gintScreenWidht / 80) / gintVelocityFactor))
    
    SaveSetting REG_APP, REG_SYSTEM, "Icon Steps", txtIconMoveSteps.Text
    gintIconSteps = CInt(txtIconMoveSteps.Text)
    
    SaveSetting REG_APP, REG_SYSTEM, "Timer2 Interval", txtTimer2Interval.Text
    gintTimer2Interval = CInt(txtTimer2Interval.Text)
    
    SaveSetting REG_APP, REG_SYSTEM, "Bottom", txtBottom.Text
    gsngBottom = GetValue(txtBottom.Text)
    
    SaveSetting REG_APP, REG_SYSTEM, "Gun Position", txtGunPosition.Text
    gsngGunPosition = GetValue(txtGunPosition.Text)
    
    SaveSetting REG_APP, REG_SYSTEM, "Battlefield Width", txtBattlefieldWidth.Text
    gsngBattlefieldWidth = GetValue(txtBattlefieldWidth.Text)
    
    Unload Me
End Sub

Private Sub txtBattlefieldWidth_Validate(Cancel As Boolean)
    If Not IsNumeric(txtBattlefieldWidth.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    Else
        If CInt(txtBattlefieldWidth.Text) = 0 Then txtBattlefieldWidth.Text = 0.01
    End If
End Sub

Private Sub txtBottom_Validate(Cancel As Boolean)
    If Not IsNumeric(txtBottom.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtGravity_Validate(Cancel As Boolean)
    If Not IsNumeric(txtGravity.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtGunPosition_Validate(Cancel As Boolean)
    If Not IsNumeric(txtGunPosition.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    Else
        If CInt(txtGunPosition.Text) = 0 Then txtGunPosition.Text = 0.05
    End If
End Sub

Private Sub txtIconMoveSteps_Validate(Cancel As Boolean)
    If Not IsNumeric(txtIconMoveSteps.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    Else
        If CInt(txtIconMoveSteps.Text) = 0 Then txtIconMoveSteps.Text = 1
    End If
End Sub

Private Sub txtTimer2Interval_Validate(Cancel As Boolean)
    If Not IsNumeric(txtTimer2Interval.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    Else
        If CInt(txtTimer2Interval.Text) = 0 Then txtTimer2Interval.Text = 1
    End If
End Sub

Private Sub txtVelocityFactor_Validate(Cancel As Boolean)
    If Not IsNumeric(txtVelocityFactor.Text) Then
        MsgBox "Value is not numeric", vbInformation
        Cancel = True
    Else
        If CInt(txtVelocityFactor.Text) = 0 Then txtVelocityFactor.Text = 1
    End If
End Sub

