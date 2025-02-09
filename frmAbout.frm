VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Guns.exe"
   ClientHeight    =   2970
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2049.947
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4470
      TabIndex        =   0
      Top             =   2445
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1563.343
      Y2              =   1563.343
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   1050
      TabIndex        =   2
      Top             =   930
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   585
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Check out the code at http://emielnijhuis.tripod.com"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   255
      TabIndex        =   3
      Top             =   2490
      Width           =   3945
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblTitle.Caption = App.Title & "  (Freeware)"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription = "Guns is a shooting game that has been tested on most Windows platforms." & vbCrLf & _
                    "I made it 'cause it's more fun than studying for a MCSD course." & vbCrLf & vbCrLf & App.Comments
End Sub


