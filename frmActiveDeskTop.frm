VERSION 5.00
Begin VB.Form frmActiveDeskTop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Active Desktop"
   ClientHeight    =   2220
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5100
   Icon            =   "frmActiveDeskTop.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Don't display this message again"
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   1635
      Width           =   2700
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   3765
      TabIndex        =   0
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message..."
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   810
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "This game is best played when the Active Desktop is not viewed as a web page."
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   3810
   End
End
Attribute VB_Name = "frmActiveDeskTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOK_Click()
    
    If chkDisplay.Value = vbChecked Then
        SaveSetting REG_APP, REG_SYSTEM, "Display Message", 0
    Else
        SaveSetting REG_APP, REG_SYSTEM, "Display Message", 1
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()

    If Platform = "Windows 95/98" Then
        lblMessage.Caption = "(Start button ; Settings ; Active Desktop ; View as Web Page)"
    Else
        lblMessage.Caption = "(Right-click desktop ; Active Desktop ; Customize My Desktop ;" & vbCrLf & _
                                "Uncheck box 'Show Web content on my Active Desktop')"
    End If
    
End Sub

