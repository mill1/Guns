VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2520
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   2355
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   4920
      Begin VB.Timer timSplash 
         Interval        =   1000
         Left            =   120
         Top             =   225
      End
      Begin VB.Image imgBuhrmann 
         Height          =   960
         Left            =   555
         Picture         =   "frmSplash.frx":0442
         Stretch         =   -1  'True
         Top             =   465
         Width           =   960
      End
      Begin VB.Label lblGun 
         BackStyle       =   0  'Transparent
         Caption         =   "Guns"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   44.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1020
         Left            =   1920
         TabIndex        =   2
         Top             =   435
         Width           =   2250
      End
      Begin VB.Label lblApplication 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "The real desktop game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   525
         TabIndex        =   1
         Top             =   1530
         Width           =   3945
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub timSplash_Timer()
    
Static i As Integer
    
    If i > 0 Then
        Unload Me
        frmGuns.Init
    End If
    
    i = i + 1

End Sub
