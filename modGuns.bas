Attribute VB_Name = "modGuns"
Option Explicit

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public glngDTSysListview32h  As Long
Public gintDTIconsCount As Integer
Public gudtOrigPointArray() As POINTAPI
Public gudtPointArray() As POINTAPI
Public gudtPointArrayOld() As POINTAPI
Public gudtTrajectoryArray() As POINTAPI
Public gintScreenWidht As Integer
Public gintScreenHeight As Integer
Public gintIconWidth As Integer ' necessary to determine hit
Public gintIconHeight As Integer 'idem
Public gintMaxVelocity As Integer
Public gblnShots As Boolean
Public gblnSoundOff As Boolean
Public gstrHighScore As String
Public gsngGravity As Single
Public gintVelocityFactor As Integer
Public gintIconSteps As Integer
Public gintTimer2Interval As Integer
Public gsngBottom As Single
Public gsngGunPosition As Single
Public gsngBattlefieldWidth As Single
Public gintNrOfHits As Integer
Private mintSpace As Integer
Private mblnAutoArrange As Boolean

Public Sub Main()
    
    MinimizeAllWins
    
    GetGlobals
    
    'bottleneck is the screen height when determining space between icons
    mintSpace = (gintScreenHeight * gsngBattlefieldWidth) / gintDTIconsCount
    
    If Platform = "Windows 95/98" Then
        
        mblnAutoArrange = GetOriginalDTIconCoors
        
        If mblnAutoArrange Then
            MsgBox _
            "Please turn off the Auto Arrange property of your desktop." & vbCrLf & _
            "(right-click the desktop ; Arrange Icons ; Auto Arrange)", vbInformation
        End If
    Else
        'There's no way to check the desktop's auto arrange property since
        'we can't establish the icon's coordinates...
        'We can, however, turn it off on win NT and win 2000.
        AutoArrange glngDTSysListview32h, False
        
        SetOriginalDTIconCoors
    End If
    
    'In most case only 0.7 of the icon width is used for dislay
    'But we needed the full width to set the original positions (win NT and win 2000).
    gintIconWidth = gintIconWidth * 0.7
    
    'In case user decides to exit immediately
    gudtPointArray = gudtOrigPointArray
    
    frmSplash.Show
    
End Sub

Private Sub GetGlobals()
    
    'glngDTSysListview32h  is the handle of the Desktop's syslistview32
    glngDTSysListview32h = FindWindow("progman", vbNullString)
    glngDTSysListview32h = FindWindowEx(glngDTSysListview32h, 0, _
                            "shelldll_defview", vbNullString)
    glngDTSysListview32h = FindWindowEx(glngDTSysListview32h, 0, _
                            "syslistview32", vbNullString)
    
    gintDTIconsCount = SendMessageByLong(glngDTSysListview32h, LVM_GETTITEMCOUNT, 0, 0)
    
    If gintDTIconsCount = 0 Then
        DoEvents
        gintDTIconsCount = SendMessageByLong(glngDTSysListview32h, LVM_GETTITEMCOUNT, 0, 0)
        If gintDTIconsCount = 0 Then
            MsgBox "Restart the game or get some desktop shotcut icons...", vbInformation
            End
        End If
    End If
    
    ReDim gudtPointArray(gintDTIconsCount - 1)
    ReDim gudtOrigPointArray(gintDTIconsCount - 1)
    
    'Get the settings...
    GetSettings
    
    gintScreenWidht = GetSystemMetrics(SM_XSCREEN)
    gintScreenHeight = GetSystemMetrics(SM_YSCREEN)
    
    gintMaxVelocity = gintScreenWidht / (1.7 + ((gintScreenWidht / 80) / gintVelocityFactor))
    
    gintIconWidth = GetSystemMetrics(SM_XICON)
    gintIconHeight = GetSystemMetrics(SM_YICON)
    
End Sub

Public Sub GetSettings()

    Dim strSetting As String

    'Type of game highscore
    strSetting = GetSetting(REG_APP, REG_SETTINGS, "HighScoreInShots", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SETTINGS, "HighScoreInShots", 1
        gblnShots = True
    Else
        gblnShots = CBool(strSetting)
    End If
    
    frmMenu.mnuHighScoreInShots.Checked = IIf(gblnShots, True, False)
    
    'Sound
    strSetting = GetSetting(REG_APP, REG_SETTINGS, "Sound Off", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SETTINGS, "Sound Off", 0
        gblnSoundOff = False
    Else
        gblnSoundOff = CBool(strSetting)
    End If
    
    frmMenu.mnuSoundOff.Checked = IIf(gblnSoundOff, True, False)
    
    GetHighScore

    'System settings
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Gravity", "?")
        
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Gravity", 9.81
        gsngGravity = 9.81
    Else
        gsngGravity = GetValue(strSetting)
    End If
    
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Velocity Factor", "?")
        
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Velocity Factor", 8
        gintVelocityFactor = 8
    Else
        gintVelocityFactor = CInt(strSetting)
    End If
    
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Icon Steps", "?")
        
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Icon Steps", 30
        gintIconSteps = 30
    Else
        gintIconSteps = CInt(strSetting)
    End If
    
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Timer2 Interval", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Timer2 Interval", 50
        gintTimer2Interval = 50
    Else
        gintTimer2Interval = CInt(strSetting)
    End If
    
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Bottom", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Bottom", 0.8
        gsngBottom = 0.8
    Else
        gsngBottom = GetValue(strSetting)
    End If

    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Gun Position", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Gun Position", 0.3
        gsngGunPosition = 0.3
    Else
        gsngGunPosition = GetValue(strSetting)
    End If
    
    strSetting = GetSetting(REG_APP, REG_SYSTEM, "Battlefield Width", "?")
    
    If strSetting = "?" Then
        SaveSetting REG_APP, REG_SYSTEM, "Battlefield Width", 0.75
        gsngBattlefieldWidth = 0.75
    Else
        gsngBattlefieldWidth = GetValue(strSetting)
    End If
    
End Sub

Public Sub GetHighScore()

    Dim strSetting As String

    'Highscore itself
    If gblnShots Then
    
        strSetting = GetSetting(REG_APP, REG_SHOTS, gintDTIconsCount, "?")
        
        If strSetting = "?" Then
            SaveSetting REG_APP, REG_SHOTS, gintDTIconsCount, 999
            gstrHighScore = "Highscore: -"
        Else
            gstrHighScore = "Highscore: " & strSetting & " shots"
        End If
    
    Else
        
        strSetting = GetSetting(REG_APP, REG_TIME, gintDTIconsCount, "?")
        
        If strSetting = "?" Then
            SaveSetting REG_APP, REG_TIME, gintDTIconsCount, 999
            gstrHighScore = "Highscore: -"
        Else
            gstrHighScore = "Highscore: " & strSetting & " seconds"
        End If

    End If

End Sub

Public Sub MinimizeAllWins()
  
    Dim lngHwnd As Long
   
    lngHwnd = FindWindow("Shell_TrayWnd", vbNullString)
    PostMessage lngHwnd, WM_COMMAND, MIN_ALL, 0&
  
    DoEvents
  
End Sub
  
Public Sub RestoreAllWins()

    Dim lngHwnd As Long

    lngHwnd = FindWindow("Shell_TrayWnd", vbNullString)
    PostMessage lngHwnd, WM_COMMAND, MIN_ALL_UNDO, 0&

End Sub
    
Public Function Platform() As String
    Dim OSInfo As OSVERSIONINFO
    Dim lngRet As Long
    
    'Set the structure size
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'Get the Windows version
    lngRet = GetVersionEx(OSInfo)
    
    'Check for errors
    If lngRet = 0 Then
        Platform = "Unknown"
    End If
    
    Select Case OSInfo.dwPlatformId
        Case 0
            Platform = "Windows 32s "
        Case 1
            Platform = "Windows 95/98"
        Case 2
            Platform = "Windows NT "
    End Select

End Function

Public Sub GetRandomizedArray()
    
    Dim i As Integer

    Do
        'Set the position parameters in pixels in the battlefield
        For i = 0 To gintDTIconsCount - 1
            'Generate random value between 1 and desktop widht.
            Randomize
            gudtPointArray(i).x = _
                Int(((gintScreenWidht * (gsngBattlefieldWidth * 0.8)) * Rnd) + gintIconWidth)
            Randomize
            'Start the battle at 1/5 from the top
            gudtPointArray(i).y = _
                Int(((gintScreenHeight * 0.5) * Rnd) + (gintScreenHeight * 0.2))
        Next
    Loop While Overlapping
    
End Sub

Public Sub GetTrajectory(ByVal intAngle As Integer, ByVal plngVelocity As Long)
    
    'Determine the trajectory of the bullet before it's even fired
    
    Dim i As Integer
    Dim sngX As Currency
    Dim sngY As Single
    Dim sngAngle As Single
    Dim lngStartX As Long
    Dim lngStartY As Long
    Dim strPope As String
    Dim strCatholic As String

    lngStartX = GetStartX
    lngStartY = GetStartY(intAngle)

    ReDim gudtTrajectoryArray(1)
    
    sngAngle = PI - (intAngle * (PI / 180)) '(9.81 = gravity on earth)
    
    plngVelocity = plngVelocity * 2
    
    Do While strPope = strCatholic
    
        i = (UBound(gudtTrajectoryArray) - 1)
        
        sngX = plngVelocity * i * Cos(sngAngle)
        sngY = GetY(plngVelocity, i, sngAngle, gsngGravity)
        
        sngX = sngX / 100
        sngY = sngY / 100
        
        gudtTrajectoryArray(i).x = lngStartX + sngX
        gudtTrajectoryArray(i).y = lngStartY - sngY
        
        'End when when bullet leaves the screen on the left or the bottom
        If gudtTrajectoryArray(i).x <= 0 Or _
            gudtTrajectoryArray(i).y >= CInt(frmGuns.picDesktop.Height) Then
            
            ReDim Preserve gudtTrajectoryArray(i - 1)
            Exit Do
        Else
            ReDim Preserve gudtTrajectoryArray(UBound(gudtTrajectoryArray) + 1)
        End If
    Loop

End Sub
    
Public Function GetY(plngVelocity As Long, pintCount As Integer, psngAngle As Single, psngGravity As Single) As Single
        
    GetY = plngVelocity * pintCount * Sin(psngAngle) - psngGravity * pintCount ^ 2
    
End Function
    
Public Function GetStartX() As Long
    On Error Resume Next
    GetStartX = frmGuns.picDesktop.Width * gsngBattlefieldWidth * 1.28
End Function

Public Function GetStartY(pintAngle As Integer) As Long

    GetStartY = (frmGun.Top - frmBullet.Height) / PICTOSCREEN + ((frmGun.linGun.Y2) / 4)
    
    If pintAngle > 45 Then GetStartY = GetStartY - (pintAngle / 10) 'one evening per line of code:
    
End Function

Public Function Overlapping() As Boolean
'Make sure the icons aren't on top of each other..

    Dim i1 As Integer
    Dim i2 As Integer
    
    For i1 = LBound(gudtPointArray) To UBound(gudtPointArray)
        For i2 = LBound(gudtPointArray) To UBound(gudtPointArray)
            
            If i1 <> i2 Then
                
                If Abs(gudtPointArray(i1).x - gudtPointArray(i2).x) < mintSpace And _
                  Abs(gudtPointArray(i1).y - gudtPointArray(i2).y) < mintSpace Then
                    Overlapping = True
                    Exit Function
                End If
            End If
        Next
    Next
    
End Function

Public Function GetValue(ByVal pstrValue As String) As Single
' Damn regional settings: you don't want a gravity of 981 or .0981

    Dim intPos As Integer
    
    intPos = InStrRev(pstrValue, ".")
    
    If intPos > 0 Then
        pstrValue = Replace(pstrValue, ".", "")
    Else
       intPos = InStrRev(pstrValue, ",")
       pstrValue = Replace(pstrValue, ",", "")
    End If

    If intPos > 0 Then
        GetValue = CSng(pstrValue) / (10 ^ (Len(pstrValue) - intPos + 1))
    Else
        GetValue = CSng(pstrValue)
    End If

End Function

Public Function GetDecimalSeparator() As String



'    Dim strDecimal As String
'    Dim dblTemp As Double
'
'    strDecimal = "10.5"
'
'    'Provoke the regional settings and loose grouping characters
'    dblTemp = Format(strDecimal, "###0.000")
'
'    GetDecimalSeparator = IIf(CStr(dblTemp) = strDecimal, ".", ",")

End Function

Public Sub Bye()
    Unload frmBullet
    Unload frmGuns
    
    'Restore the windows
    RestoreAllWins
    End
End Sub

Public Function NewLeft(pintStep As Integer, pintIndexIcon As Integer) As Long

    NewLeft = gudtPointArrayOld(pintIndexIcon).x + (gudtPointArray(pintIndexIcon).x - _
            gudtPointArrayOld(pintIndexIcon).x) * (pintStep / gintIconSteps)
        
End Function

Public Function NewTop(pintStep As Integer, pintIndexIcon As Integer) As Long
    
    NewTop = gudtPointArrayOld(pintIndexIcon).y + (gudtPointArray(pintIndexIcon).y - _
            gudtPointArrayOld(pintIndexIcon).y) * (pintStep / gintIconSteps)

End Function
