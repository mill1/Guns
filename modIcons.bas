Attribute VB_Name = "modIcons"
Option Explicit

Public Function GetOriginalDTIconCoors() As Boolean
'Return value determines whether AutoArrange is turned on...

    Dim i As Integer
    Dim strFileName As String
    Dim lnghNull As Long
    Dim lnghFile As Long
    'There is a difference between handles and pointers.
    Dim lnghFileMap As Long
    Dim lngpFileMap As Long
    Dim strTempPath As String
    Dim udtAutoArrange As POINTAPI
    
    'get windows temp path
    strTempPath = String(255, Chr$(0))
    'Get the temporary path
    GetTempPath 255, strTempPath
    'strip the rest of the buffer
    strTempPath = Left$(strTempPath, InStr(strTempPath, Chr$(0)) - 1)
    
    strFileName = strTempPath & "TEMP.ETN"
    ' Open file
    lnghFile = CreateFile(strFileName, GENERIC_READ Or GENERIC_WRITE, 0, _
                       ByVal lnghNull, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, lnghNull)
    ' Get handle
    lnghFileMap = CreateFileMappingTwo(lnghFile, ByVal lnghNull, _
                    PAGE_READWRITE, 0, 40, "MyMapping")
    ' Get pointer to memory representing file
    lngpFileMap = MapViewOfFile(lnghFileMap, FILE_MAP_WRITE, 0, 0, 0)

    'loop through icon count
    For i = 0 To gintDTIconsCount - 1
        
        SendMessageByLong glngDTSysListview32h, LVM_GETITEMPOSITION, i, lngpFileMap
        CopyMemoryTwo gudtOrigPointArray(i), lngpFileMap, 8
        
    Next i
            
    'Check if Auto Arranged is turned on...
    'Try to move the "My Computer" icon
    MoveIcon 0, gudtOrigPointArray(0).x + 1, gudtOrigPointArray(0).y
    
    'Get the coordinates (again)
    SendMessageByLong glngDTSysListview32h, LVM_GETITEMPOSITION, 0, lngpFileMap
    CopyMemoryTwo udtAutoArrange, lngpFileMap, 8

    If gudtOrigPointArray(0).x - udtAutoArrange.x = 0 Then
        GetOriginalDTIconCoors = True
    Else
        GetOriginalDTIconCoors = False
        'Move the "My Computer" icon back again
        MoveIcon i, gudtOrigPointArray(0).x, gudtOrigPointArray(0).y
    End If
    
    'Release resources back to windows
    FlushViewOfFile lngpFileMap, 40
    UnmapViewOfFile lngpFileMap
    CloseHandle lnghFileMap
    CloseHandle lnghFile
    'Thanx Paul Pavlic...

End Function

Public Sub SetOriginalDTIconCoors()

'Mimic Auto Arrange

    Dim i As Integer
    Dim i2 As Integer
    Dim lngX As Long
    Dim lngY As Long
    Dim intMaxNrOfIconsPerColumn As Integer
    
    intMaxNrOfIconsPerColumn = Fix((gintScreenHeight * 0.94) / (2.5 * gintIconHeight))
    
    lngX = gintIconWidth
    lngY = 2
    
    For i = 0 To gintDTIconsCount - 1
    
        gudtOrigPointArray(i).x = lngX
        gudtOrigPointArray(i).y = lngY
    
        MoveIcon i, lngX, lngY
        lngY = lngY + (2.5 * gintIconHeight)
        
        i2 = i2 + 1
        
        If i2 = intMaxNrOfIconsPerColumn Then
            lngX = lngX + (3 * gintIconWidth)
            lngY = 2
            i2 = 0
        End If
    Next

End Sub

Public Sub AutoArrange(ByVal lnghandle As Long, ByVal blnState As Boolean)
   
   'Doesn't work under W 95 and 98
   
   If blnState Then
      SetStyle lnghandle, LVS_AUTOARRANGE, 0
   Else
      SetStyle lnghandle, 0, LVS_AUTOARRANGE
   End If
   
End Sub

Public Sub SetStyle(ByVal plnghandle As Long, ByVal plngStyle As Long, ByVal plngStyleNot As Long)

Dim lngStyle As Long

   If Not plnghandle = 0 Then
      lngStyle = GetWindowLong(plnghandle, GWL_STYLE)
      lngStyle = lngStyle And Not plngStyleNot
      lngStyle = lngStyle Or plngStyle
      
      SetWindowLong plnghandle, GWL_STYLE, lngStyle
      
      SetWindowPos plnghandle, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
        'dus
   End If
   
End Sub


Public Sub MoveIcon(pintIndexIcon As Integer, plngX As Long, plngY As Long)
    
'Finally move the damn thing

    SendMessageByLong glngDTSysListview32h, LVM_SETITEMPOSITION, _
                    pintIndexIcon, plngX + plngY * &H10000
    
End Sub
