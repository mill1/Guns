Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As _
Long, ByVal hTemplateFile As Long) As Long
    
Public Declare Function CreateFileMappingTwo Lib "kernel32" Alias _
"CreateFileMappingA" (ByVal lnghFile As Long, lpFileMappigAttributes _
As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, _
ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long

Public Declare Function MapViewOfFile Lib "kernel32" (ByVal _
hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal _
dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal _
dwNumberOfBytesToMap As Long) As Long

Public Declare Sub CopyMemoryTwo Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal hpvSource&, ByVal cbCopy As Long)

Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress _
As Any) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long

Public Declare Function FlushViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long, _
ByVal dwNumberOfBytesToFlush As Long) As Long

Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long, ByVal nwidth As Long, ByVal nheight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SendMessageByLong& Lib "user32" Alias "SendMessageA" _
(ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
