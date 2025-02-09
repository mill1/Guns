Attribute VB_Name = "modConstants"
Option Explicit

'No matter which resolution the ratio between a maximized form and the nr of pixels
'of the screen is always 15 for either width or height
Public Const SCREENTOFORM = 15
Public Const PICTOSCREEN = 60
Public Const PI = 3.141592654

'API Constants
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const PAGE_READWRITE As Long = &H4
Public Const FILE_MAP_WRITE = &H2
Public Const SCRCOPY = &HCC0020
Public Const LVM_FIRST = &H1000
Public Const LVM_SETITEMPOSITION& = (&H1000 + 15)
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_GETTITEMCOUNT& = (&H1000 + 4)
Public Const WM_COMMAND As Long = &H111
Public Const SM_XSCREEN = 0 'X Size of screen
Public Const SM_YSCREEN = 1 'Y Size of Screen
Public Const SM_XICON = 11 'Width of standard icon
Public Const SM_YICON = 12 'Height of standard icon
Public Const MIN_ALL As Long = 419
Public Const MIN_ALL_UNDO As Long = 416
Public Const LVS_AUTOARRANGE = &H100
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const GWL_STYLE = (-16)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_FRAMECHANGED = &H20

'Registry Settings
Public Const REG_APP = "fgth"
Public Const REG_SETTINGS = "Settings"
Public Const REG_SYSTEM = "System"
Public Const REG_SHOTS = "HighScores\Shots"
Public Const REG_TIME = "HighScores\Time"

