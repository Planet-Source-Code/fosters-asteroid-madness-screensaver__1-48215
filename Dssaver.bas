Attribute VB_Name = "dSSaver"
Option Explicit
'----------------------------------------------------------------------
' Application Specific Constants...
'----------------------------------------------------------------------
'''Public gSpriteCollection As Collection              ' collection of active sprites...

Public gSpriteCount As Long                         ' count of active sprites..
Public gTracers As Boolean                          ' tracers option (sprite doesn't clean up trails)
Public gRefreshRate As Long                         ' sprite animation frame movement rate
Public gRefreshRND As Boolean                       ' random refresh rate option
Public gSpriteSize As Long                          ' relative sprite size option
Public gSizeRND As Boolean                          ' randomize sprite size
Public gSpriteSpeed As Long                         ' active sprite velosity
Public gSpeedRND As Boolean                         ' randomize sprite speed
Public gSprite As ResBitmap                         ' bitmap resource loading bucket

Public Const gREGKEY_APPROOT = "SOFTWARE\VB 5 Samples\VB 5 Saver" ' ScreenSaver registry subkey
Public Const gREGVAL_SPRITECOUNT = "SpriteCount"    ' Sprite count registry setting key
Public Const DEF_SPRITECOUNT = 8                    ' default
Public Const MIN_SPRITECOUNT = 1                    ' min possible value
Public Const MAX_SPRITECOUNT = 30                   ' max possible value

Public Const gREGVAL_TRACERSON = "TracersOn"        ' Tracers on regkey

Public Const gREGVAL_REFRESHRATE = "RefreshRate"    ' Animation refresh rate registry setting key
Public Const MIN_REFRESHRATE = 1                    ' 1 / 1000 sec
Public Const MAX_REFRESHRATE = 100                  ' 1 / 10   sec
Public Const gREGVAL_RATERANDOM = "RateRandom"      ' Random refresh rate registry setting key

Public Const gREGVAL_SPRITESIZE = "SpriteSize"      ' Sprite size registry setting key
Public Const MIN_SPRITESIZE = 25                    ' 25% normal size
Public Const MAX_SPRITESIZE = 150                   ' 150% normal size
Public Const gREGVAL_SIZERANDOM = "SizeRandom"      ' Sprite size random registry setting key

Public Const gREGVAL_SPRITESPEED = "SpriteSpeed"    ' Sprite speed registry setting key
Public Const MIN_SPRITESPEED = 1                    ' Move in 1 pixel increments
Public Const MAX_SPRITESPEED = 50                   ' Move in 50 pixel increments
Public Const gREGVAL_SPEEDRANDOM = "SpeedRandom"    ' Sprite speed random registry setting key
Public Const sTRUE = "TRUE"                         ' Boolean TRUE registry value
Public Const sFALSE = "FALSE"                       ' Boolean FALSE registry value

Public Const BASE_MASS = 100                        ' Relative base mass for sprite size

'----------------------------------------------------------------------
'Public API Declares...
'----------------------------------------------------------------------
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal fShow As Integer) As Integer
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Any) As Long
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'----------------------------------------------------------------------
'Public Constants...
'----------------------------------------------------------------------
Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Public Const HWND_TOPMOST = -1&
Public Const HWND_TOP = 0&
Public Const HWND_BOTTOM = 1&

Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' Windows messages...
Public Const WM_PAINT = &HF&
Public Const WM_ACTIVATEAPP = &H1C&
Public Const SW_SHOWNOACTIVATE = 4&

' Get Windows Long Constants
Public Const GWL_USERDATA = (-21&)
Public Const GWL_WNDPROC = (-4&)

' ScreenSaver Running Modes
Public Const RM_NORMAL = 1
Public Const RM_CONFIGURE = 2
Public Const RM_PREVIEW = 4

' Reg Create Type Values...
Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore

' Reg Key Security Options...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const ERROR_SUCCESS = 0                  ' Return Value...
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

'----------------------------------------------------------------------
'Public Type Defs...
'----------------------------------------------------------------------
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Public Type ResBitmap
    ResID As Long
    Sprite As StdPicture
End Type
