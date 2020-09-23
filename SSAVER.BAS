Attribute VB_Name = "mSSaver"
Option Explicit

Public DisplayHwnd As Long                      ' Hwnd of display form
Public DispRec As RECT                          ' Rectangle values of display form
Public PrevWndProc As Long                      ' Previous window proc (used in subclassing)
Public RunMode As Long                          ' Screen saver running mode (run, preview, setup)
Public DeskBmp As BITMAP                        ' Bitmap copy of the desktop
Public DeskDC As Long                           ' Desktop device context handle
Public cmd As String
Sub InitRegistry()
    If Len(GetSetting(App.Title, "Settings", "Ships")) = 0 Then
        SaveSetting App.Title, "Settings", "Ships", 30
    End If
    If Len(GetSetting(App.Title, "Settings", "Alpha")) = 0 Then
        SaveSetting App.Title, "Settings", "Alpha", 1
    End If
    iNumShips = CInt(GetSetting(App.Title, "Settings", "Ships"))
    bAlphaTrails = IIf(CInt(GetSetting(App.Title, "Settings", "Alpha")) = 1, True, False)
End Sub
'-----------------------------------------------------------------
Sub Main()
'-----------------------------------------------------------------
    Dim rc As Long                              ' function return code
                              ' command line arguments
    Dim Style As Long                           ' window style of display form
'-----------------------------------------------------------------
    App.Title = "Asteroid Madness"
    InitRegistry
    
    cmd = LCase$(Trim$(Command$))               ' copy command line parameters in lowercase...
    'cmd = "/c"
    Select Case Mid$(cmd, 1, 2)                 ' Parse 1st 2 chars from cmd line
    '------------------------------------------------------------
    Case "", "/s"   '[Normal Run Mode]            Run as Screen Saver on desktop.
    '------------------------------------------------------------
        RunMode = RM_NORMAL                     ' Store screen saver's run mode

        GetWindowRect GetDesktopWindow(), DispRec ' Get DeskTop Rectangle dimentions
        
        Load Form1                          ' Load Screen saver
#If DebugOn Then                                ' Do this only when debugging
        Form1.Show
#Else                                           ' Do this only when NOT debugging
        SetWindowPos Form1.hwnd, _
                     HWND_TOPMOST, 0&, 0&, DispRec.Right, DispRec.Bottom, _
                     SWP_SHOWWINDOW             ' Size window and make top most
#End If
    '------------------------------------------------------------
     Case "/p"      '[Win 95 & NT 4 Preview Mode] Run inside of the Screen Saver Config Viewer.
    '------------------------------------------------------------
    '- Run the screen saver in the windows preview dialog, YES in VB!
    '------------------------------------------------------------
        RunMode = RM_PREVIEW                    ' Store screen saver's run mode...
        
        DisplayHwnd = GetHwndFromCmd(cmd)       ' ** Get HWND of Preview  DeskTop
        GetClientRect DisplayHwnd, DispRec      ' Get Display Rectangle dimentions
        
        Load Form1                          ' Load Screen saver form
        Form1.Caption = "Preview"           ' Consistant with Win 95 screen savers(what the heck)
        
        Style = GetWindowLong(Form1.hwnd, GWL_STYLE) ' ** Get current window style
        Style = Style Or WS_CHILD                        ' ** Append "WS_CHILD" style to the hWnd window style
        SetWindowLong Form1.hwnd, GWL_STYLE, Style   ' ** Add new style to window
        
        SetParent Form1.hwnd, DisplayHwnd   ' ** Set preview window as parent window
        SetWindowLong Form1.hwnd, GWL_HWNDPARENT, DisplayHwnd ' ** Save the hWnd Parent in hWnd's window struct.
        
        ' ** Show screensaver in the preview window...
        SetWindowPos Form1.hwnd, _
                     HWND_TOP, 0&, 0&, DispRec.Right, DispRec.Bottom, _
                     SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    '------------------------------------------------------------
    ' lines prefixed with ** are necessary for the preview dialog to work correctly.
    '------------------------------------------------------------
    Case "/c"       '[ScreenSaver Configuration Mode] Run Screen Saver Settings Dialog.
    '------------------------------------------------------------
        
        Load Form2
        Form2.Show
    '------------------------------------------------------------
    Case Else
    '------------------------------------------------------------
#If DebugOn Then                                ' Do this only when debugging
        MsgBox "Unknown Command Line Param: [" & Command$ & "]" ' Debug/display unknown param...
#End If
    End Select
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'------------------------------------------------------------
Public Function SubWndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------------------
'- Subclassing not implemented but reservered for furture use...
'------------------------------------------------------------
'    Select Case MSG
'    Case WM_PAINT
'        SubWndProc = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
'        PaintDeskDC DeskDC, DeskBmp, hwnd
'        Exit Function
'    End Select

'    SubWndProc = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
'------------------------------------------------------------
End Function
'------------------------------------------------------------

'-----------------------------------------------------------------
Private Function GetHwndFromCmd(cmd As String) As Long
'-----------------------------------------------------------------
    Dim Str As String                           ' substring variable
    Dim lenStr As Long                          ' length of substring
    Dim Idx As Long                             ' Index variable
'-----------------------------------------------------------------
    Str = Trim$(cmd)                            ' copy command line
    lenStr = Len(Str)                           ' get size of string
    
    For Idx = lenStr To 1 Step -1               ' for each char in string
        Str = Right$(Str, Idx)                  ' chop off the rightmost char
        If IsNumeric(Str) Then                  ' if substring is numeric then value is an hWnd
            GetHwndFromCmd = Val(Str)           ' return hWnd value
            Exit For                            ' exit for loop
        End If
    Next
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, "Visual Basic 5.0 - Screen Saver...", _
               vbCrLf & "Building Applications in Visual Basic 5.0", 0
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Private Sub AssertRC(bool As Boolean, rc As Long, fcnName As String)
'-----------------------------------------------------------------
#If DebugOn Then
    If Not bool Then
        MsgBox "Assertion Failed::" & vbCrLf & _
               "       In Module:: " & fcnName & vbCrLf & _
               "     Return Code:: " & CStr(rc), vbCritical
    End If
#End If
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'------------------------------------------------------------
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                               ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
'------------------------------------------------------------
GetKeyError:    ' Cleanup After An Error Has Occured...
'------------------------------------------------------------
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------

'------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
'------------------------------------------------------------
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
'------------------------------------------------------------
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
    
    '------------------------------------------------------------
    '- Create/Modify Key Value...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...
    
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, Len(SubKeyValue))   ' Create/Modify Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key
    
    UpdateKey = True                                    ' Return Success
    Exit Function                                       ' Exit
'------------------------------------------------------------
CreateKeyError:
'------------------------------------------------------------
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------

'------------------------------------------------------------
Public Sub SaveSettings()
'------------------------------------------------------------
    Dim RegVal As String                                ' String value of registry key
    Dim lRegVal As Long                                 ' long value of registry key
'------------------------------------------------------------
    ' Save Sprite Count Value
    RegVal = CStr(gSpriteCount)
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITECOUNT, RegVal)
    
    ' Save Tracers on Value
    RegVal = sFALSE
    If gTracers Then RegVal = sTRUE
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_TRACERSON, RegVal)
    
    ' Save Refresh Rate Value
    RegVal = CStr(gRefreshRate)
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_REFRESHRATE, RegVal)
    
    ' Save Rate Random Value
    RegVal = sFALSE
    If gRefreshRND Then RegVal = sTRUE
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_RATERANDOM, RegVal)
    
    ' Save Sprite Size Value
    RegVal = CStr(gSpriteSize)
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITESIZE, RegVal)
    
    ' Save Size Random Value
    RegVal = sFALSE
    If gSizeRND Then RegVal = sTRUE
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SIZERANDOM, RegVal)
    
    ' Save Sprite Speed Value
    RegVal = CStr(gSpriteSpeed)
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITESPEED, RegVal)
    
    ' Save Speed Random Value
    RegVal = sFALSE
    If gSpeedRND Then RegVal = sTRUE
    Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPEEDRANDOM, RegVal)
'------------------------------------------------------------
End Sub
'------------------------------------------------------------

'------------------------------------------------------------
Public Sub LoadSettings()
'------------------------------------------------------------
    Dim RegVal As String
    Dim iRegVal As Long
'------------------------------------------------------------
    ' Get Sprite Count Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITECOUNT, RegVal)
    gSpriteCount = Val(RegVal)
    If (gSpriteCount < MIN_SPRITECOUNT) Then gSpriteCount = DEF_SPRITECOUNT ' Default value.
    If (gSpriteCount > MAX_SPRITECOUNT) Then gSpriteCount = MAX_SPRITECOUNT
    
    ' Get Tracers on Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_TRACERSON, RegVal)
    gTracers = (RegVal = sTRUE)

    ' Get Refresh Rate Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_REFRESHRATE, RegVal)
    gRefreshRate = Val(RegVal)
    If (gRefreshRate < MIN_REFRESHRATE) Then gRefreshRate = MAX_REFRESHRATE ' Default value ...fast
    If (gRefreshRate > MAX_REFRESHRATE) Then gRefreshRate = MAX_REFRESHRATE
    
    ' Get Rate Random Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_RATERANDOM, RegVal)
    gRefreshRND = (RegVal = sTRUE)
       
    ' Get Sprite Size Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITESIZE, RegVal)
    gSpriteSize = Val(RegVal)
    If (gSpriteSize < MIN_SPRITESIZE) Then gSpriteSize = MIN_SPRITESIZE
    If (gSpriteSize > MAX_SPRITESIZE) Then gSpriteSize = MAX_SPRITESIZE
    
    ' Get Size Random Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SIZERANDOM, RegVal)
    gSizeRND = (RegVal = sTRUE) Or (RegVal = "")    ' Default to TRUE
    
    ' Get Sprite Speed Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPRITESPEED, RegVal)
    gSpriteSpeed = Val(RegVal)
    If (gSpriteSpeed < MIN_SPRITESPEED) Then gSpriteSpeed = MIN_SPRITESPEED
    If (gSpriteSpeed > MAX_SPRITESPEED) Then gSpriteSpeed = MAX_SPRITESPEED
    
    ' Get Speed Random Value
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, gREGVAL_SPEEDRANDOM, RegVal)
    gSpeedRND = (RegVal = sTRUE)
'------------------------------------------------------------
End Sub
'------------------------------------------------------------
