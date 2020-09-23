Attribute VB_Name = "PaintSup"
Option Explicit

'-----------------------------------------------------------------
Public Function ShrinkBmp(dispHdc As Long, hBmp As Long, RatioX As Single, RatioY As Single) As Long
'-----------------------------------------------------------------
    Dim hBmpOut As Long                             ' output bitmap handle
    Dim bm1 As BITMAP, bm2 As BITMAP                ' temporary bitmap structs
    Dim hdcMem1 As Long, hdcMem2 As Long            ' temporary memory bitmap handles...
'-----------------------------------------------------------------
    hdcMem1 = CreateCompatibleDC(dispHdc)           ' create mem DC compatible to the display DC
    hdcMem2 = CreateCompatibleDC(dispHdc)           ' create mem DC compatible to the display DC
  
    GetObject hBmp, LenB(bm1), bm1                  ' select bitmap object
  
    LSet bm2 = bm1                                  ' copy bitmap object
  
    bm2.bmWidth = CLng(bm2.bmWidth * RatioX)        ' scale output bitmap width
    bm2.bmHeight = CLng(bm2.bmHeight * RatioY)      ' scale output bitmap height
    bm2.bmWidthBytes = ((((bm2.bmWidth * bm2.bmBitsPixel) + 15) \ 16) * 2) ' calculate bitmap width bytes

    hBmpOut = CreateBitmapIndirect(bm2)             ' create handle to output bitmap indirectly from new bm2
    
    SelectObject hdcMem1, hBmp                      ' select original bitmap into mem dc
    SelectObject hdcMem2, hBmpOut                   ' select new bitmap into mem dc

    ' stretch old bitmap into new bitmap
    StretchBlt hdcMem2, 0, 0, bm2.bmWidth, bm2.bmHeight, _
               hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy
    
    DeleteDC hdcMem1                                ' delete memory dc
    DeleteDC hdcMem2                                ' delete memory dc

    ShrinkBmp = hBmpOut                             ' return handle to new bitmap
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub InitDeskDC(OutHdc As Long, OutBmp As BITMAP, DispRec As RECT)
'-----------------------------------------------------------------
    Dim DskHwnd As Long                     ' hWnd of desktop
    Dim DskRect As RECT                     ' rect size of desktop
    Dim DskHdc As Long                      ' hdc of desktop
    Dim hOutBmp As Long                     ' handle to output bitmap
    Dim rc As Long                          ' function return code
'-----------------------------------------------------------------
    DskHwnd = GetDesktopWindow()            ' Get src - HWND of Desktop
    DskHdc = GetWindowDC(DskHwnd)           ' Get src HDC - Handle to device context
    rc = GetWindowRect(DskHwnd, DskRect)    ' Get src Rectangle dimentions
    
    With DispRec
        ' Create handle to compatible output bitmap
        hOutBmp = CreateCompatibleBitmap(DskHdc, (.Right - .Left + 1), (.Bottom - .Top + 1))
    
        rc = GetObject(hOutBmp, Len(OutBmp), OutBmp)    ' Get handle to bitmap
        OutHdc = CreateCompatibleDC(DskHdc)             ' Create compatible hdc
        rc = SelectObject(OutHdc, hOutBmp)              ' copy bitmap structure into output dc
                
        rc = StretchBlt(OutHdc, 0, 0, _
                       (.Right - .Left + 1), _
                       (.Bottom - .Top + 1), _
                        DskHdc, 0, 0, _
                       (DskRect.Right - DskRect.Left + 1), _
                       (DskRect.Bottom - DskRect.Top + 1), _
                        vbSrcCopy)          ' Paint bitmap desk dc to output dc
    End With
    
    rc = DeleteObject(hOutBmp)              ' delete handle to output bitmap
    rc = ReleaseDC(DskHwnd, DskHdc)         ' Clean up - Release src HDC
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub PaintDeskDC(InHdc As Long, InBmp As BITMAP, OutHwnd As Long)
'-----------------------------------------------------------------
    Dim OutRect As RECT                     ' rect. size of output window
    Dim OutHdc As Long                      ' hdc of output window
    Dim rc As Long                          ' function return code
'-----------------------------------------------------------------
    rc = GetClientRect(OutHwnd, OutRect)    ' Get Dest Rectangle dimentions
    OutHdc = GetWindowDC(OutHwnd)           ' get Dest HDC
        
    With OutRect
        ' Paint the desktop picture to the output window...
        rc = StretchBlt(OutHdc, 0, 0, _
                       (.Right - .Left + 1), _
                       (.Bottom - .Top + 1), _
                       InHdc, 0, 0, _
                       InBmp.bmWidth, InBmp.bmHeight, vbSrcCopy)
    End With
    
    rc = ReleaseDC(OutHwnd, OutHdc)         ' Clean up - Release src HDC
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub DelDeskDC(OutHdc As Long)
'-----------------------------------------------------------------
    Dim rc As Long
'-----------------------------------------------------------------
    
    rc = DeleteDC(OutHdc)          ' Clean up - Release src HDC
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub DrawTransparentBitmap(lHDCDest As Long, _
                                 lBmSource As Long, _
                                 lMaskColor As Long, _
                                 Optional lDestStartX As Long, _
                                 Optional lDestStartY As Long, _
                                 Optional lDestWidth As Long, _
                                 Optional lDestHeight As Long, _
                                 Optional lSrcStartX As Long, _
                                 Optional lSrcStartY As Long, _
                                 Optional BkGrndHdc As Long)
'-----------------------------------------------------------------
    Dim udtBitMap As BITMAP
    Dim lColorRef As Long 'COLORREF
    Dim lBmAndBack As Long 'HBITMAP
    Dim lBmAndObject As Long
    Dim lBmAndMem As Long
    Dim lBmSave As Long
    Dim lBmBackOld As Long
    Dim lBmObjectOld As Long
    Dim lBmMemOld As Long
    Dim lBmSaveOld As Long
    Dim lHDCMem As Long 'HDC
    Dim lHDCBack As Long
    Dim lHDCObject As Long
    Dim lHDCTemp As Long
    Dim lHDCSave As Long
    Dim udtSize As POINTAPI 'POINT
    Dim x As Long, y As Long
'-----------------------------------------------------------------
    lHDCTemp = CreateCompatibleDC(lHDCDest)     'Create a temporary HDC compatible to the Destination HDC
    SelectObject lHDCTemp, lBmSource             'Select the bitmap
    GetObject lBmSource, Len(udtBitMap), udtBitMap
    
    With udtSize
        .x = udtBitMap.bmWidth                  'Get width of bitmap
        .y = udtBitMap.bmHeight                 'Get height of bitmap
        'Use passed width and height parameters
        If lDestWidth <> 0 Then .x = lDestWidth
        If lDestHeight <> 0 Then .y = lDestHeight
        x = .x
        y = .y
    End With
    
    'Create some DCs to hold temporary data
    lHDCBack = CreateCompatibleDC(lHDCDest)
    lHDCObject = CreateCompatibleDC(lHDCDest)
    lHDCMem = CreateCompatibleDC(lHDCDest)
    lHDCSave = CreateCompatibleDC(lHDCDest)
    
    'Create a bitmap for each DC.  DCs are required for
    'a number of GDI functions
    
    'Monochrome DC
    lBmAndBack = CreateBitmap(x, y, 1&, 1&, 0&)
    'Monochrome DC
    lBmAndObject = CreateBitmap(x, y, 1&, 1&, 0&)
    'Compatible DC's
    lBmAndMem = CreateCompatibleBitmap(lHDCDest, x, y)
    lBmSave = CreateCompatibleBitmap(lHDCDest, x, y)

    'Each DC must select a bitmap object to store pixel data.
    lBmBackOld = SelectObject(lHDCBack, lBmAndBack)
    lBmObjectOld = SelectObject(lHDCObject, lBmAndObject)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmSaveOld = SelectObject(lHDCSave, lBmSave)
    
    'Set proper mapping mode.
    SetMapMode lHDCTemp, GetMapMode(lHDCDest)
    
    'Save the bitmap sent here, because it will be overwritten
    BitBlt lHDCSave, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcCopy
    
    'Set the background color of the source DC to the color
    'contained in the parts of the bitmap that should be transparent
    lColorRef = SetBkColor(lHDCTemp, lMaskColor)
    
    'Create the object mask for the bitmap by performaing a BitBlt
    'from the source bitmap to a monochrome bitmap.
    BitBlt lHDCObject, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcCopy
    
    'Set the background color of the source DC back to the original color
    SetBkColor lHDCTemp, lColorRef
    
    'Create the inverse of the object mask.
    BitBlt lHDCBack, 0&, 0&, x, y, lHDCObject, 0&, 0&, vbNotSrcCopy
    
    'Copy the background of the main DC to the destination
    If (BkGrndHdc <> 0) Then
        BitBlt lHDCMem, 0&, 0&, x, y, BkGrndHdc, lDestStartX, lDestStartY, vbSrcCopy
    Else
        BitBlt lHDCMem, 0&, 0&, x, y, lHDCDest, lDestStartX, lDestStartY, vbSrcCopy
    End If
    
    'Mask out the places where the bitmap will be placed
    BitBlt lHDCMem, 0&, 0&, x, y, lHDCObject, 0&, 0&, vbSrcAnd
    
    'Mask out the transparent colored pixels on the bitmap
    BitBlt lHDCTemp, lSrcStartX, lSrcStartY, x, y, lHDCBack, 0&, 0&, vbSrcAnd
    
    'XOR the bitmap with the background on the destination DC
    BitBlt lHDCMem, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcPaint
    
    'Copy the destination to the screen
    BitBlt lHDCDest, lDestStartX, lDestStartY, x, y, lHDCMem, 0&, 0&, vbSrcCopy
    
    'Place the original bitmap back into the bitmap sent here
    BitBlt lHDCTemp, lSrcStartX, lSrcStartY, x, y, lHDCSave, 0&, 0&, vbSrcCopy
    
    'Delete memory bitmaps
    DeleteObject SelectObject(lHDCBack, lBmBackOld)
    DeleteObject SelectObject(lHDCObject, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    DeleteObject SelectObject(lHDCSave, lBmSaveOld)
    
    'Delete memory DC's
    DeleteDC lHDCMem
    DeleteDC lHDCBack
    DeleteDC lHDCObject
    DeleteDC lHDCSave
    DeleteDC lHDCTemp
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------
