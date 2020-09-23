Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type udtPoint
    X As Long
    Y As Long
    Speed As Long
    Angle As Single
    Process As Boolean
End Type
Public Type udtExplodePixel
    X As Long
    Y As Long
    Angle As Single
    Speed As Long
End Type
Public Type udtExplosion
    Star(50) As udtExplodePixel
    Color As Long
    xPos As Long
    yPos As Long
    Free As Boolean
End Type
Public Const Pi = 3.1415926535898
Public iNumShips As Integer
Public bAlphaTrails As Boolean
Function GimmeX(ByRef aIn As Single, ByRef lIn As Long) As Long
    GimmeX = Sin(aIn * (Pi / 180)) * lIn
End Function
Function GimmeY(ByRef aIn As Single, ByRef lIn As Long) As Long
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function
Public Function AdjustBrightness(ByRef RGB_In As Long, ByRef ShiftPercentage As Integer, Optional GotoExtreme As Boolean = False) As Long
Dim lColor As Long
Dim r As Single, G As Single, B As Single

    lColor = RGB_In
    r = lColor Mod &H100
    lColor = lColor \ &H100
    G = lColor Mod &H100
    lColor = lColor \ &H100
    B = lColor Mod &H100

    If r > 0 Then r = r + ((r / 100) * ShiftPercentage)
    If G > 0 Then G = G + ((G / 100) * ShiftPercentage)
    If B > 0 Then B = B + ((B / 100) * ShiftPercentage)
    
    If r > 255 Or G > 255 Or B > 255 Then
        If GotoExtreme Then
            If r > 255 Then r = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            AdjustBrightness = RGB(r, G, B)
        Else
            AdjustBrightness = RGB_In
        End If
    ElseIf r < 10 Or G < 10 Or B < 10 Then
        If GotoExtreme Then
            If r < 10 Then r = 0
            If G < 10 Then G = 0
            If B < 10 Then B = 0
            AdjustBrightness = RGB(r, G, B)
        Else
            AdjustBrightness = RGB_In
        End If
    Else
        AdjustBrightness = RGB(r, G, B)
    End If
End Function
