VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1740
      Top             =   1380
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   4740
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   1
      Top             =   1020
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AnimObj() As udtPoint
Dim Explosion() As udtExplosion
Private Ship(0 To 4) As POINTAPI
Const ShipRadius As Long = 8
Const AllRunOverTime = 3
Dim RunOverTimer As Long

Private Sub Form_Click()
    Timer1.Enabled = False
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------
    Dim Idx As Integer                          ' Array index
'-----------------------------------------------------------------
    ' [* YOU MUST TURN OFF THE TIMER BEFORE DESTROYING THE SPRITE OBJECT *]
    Timer1.Enabled = False                     ' [* YOU MAY DEADLOCK!!! *]
'   Set gSpriteCollection = Nothing             ' Not sure if this would work...

    DelDeskDC DeskDC                            ' Cleanup the DeskDC (Memleak will occure if not done)
    
    If (RunMode = RM_NORMAL) Then ShowCursor -1 ' Show MousePointer
    Screen.MousePointer = vbDefault             ' Reset MousePointer
    End
'-----------------------------------------------------------------
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Static X0 As Integer, Y0 As Integer
'-----------------------------------------------------------------
    If (RunMode = RM_NORMAL) Then           ' Determine screen saver mode
        If ((X0 = 0) And (Y0 = 0)) Or _
           ((Abs(X0 - X) < 5) And (Abs(Y0 - Y) < 5)) Then ' small mouse movement...
            X0 = X                          ' Save current x coordinate
            Y0 = Y                          ' Save current y coordinate
            Exit Sub                        ' Exit
        End If
    
        Unload Me
        End ' Large mouse movement (terminate screensaver)
    End If
End Sub
Private Sub Form_Load()

    Randomize Timer
    'Me.Show
    
    If (RunMode = RM_NORMAL) Then ShowCursor 0
    
    InitDeskDC DeskDC, DeskBmp, DispRec
    
    Me.Move 0, 0, Screen.Width, Screen.Height
    picBuffer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picBlank.Move 0, 0, picBuffer.Width, picBuffer.Height
    
    SetupAnimObj 'define ships
    
    'set 1 explosion and set it free for use
    ReDim Explosion(0)
    Explosion(0).Free = True
    
    'used when all the ships have been 'run over'
    RunOverTimer = 0
    
    Timer1.Enabled = True
End Sub
Sub SetupAnimObj()
Dim lX As Long
    ReDim AnimObj(iNumShips)
    For lX = 0 To UBound(AnimObj)
        With AnimObj(lX)
            .X = Int(Rnd * (picBuffer.Width - (2 * ShipRadius))) + ShipRadius
            .Y = Int(Rnd * (picBuffer.Height - (2 * ShipRadius))) + ShipRadius
            .Speed = Int(Rnd * 30) + 1
            .Angle = Rnd * 360
            .Process = True
        End With
    Next
End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub
Sub BufferToScreen()
    BitBlt Me.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Timer1_Timer()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    If Not bAlphaTrails Then
        ClearBuffer
    Else
        Blend.SourceConstantAlpha = 100
        
        CopyMemory BlendPtr, Blend, 4
        
        AlphaBlend picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, picBuffer.Width, picBuffer.Height, BlendPtr
    End If
    CreateFrame
    BufferToScreen
End Sub

Sub CreateFrame()
Dim lX As Long
Dim lY As Long
Dim hrPen As Long
Dim xDistance As Long
Dim yDistance As Long
Dim hDistance As Long

Dim EveryoneGotRunover As Boolean

    For lX = 0 To UBound(AnimObj)
        With AnimObj(lX)
            If .Process Then 'ship is alive
                DefineShip .Angle, .X, .Y
                hrPen = CreatePen(0, 1, vbWhite) '0=solid brush
                SelectObject picBuffer.hdc, hrPen
                Polyline picBuffer.hdc, Ship(0), 5 'draw the ship
                DeleteObject hrPen
            
                'move the ship along its course
                .X = .X + GimmeX(.Angle, .Speed)
                .Y = .Y + GimmeY(.Angle, .Speed)
                'hit the boundaries? deflect the ship's angle
                If .X < ShipRadius Or .X > (picBuffer.Width - ShipRadius) Then 'off x axis
                    If .X < ShipRadius Then
                        If .Angle < 360 And .Angle > 270 Then
                            .Angle = 90 - (.Angle - 270)
                        Else
                            .Angle = 90 + (270 - .Angle)
                        End If
                    Else
                        If .Angle >= 0 And .Angle < 90 Then
                            .Angle = 270 + (90 - .Angle)
                        Else
                            .Angle = 270 - (.Angle - 90)
                        End If
                    End If
                End If
                If .Y < ShipRadius Or .Y > (picBuffer.Height - ShipRadius) Then 'off y axis
                    If .Y < ShipRadius Then
                        If .Angle > 0 And .Angle < 90 Then
                            .Angle = 90 - .Angle
                        Else
                            .Angle = 270 - (.Angle - 270)
                        End If
                    Else
                        If .Angle >= 90 And .Angle < 180 Then
                            .Angle = 180 - .Angle
                        Else
                            .Angle = 360 - (.Angle - 180)
                        End If
                    End If
                End If
                'did anyone get run over?
                EveryoneGotRunover = True
                For lY = 0 To UBound(AnimObj)
                    If lY <> lX Then
                        If AnimObj(lY).Process Then
                            EveryoneGotRunover = False
                            xDistance = Abs(AnimObj(lY).X - AnimObj(lX).X)
                            yDistance = Abs(AnimObj(lY).Y - AnimObj(lX).Y)
                            'bit of pythagoras to get their distance apart...
                            hDistance = (((xDistance * xDistance) + (yDistance * yDistance)) ^ 0.5) \ 2
                            If hDistance <= ShipRadius Then
                                AnimObj(lY).Process = False
                                'start an explosion where the run over ship was
                                SetupExplosion .X, .Y
                            End If
                        End If
                    End If
                Next
                If EveryoneGotRunover Then
                    If RunOverTimer = 0 Then
                        RunOverTimer = Timer
                    ElseIf Timer - RunOverTimer > AllRunOverTime Then
                        For lY = 0 To UBound(AnimObj)
                            AnimObj(lY).Process = True
                        Next
                        RunOverTimer = 0
                        ReDim Explosion(0)
                        Explosion(0).Free = True
                    End If
                End If
            End If
        End With
        
    Next
    'Draw the explosions
    For lX = 0 To UBound(Explosion)
        If Explosion(lX).Free = False Then 'explosion in progress
            With Explosion(lX)
                For lY = 0 To UBound(.Star)
                    With .Star(lY)
                        SetPixel picBuffer.hdc, .X, .Y, Explosion(lX).Color
                        .X = .X + GimmeX(.Angle, .Speed)
                        .Y = .Y + GimmeY(.Angle, .Speed)
                    End With
                Next
                .Color = AdjustBrightness(.Color, -5, True)
                If .Color = vbBlack Then
                    .Free = True
                End If
            End With
        End If
    Next lX
End Sub
Sub SetupExplosion(xIN As Long, yIN As Long)
Dim iX As Integer
Dim iStar As Integer
Dim FreeSlot As Boolean
    'is there a free slot?
    FreeSlot = False
    For iX = 0 To UBound(Explosion)
        If Explosion(iX).Free Then
            FreeSlot = True
            Exit For
        End If
    Next
    
    If FreeSlot Then
        With Explosion(iX)
            .Free = False
            .xPos = xIN
            .yPos = yIN
            .Color = vbWhite
            For iStar = 0 To UBound(.Star)
                With .Star(iStar)
                    .X = xIN ' + (Int(Rnd * 10) \ 2)
                    .Y = yIN ' + (Int(Rnd * 10) \ 2)
                    .Angle = Int(Rnd * 360)
                    .Speed = Int(Rnd * 5) + 5
                End With
            Next
        End With
    Else 'instantiate new explosion
        ReDim Preserve Explosion(UBound(Explosion) + 1)
        Explosion(UBound(Explosion)).Free = True
        SetupExplosion xIN, yIN
    End If
End Sub
Sub DefineShip(ByVal sAngle As Single, X As Long, Y As Long)
    Ship(0).X = GimmeX(sAngle, 12) + X
    Ship(0).Y = GimmeY(sAngle, 12) + Y
    
    sAngle = sAngle + 135
    If sAngle > 359 Then sAngle = sAngle - 360
    Ship(1).X = GimmeX(sAngle, 9) + X
    Ship(1).Y = GimmeY(sAngle, 9) + Y
    
    sAngle = sAngle + 45
    If sAngle > 359 Then sAngle = sAngle - 360
    Ship(2).X = GimmeX(sAngle, 2) + X
    Ship(2).Y = GimmeY(sAngle, 2) + Y
    
    sAngle = sAngle + 45
    If sAngle > 359 Then sAngle = sAngle - 360
    Ship(3).X = GimmeX(sAngle, 9) + X
    Ship(3).Y = GimmeY(sAngle, 9) + Y
    
    Ship(4).X = Ship(0).X
    Ship(4).Y = Ship(0).Y
    

End Sub

