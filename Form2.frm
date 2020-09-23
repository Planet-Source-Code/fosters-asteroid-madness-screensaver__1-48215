VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asteroid Madness"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAlpha 
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   660
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   900
      Width           =   735
   End
   Begin VB.TextBox txtNumShips 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "100"
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "Alpha blended trails"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Ships"
      Height          =   195
      Left            =   495
      TabIndex        =   2
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "by Mike Toye, mike@qucami.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "by Mike Toye, mike@qucami.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1695
      TabIndex        =   1
      Top             =   1575
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   1500
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    SaveSetting App.Title, "Settings", "Ships", txtNumShips
    SaveSetting App.Title, "Settings", "Alpha", chkAlpha.Value
    Unload Me
End Sub

Private Sub Form_Load()
    txtNumShips = GetSetting(App.Title, "Settings", "Ships")
    chkAlpha.Value = CInt(GetSetting(App.Title, "Settings", "Alpha"))
End Sub

Private Sub txtNumShips_LostFocus()
    If Len(txtNumShips) = 0 Then
        MsgBox "Please specify the number of ships to display", vbInformation, App.Title
        txtNumShips.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtNumShips) Then
        MsgBox "Please specify a numeric!", vbInformation, App.Title
        txtNumShips.SelStart = 0
        txtNumShips.SelLength = Len(txtNumShips)
        txtNumShips.SetFocus
        Exit Sub
    End If
    If CInt(txtNumShips) < 2 Then
        MsgBox "Please specify more than 1 ship", vbInformation, App.Title
        txtNumShips.SelStart = 0
        txtNumShips.SelLength = Len(txtNumShips)
        txtNumShips.SetFocus
        Exit Sub
    End If
    

End Sub
