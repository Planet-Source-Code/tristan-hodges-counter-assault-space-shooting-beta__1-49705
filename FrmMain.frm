VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrEDir 
      Interval        =   500
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer TmrInvincible 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Interval        =   32000
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2040
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer TmrShoot 
      Interval        =   200
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer TmrKeys 
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer TmrRefresh 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Path$ = App.Path & "\data"
Init.DX_Init
InitSurfaces
ShipX = 350
ShipY = 450
Health = 60
Energy = 60
Lives = 4
Randomize
End Sub

Private Sub Timer1_Timer()
Ticks = Ticks + 1
End Sub

Private Sub Timer2_Timer()
Dim RandomShip As Integer
For i = 0 To 10
    If EnemiesOn(i) = True And EnemyShot(i) = False Then
        RandomShip = Int(Rnd * 6)
        If RandomShip = 3 Then
            EnemyShot(i) = True
            ESX(i) = EnemyX(i) + 28
            ESY(i) = EnemyY(i) + 59
        End If
    End If
Next i
End Sub

Private Sub Timer3_Timer()
Randomize
RndP = Int(Rnd * 3)
If RndP = 0 Then
    RT_P.Left = 0
    RT_P.Right = 20
ElseIf RndP = 1 Then
    RT_P.Left = 20
    RT_P.Right = 40
ElseIf RndP = 2 Then
    RT_P.Left = 40
    RT_P.Right = 60
End If
PGo = True
PX = Int(Rnd * 450) + 100
PY = 0
End Sub

Private Sub TmrEDir_Timer()
For i = 0 To 10
    EDir(i) = Int(Rnd * 3)
Next i
End Sub

Private Sub TmrInvincible_Timer()
Inv = False
TmrInvincible.Enabled = False
End Sub

Private Sub TmrKeys_Timer()
If GetKey(vbKeyEscape) Then
    End
End If
If GetKey(vbKeyRight) Then
    ShipX = ShipX + 3
End If
If GetKey(vbKeyLeft) Then
    ShipX = ShipX - 3
End If
If GetKey(vbKeyUp) Then
    ShipY = ShipY - 3
    If Health > 25 Then
        RT_Ship.Right = 118
        RT_Ship.Left = 59
    End If
End If
If GetKey(vbKeyDown) Then
    ShipY = ShipY + 3
End If
If GetKey(vbKeyUp) = False Then
    RT_Ship.Left = 0
    RT_Ship.Right = 59
End If
If ShipX < 50 Then
    ShipX = 50
End If
If ShipX > 691 Then
    ShipX = 691
End If
If ShipY > 490 Then
    ShipY = 490
End If
If ShipY < 50 Then
    ShipY = 50
End If
End Sub

Private Sub TmrRefresh_Timer()
On Error Resume Next
D3D_Viewport.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER
D3D_Device.Update
D3D_Viewport.Render FR_Root
BG
Ship
Enemies
DS_Back.setDrawWidth 50
DS_Back.DrawLine 0, 25, 800, 25
DS_Back.DrawLine 0, 575, 800, 575
DS_Back.DrawLine 25, 0, 25, 600
DS_Back.DrawLine 775, 0, 775, 600

DS_Back.SetForeColor vbWhite
DS_Back.DrawText 10, 10, "Lives: " & Lives, False
DS_Back.DrawText 140, 10, "Score: " & Score, False
DS_Back.DrawText 410, 13, "Health:", False
DS_Back.DrawText 410, 36, "Energy:", False
DS_Back.SetForeColor vbBlack

DS_Back.SetForeColor &H808080
DS_Back.setDrawWidth 20
DS_Back.DrawLine 500, 25, 680, 25

DS_Back.SetForeColor vbRed
DS_Back.setDrawWidth 15
DS_Back.DrawLine 500, 25, (500 + (Health * 3)), 25

DS_Back.SetForeColor &H808080
DS_Back.setDrawWidth 20
DS_Back.DrawLine 500, 50, 680, 50

DS_Back.SetForeColor vbGreen
DS_Back.setDrawWidth 15
DS_Back.DrawLine 500, 50, (500 + (Energy * 3)), 50
DS_Back.SetForeColor vbBlack
DS_Front.Flip Nothing, DDFLIP_WAIT
End Sub

Private Sub TmrShoot_Timer()
If GetKey(vbKeySpace) Then
    For i = 0 To 20
        If Laser(i) = False Then
            TakeEnergy
            Laser(i) = True
            LasersX(i) = ShipX + 30
            LasersY(i) = ShipY - 10
            Exit For
        End If
    Next i
End If
End Sub
