Attribute VB_Name = "Game"
Public Sub InitSurfaces()
ModDX7.CreateSurfaceFromFile DS_BG, SD_BG, Path$ & "\redgiant.bmp", 700, 500
RT_BG.Bottom = 500
RT_BG.Right = 700

ModDX7.CreateSurfaceFromFile DS_Ship, SD_Ship, Path$ & "\ship.bmp", 236, 67
RT_Ship.Bottom = 67
RT_Ship.Right = 59
ModDX7.AddColorKey DS_Ship, CK_Ship, 0, 0

ModDX7.CreateSurfaceFromFile DS_L, SD_L, Path$ & "\laser2.bmp", 4, 10
RT_L.Bottom = 10
RT_L.Right = 4
ModDX7.AddColorKey DS_L, CK_L, 0, 0

ModDX7.CreateSurfaceFromFile DS_E, SD_E, Path$ & "\e1.bmp", 118, 59
RT_E.Bottom = 59
RT_E.Right = 59
ModDX7.AddColorKey DS_E, CK_E, 0, 0

ModDX7.CreateSurfaceFromFile DS_EL, SD_EL, Path$ & "\elaser.bmp", 4, 10
RT_EL.Bottom = 10
RT_EL.Right = 4
ModDX7.AddColorKey DS_EL, CK_EL, 0, 0

ModDX7.CreateSurfaceFromFile DS_P, SD_P, Path$ & "\power.bmp", 60, 30
RT_P.Bottom = 30
RT_P.Right = 20

DS_Back.SetFont FrmMain.Font
End Sub
Public Sub BG()
DS_Back.BltFast 50, 50, DS_BG, RT_BG, DDBLTFAST_WAIT
End Sub
Public Sub Wait(Length As Double)
Dim StartTime
StartTime = Timer
Do While Timer - StartTime < Length
    DoEvents
Loop
End Sub
Public Function GetKey(lngKey As Long) As Boolean
If GetKeyState(lngKey&) < 0 Then
    GetKey = True
Else
    GetKey = False
End If
End Function
Public Sub Ship()
If Health > 25 Then
    DS_Back.BltFast ShipX, ShipY, DS_Ship, RT_Ship, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
If Health < 25 Then
    RT_Ship.Left = 118
    RT_Ship.Right = 177
    DS_Back.BltFast ShipX, ShipY, DS_Ship, RT_Ship, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
For i = 0 To 20
    If Laser(i) = True Then
        LasersY(i) = LasersY(i) - 6
        DS_Back.BltFast LasersX(i), LasersY(i), DS_L, RT_L, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    If LasersY(i) < 50 Then
        Laser(i) = False
        LasersX(i) = -300
        LasersY(i) = -300
    End If
Next i

If PGo = True Then
    DS_Back.BltFast PX, PY, DS_P, RT_P, DDBLTFAST_WAIT
    PY = PY + 2
    If PX > ShipX And PX < ShipX + 59 Then
        If PY > ShipY And PY < ShipY + 67 Then
            If RndP = 0 Then
                Health = Health + 30
                If Health > 60 Then
                    Health = 60
                End If
            ElseIf RndP = 1 Then
                Energy = Energy + 30
                If Energy > 60 Then
                    Energy = 60
                End If
            ElseIf RndP = 2 Then
                Inv = True
                FrmMain.TmrInvincible.Enabled = True
            End If
            PGo = False
        End If
    End If
End If
End Sub
Public Sub Enemies()
If Ticks = 15 Then
    SendEnemy 330, 0
ElseIf Ticks = 25 Then
    SendEnemy 200, 1
    SendEnemy 500, 2
End If
Select Case Ticks
Case 40
    SendEnemy 100, 3
Case 45
    SendEnemy 200, 4
Case 50
    SendEnemy 300, 5
Case 65
    SendEnemy 600, 6
Case 70
    SendEnemy 500, 7
Case 75
    SendEnemy 400, 8
Case 100
    SendEnemy 100, 0
Case 110
    SendEnemy 200, 1
Case 120
    SendEnemy 300, 2
Case 130
    SendEnemy 400, 3
Case 140
    SendEnemy 500, 4
Case 150
    SendEnemy 600, 5
Case 180
    SendEnemy 330, 0
Case 188
    SendEnemy 330, 1
Case 196
    SendEnemy 330, 2
Case 204
    SendEnemy 330, 3
Case 212
    SendEnemy 330, 4
Case 220
    SendEnemy 330, 5
Case 250
    SendEnemy 100, 0, True
    SendEnemy 600, 1, True
Case 280
    SendEnemy 100, 2, True
    SendEnemy 600, 3, True
Case 290
    SendEnemy 200, 4, True
    SendEnemy 500, 5, True
Case 300
    SendEnemy 300, 6, True
    SendEnemy 400, 7, True
Case 330
    SendEnemy 330, 0
Case 335
    SendEnemy 100, 1
Case 340
    SendEnemy 600, 2
Case 350
    SendEnemy 330, 3, True
Case 355
    SendEnemy 330, 4, True
Case 360
    SendEnemy 330, 5, True
Case 370
    SendEnemy 100, 0
Case 380
    SendEnemy 200, 1
Case 390
    SendEnemy 300, 2, True
Case 400
    SendEnemy 400, 3, True
Case 430
    SendEnemy 100, 0
    SendEnemy 200, 1
    SendEnemy 300, 2, True
    SendEnemy 400, 3
    SendEnemy 500, 4
Case 440
    SendEnemy 330, 5
Case 445
    SendEnemy 230, 6
Case 450
    SendEnemy 430, 0
Case 460
    SendEnemy 330, 1, True
Case 470
    SendEnemy 130, 2
Case 480
    SendEnemy 230, 3
Case 495
    SendEnemy 570, 4
Case 500
    SendEnemy 330, 5, True
Case 520
    SendEnemy 100, 0
Case 530
    SendEnemy 200, 1
Case 540
    SendEnemy 300, 2, True
Case 550
    SendEnemy 400, 3
Case 560
    SendEnemy 500, 4
Case 570
    SendEnemy 600, 5
Case 600
    SendEnemy 330, 0
    SendEnemy 330, 1
    SendEnemy 330, 2
Case 610
    SendEnemy 230, 3
    SendEnemy 230, 4
Case 620
    SendEnemy 430, 5
    SendEnemy 430, 6
Case 640
    SendEnemy 230, 7, True
Case 650
    SendEnemy 400, 8, True
Case 660
    SendEnemy 130, 0, False
    SendEnemy 630, 1, False
Case 670
    SendEnemy 230, 2, True
    SendEnemy 530, 3, True
Case 680
    SendEnemy 330, 4, True
    SendEnemy 430, 5, True
Case 700
    SendEnemy 130, 0
    SendEnemy 130, 1
Case 705
    SendEnemy 230, 2
    SendEnemy 230, 3
Case 710
    SendEnemy 330, 4
    SendEnemy 330, 5
Case 715
    SendEnemy 430, 6
    SendEnemy 430, 7
Case 720
    SendEnemy 530, 8
    SendEnemy 630, 9
Case 745
    Ticks = 0
End Select
For i = 0 To 10
    If EnemiesOn(i) = True Then
        EnemyY(i) = EnemyY(i) + 2
        If SE(i) = True Then
            If EDir(i) = 1 Then
                EnemyX(i) = EnemyX(i) - 2
            ElseIf EDir(i) = 2 Then
                EnemyX(i) = EnemyX(i) + 2
            End If
        End If
        For l = 0 To 20
            If Laser(l) = True Then
                If LasersX(l) > EnemyX(i) And LasersX(l) < EnemyX(i) + 59 Then
                    If LasersY(l) < EnemyY(i) + 59 And LasersY(l) > EnemyY(i) Then
                        EnemyHit(i) = EnemyHit(i) + 1
                        LasersX(l) = -200
                        LasersY(l) = -200
                        Laser(l) = False
                    End If
                End If
            End If
        Next l
        If EnemyHit(i) = 0 Then
            RT_E.Left = 0
            RT_E.Right = 59
        ElseIf EnemyHit(i) = 1 Then
            RT_E.Left = 59
            RT_E.Right = 118
        ElseIf EnemyHit(i) = 2 Then
            EnemiesOn(i) = False
            EnemyX(i) = -300
            EnemyY(i) = -300
            Score = Score + 50
            EnemyHit(i) = 3
            SE(i) = False
        End If
        DS_Back.BltFast EnemyX(i), EnemyY(i), DS_E, RT_E, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    If EnemyY(i) > 600 Then
        EnemiesOn(i) = False
    End If
Next i
For i = 0 To 10
    If EnemyShot(i) = True Then
        DS_Back.BltFast ESX(i), ESY(i), DS_EL, RT_EL, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ESY(i) = ESY(i) + 5
    End If
    If EnemyShot(i) = True Then
        If ESX(i) > ShipX And ESX(i) < ShipX + 59 Then
            If ESY(i) > ShipY And ESY(i) < ShipY + 67 Then
                TakeHealth 10
                EnemyShot(i) = False
            End If
        End If
    End If
Next i
End Sub
Public Sub SendEnemy(X As Double, Enemy As Integer, Optional Swurve As Boolean)
EnemiesOn(Enemy) = True
EnemyX(Enemy) = X
EnemyY(Enemy) = 0
EnemyHit(Enemy) = 0
EnemyShot(Enemy) = False
If Swurve = True Then
    SE(Enemy) = True
End If
End Sub
Public Sub TakeEnergy()
If Energy > 0 Then
    Energy = Energy - 0.2
Else
    Health = Health - 1
    If Health < 0 Then
        ResetShip
    End If
End If
End Sub
Public Sub TakeHealth(H As Double)
If Inv = False Then
If Health - H > -1 Then
    Health = Health - H
ElseIf Health - H < 0 Then
    Health = 0
    ResetShip
End If
End If
End Sub
Public Sub ResetShip()
If Lives > 0 Then
    ShipX = -500
    RefreshScreen
    Wait 1.5
ElseIf Lives = 0 Then
    Lives = 5
    RefreshScreen2
    Wait 4#
End If
ShipX = 330
ShipY = 450
Health = 60
Energy = 60
Lives = Lives - 1
End Sub
Public Sub RefreshScreen()
D3D_Viewport.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER
D3D_Device.Update
D3D_Viewport.Render FR_Root
BG
DS_Back.SetForeColor vbWhite
DS_Back.DrawText 340, 290, "YOU SUCK!", False
DS_Front.Flip Nothing, DDFLIP_WAIT
End Sub
Public Sub SendStrong(X As Double, Enemy As Integer, HowStrong As Integer)
For i = 0 To HowStrong
    SendEnemy X, Enemy + i, False
Next i
End Sub
Public Sub RefreshScreen2()
D3D_Viewport.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER
D3D_Device.Update
D3D_Viewport.Render FR_Root
BG
DS_Back.SetForeColor vbWhite
DS_Back.DrawText 240, 290, "GAME OVER, MAN, GAME OVER, YEAH!", False
DS_Front.Flip Nothing, DDFLIP_WAIT
Score = 0
End Sub
