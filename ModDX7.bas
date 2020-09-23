Attribute VB_Name = "ModDX7"
'ModDX7 - by Simon Price
'a module of simple funtions to make DX_Main 7 easier to program
Dim InExMode As Boolean

Sub SetDisplayMode(Width As Integer, Height As Integer, Colors As Byte)
'set's the display mode to the required size and colors
 DD_Main.SetDisplayMode Width, Height, Colors, 0, DDSDM_DEFAULT
End Sub

Sub RestoreDisplayMode()
'puts the screen back to normal
 DD_Main.RestoreDisplayMode
End Sub

Sub CreateSurfaceFromFile(Surface As DirectDrawSurface4, Surfdesc As DDSURFACEDESC2, Filename As String, Width As Integer, Height As Integer)
'loads a bitmap from a file and makes a pic from it
     Surfdesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
     Surfdesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
     Surfdesc.lWidth = Width
     Surfdesc.lHeight = Height
     Set Surface = DD_Main.CreateSurfaceFromFile(Filename, Surfdesc)
End Sub

Sub StretchBlt(Pic As DirectDrawSurface4, X As Integer, y As Integer, Width As Integer, Height As Integer, DestPic As DirectDrawSurface4, DestX As Integer, DestY As Integer, DestWidth As Integer, DestHeight As Integer, Optional AddColorKeyTag As Boolean)
Dim Box As RECT
Box.Left = X
Box.Top = y
Box.Right = X + Width
Box.Bottom = y + Height

Dim DestBox As RECT
DestBox.Left = DestX
DestBox.Top = DestY
DestBox.Right = DestX + DestWidth
DestBox.Bottom = DestY + DestHeight
Pic.Blt DestBox, DestPic, Box, DDBLT_WAIT
End Sub

Sub BitBlt(Pic As DirectDrawSurface4, X As Integer, y As Integer, Width As Integer, Height As Integer, DestPic As DirectDrawSurface7, DestX As Integer, DestY As Integer)
Dim DestBox As RECT
DestBox.Left = DestX
DestBox.Top = DestY
DestBox.Right = DestX + DestWidth
DestBox.Bottom = DestY + DestHeight

Pic.BltFast X, y, DestPic, DestBox, DDBLTFAST_WAIT
End Sub

Function ExModeActive() As Boolean
     Dim TestCoopRes As Long 'holds the return value of the test.

     TestCoopRes = DD_Main.TestCooperativeLevel 'Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True 'everything is fine
     Else
         ExModeActive = False 'this computer doesn't support this mode
     End If
 End Function
 
Sub EndIt(hwnd As Long)
DD_Main.SetCooperativeLevel hwnd, DDSCL_NORMAL
InExMode = False
End Sub

Sub AddColorKey(Surface As DirectDrawSurface4, ColorKey As DDCOLORKEY, low As Long, high As Long)
ColorKey.low = 0
ColorKey.high = 0
Surface.SetColorKey DDCKEY_SRCBLT, ColorKey
End Sub
