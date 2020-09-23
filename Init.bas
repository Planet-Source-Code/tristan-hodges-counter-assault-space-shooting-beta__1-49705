Attribute VB_Name = "Init"
Public Sub DX_Init()
Set DD_Main = Nothing
Set DD_Main = DX_Main.DirectDraw4Create("")
DD_Main.SetCooperativeLevel FrmMain.hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
DD_Main.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT

SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
SD_Front.lBackBufferCount = 1
Set DS_Front = Nothing
Set DS_Front = DD_Main.CreateSurface(SD_Front)

DD_Back.lCaps = DDSCAPS_BACKBUFFER
Set DS_Back = Nothing
Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)

Set D3D_Main = Nothing
Set D3D_Device = Nothing
Set D3D_Main = DX_Main.Direct3DRMCreate()
Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD_Main, DS_Back, D3DRMDEVICE_DEFAULT)
D3D_Device.SetBufferCount 2
D3D_Device.SetQuality D3DRMRENDER_GOURAUD
D3D_Device.SetTextureQuality D3DRMTEXTURE_LINEAR
D3D_Device.SetRenderMode D3DRMRENDERMODE_SORTEDTRANSPARENCY

Set FR_Root = Nothing
Set FR_Root = D3D_Main.CreateFrame(Nothing)
Set FR_Camera = Nothing
Set FR_Camera = D3D_Main.CreateFrame(FR_Root)

FR_Root.SetSceneBackgroundRGB 0, 0, 1

FR_Camera.SetPosition Nothing, 1, 6, -10
Set D3D_Viewport = Nothing
Set D3D_Viewport = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 800, 600)
D3D_Viewport.SetBack 300

Set DS = Nothing
Set DS = DX_Main.DirectSoundCreate("")
DS.SetCooperativeLevel FrmMain.hwnd, DSSCL_PRIORITY

InitGraphics
End Sub
Public Sub InitGraphics()

End Sub
