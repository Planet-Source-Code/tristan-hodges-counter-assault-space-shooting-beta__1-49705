Attribute VB_Name = "Variables"
Public DX_Main As New DirectX7
Public DD_Main As DirectDraw4
Public D3D_Main As Direct3DRM3

Public DI_Main As DirectInput
Public DI_Device As DirectInputDevice
Public DI_State As DIKEYBOARDSTATE

Public DS_Front As DirectDrawSurface4
Public DS_Back As DirectDrawSurface4
Public SD_Front As DDSURFACEDESC2
Public DD_Back As DDSCAPS2

Public DS As DirectSound
Public DS_Buffer As DirectSoundBuffer
Public DS_BGMusic As DirectSoundBuffer

Public D3D_Device As Direct3DRMDevice3
Public D3D_Viewport As Direct3DRMViewport2

Public FR_Root As Direct3DRMFrame3
Public FR_Camera As Direct3DRMFrame3

Public TX_BG As Direct3DRMTexture3

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Path As String
Public Frames As Integer
Public LastFrames As Integer

Public ShipX As Double
Public ShipY As Double


Public DS_BG As DirectDrawSurface4
Public SD_BG As DDSURFACEDESC2
Public CK_BG As DDCOLORKEY
Public RT_BG As RECT

Public DS_P As DirectDrawSurface4
Public SD_P As DDSURFACEDESC2
Public RT_P As RECT

Public DS_Ship As DirectDrawSurface4
Public SD_Ship As DDSURFACEDESC2
Public CK_Ship As DDCOLORKEY
Public RT_Ship As RECT

Public DS_L As DirectDrawSurface4
Public SD_L As DDSURFACEDESC2
Public CK_L As DDCOLORKEY
Public RT_L As RECT

Public DS_EL As DirectDrawSurface4
Public SD_EL As DDSURFACEDESC2
Public CK_EL As DDCOLORKEY
Public RT_EL As RECT

Public DS_E As DirectDrawSurface4
Public SD_E As DDSURFACEDESC2
Public CK_E As DDCOLORKEY
Public RT_E As RECT

Public LasersX(20) As Double
Public LasersY(20) As Double
Public Laser(20) As Boolean

Public Ticks As Integer

Public Lives As Integer
Public Score As Integer

Public EnemyX(10) As Double
Public EnemyY(10) As Double
Public EnemiesOn(10) As Boolean
Public EnemyHit(10) As Integer
Public SE(10) As Boolean
Public EDir(10) As Integer

Public EnemyShot(10) As Boolean
Public ESX(10) As Double
Public ESY(10) As Double

Public Health As Double
Public Energy As Double

Public PX As Double
Public PY As Double
Public PGo As Boolean
Public RndP As Integer
Public Inv As Boolean

Public TicksPlus As Integer
Public ECount As Integer
