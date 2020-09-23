VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General Declarations
Dim xxxx As Integer
Const Sin5 = 0.1071557!     ' Sin(5°)
Const Cos5 = 1.096195!  ' Cos(5°)
Dim Pos As D3DVECTOR
Dim Go As Boolean
Dim speed As Variant
Dim Mov As D3DVECTOR


'Textures
Dim texStreet As Direct3DRMTexture3
Dim texBetterBricks As Direct3DRMTexture3
Dim texWindows1 As Direct3DRMTexture3
Dim texWindows2 As Direct3DRMTexture3
Dim texWindowMix As Direct3DRMTexture3
Dim texStands  As Direct3DRMTexture3

'These three components are the big daddy of what your 3D Program.
Dim DX_Main As New DirectX7 ' The DirectX core file (The heart of it all)
Dim DD_Main As DirectDraw4 ' The DirectDraw Object
Dim D3D_Main As Direct3DRM3 ' The direct3D Object

'DirectInput Components
Dim DI_Main As DirectInput ' The DirectInput core
Dim DI_Device As DirectInputDevice ' The DirectInput device
Dim DI_State As DIKEYBOARDSTATE 'Array Holding the state of the keys

'DirectDraw Surfaces- Where the screen is drawn.
Dim DS_Front As DirectDrawSurface4 ' The frontbuffer (What you see on the screen)
Dim DS_Back As DirectDrawSurface4 ' The backbuffer, (Where everything is drawn before it's put on the screen.)
Dim SD_Front As DDSURFACEDESC2 ' The SurfaceDescription
Dim DD_Back As DDSCAPS2 ' General Surface Info

'ViewPort and Direct3D Device
Dim D3D_Device As Direct3DRMDevice3 'The Main Direct3D Retained Mode Device
Dim D3D_ViewPort As Direct3DRMViewport2 'The Direct3D Retained Mode Viewport (Kinda the camera)

'The Frames
Dim FR_Root As Direct3DRMFrame3 'The Main Frame (The other frames are put under this one (Like a tree))
Dim FR_Camera As Direct3DRMFrame3 'Another frame, just happens to be called 'camera'. We will use this
Dim FR_Light As Direct3DRMFrame3 'This frame contains our, guess what, spotlight!
Dim FR_Street As Direct3DRMFrame3 'Frame containing our 1st mesh that will be put in this "game".
Dim FR_WideBuilding1 As Direct3DRMFrame3
Dim FR_midBuilding1 As Direct3DRMFrame3
Dim FR_BigBuilding1 As Direct3DRMFrame3
Dim FR_MidBuilding2 As Direct3DRMFrame3
Dim FR_midBuilding3 As Direct3DRMFrame3
Dim FR_longbuilding1 As Direct3DRMFrame3
Dim FR_longbuilding2 As Direct3DRMFrame3
Dim FR_LongBuilding3 As Direct3DRMFrame3
Dim FR_LongBuilding4 As Direct3DRMFrame3
Dim FR_LOngBuilding5 As Direct3DRMFrame3
Dim FR_WideBuilding2 As Direct3DRMFrame3
Dim FR_WideBuilding3 As Direct3DRMFrame3
Dim FR_MidBuilding4 As Direct3DRMFrame3
Dim FR_WideBuilding4 As Direct3DRMFrame3
Dim FR_midBuilding5 As Direct3DRMFrame3
Dim FR_Wearhouse1 As Direct3DRMFrame3
Dim FR_MidBuilding6 As Direct3DRMFrame3
Dim FR_MidBuilding7 As Direct3DRMFrame3
Dim FR_Stands As Direct3DRMFrame3
Dim FR_Car As Direct3DRMFrame3

'Meshes (3D objects loaded from a .x file)
Dim MS_Street As Direct3DRMMeshBuilder3
Dim MS_WideBuilding1 As Direct3DRMMeshBuilder3
Dim MS_midBuilding1 As Direct3DRMMeshBuilder3
Dim MS_BigBuilding1 As Direct3DRMMeshBuilder3
Dim MS_midBuilding2 As Direct3DRMMeshBuilder3
Dim MS_midBuilding3 As Direct3DRMMeshBuilder3
Dim MS_longbuilding1 As Direct3DRMMeshBuilder3
Dim MS_longbuilding2 As Direct3DRMMeshBuilder3
Dim MS_LongBuilding3 As Direct3DRMMeshBuilder3
Dim MS_LongBuilding4 As Direct3DRMMeshBuilder3
Dim MS_LongBuilding5 As Direct3DRMMeshBuilder3
Dim MS_WideBuilding2 As Direct3DRMMeshBuilder3
Dim MS_WideBuilding3 As Direct3DRMMeshBuilder3
Dim MS_MidBuilding4 As Direct3DRMMeshBuilder3
Dim MS_Wearhouse1 As Direct3DRMMeshBuilder3
Dim MS_Stands As Direct3DRMMeshBuilder3
Dim MS_Car As Direct3DRMMeshBuilder3

'Lights
Dim LT_Ambient As Direct3DRMLight 'Our Main (Ambient) light that illuminates everything (not just part of something
                                                'like a spotlight.
Dim LT_Spot As Direct3DRMLight 'Our Spot light, makes it look more realistic.

'Camera Positions
Dim xx As Long
Dim yy As Long
Dim zz As Long

Dim esc As Boolean 'If Escape is pressed, the DX_Input sub will make it true and the main loop will end.
'Incase you haven't caught on, the prefix FR = frame, MS = Mesh, & LT = light.
'======================================================================================

Private Sub DX_Init()
 'Type, not copy and paste!
 'This sub will initialize all your components and set them up.
 Set DD_Main = DX_Main.DirectDraw4Create("") 'Create the DirectDraw Object

 DD_Main.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE 'Set Screen Mode (Full 'Screen)
 DD_Main.SetDisplayMode 640, 480, 32, 0, DDSDM_DEFAULT 'Set Resolution and BitDepth (Lets use 32-bit color)
 
 SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
 SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or _
 DDSCAPS_FLIP 'I used the line-continuation ( _ ) because the whole thing wouldn't fit on one line on the HTML tutorial...
 SD_Front.lBackBufferCount = 1 'Make one backbuffer
 Set DS_Front = DD_Main.CreateSurface(SD_Front) 'Initialize the front buffer (the screen)
 'The Previous block of code just created the screen and the backbuffer.
 
 DD_Back.lCaps = DDSCAPS_BACKBUFFER
 Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)
 DS_Back.SetForeColor RGB(255, 255, 255)
 'The backbuffer was initialized and the DirectDraw text color was set to white.

 Set D3D_Main = DX_Main.Direct3DRMCreate() 'Creates the Direct3D Retained Mode Object!!!!!

 Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD_Main, DS_Back, _
 D3DRMDEVICE_DEFAULT) 'Tell the Direct3D Device that we are using hardware rendering (HALDevice) instead
                                   'of software enumeration (RGBDevice).
 D3D_Device.SetBufferCount 2 'Set the number of buffers
 D3D_Device.SetQuality D3DRMRENDER_GOURAUD 'Set Rendering Quality. Can use Flat, or WireFrame, but
                                                                  'GOURAUD has the best rendering quality.
 D3D_Device.SetTextureQuality D3DRMTEXTURE_NEAREST 'Set the texture quality
 D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY 'Set the render mode.

 Set DI_Main = DX_Main.DirectInputCreate() 'Create the DirectInput Device
 Set DI_Device = DI_Main.CreateDevice("GUID_SysKeyboard") 'Set it to use the keyboard.
 DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD 'Set the data format to the keyboard format.
 DI_Device.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'Set Coperative Level.
 DI_Device.Acquire
 'The above block of code configures the DirectInput Device and starts it.
 DS_Front.SetForeColor RGB(255, 255, 255) '220 310
 DS_Front.DrawBox 200, 220, 468, 330
 DS_Front.DrawText 200, 220, "Loading Matthew's Racing Game", False
 DS_Front.DrawText 200, 250, "Creating Frames...", False
 
End Sub

'=====================================================================================

Private Sub DX_MakeObjects()
 Set FR_Root = D3D_Main.CreateFrame(Nothing) 'This will be the root frame of the 'tree'
 Set FR_Camera = D3D_Main.CreateFrame(FR_Root) 'Our Camera's Sub Frame. It goes under FR_Root in the 'Tree'.
 Set FR_Light = D3D_Main.CreateFrame(FR_Root) 'Our Light's Sub Frame
 Set FR_Street = D3D_Main.CreateFrame(FR_Root) 'Our Building (the 3D Thingy that will be placed in our world's)
 Set FR_WideBuilding1 = D3D_Main.CreateFrame(FR_Root)
 Set FR_midBuilding1 = D3D_Main.CreateFrame(FR_Root)
 Set FR_BigBuilding1 = D3D_Main.CreateFrame(FR_Root)
 Set FR_MidBuilding2 = D3D_Main.CreateFrame(FR_Root)
 Set FR_midBuilding3 = D3D_Main.CreateFrame(FR_Root)
 Set FR_longbuilding1 = D3D_Main.CreateFrame(FR_Root)
 Set FR_longbuilding2 = D3D_Main.CreateFrame(FR_Root)
 Set FR_LongBuilding3 = D3D_Main.CreateFrame(FR_Root)
 Set FR_LongBuilding4 = D3D_Main.CreateFrame(FR_Root)
 Set FR_LOngBuilding5 = D3D_Main.CreateFrame(FR_Root)
 Set FR_WideBuilding2 = D3D_Main.CreateFrame(FR_Root)
 Set FR_WideBuilding3 = D3D_Main.CreateFrame(FR_Root)
 Set FR_MidBuilding4 = D3D_Main.CreateFrame(FR_Root)
 Set FR_WideBuilding4 = D3D_Main.CreateFrame(FR_Root)
 Set FR_midBuilding5 = D3D_Main.CreateFrame(FR_Root)
 Set FR_Wearhouse1 = D3D_Main.CreateFrame(FR_Root)
 Set FR_MidBuilding6 = D3D_Main.CreateFrame(FR_Root)
 Set FR_MidBuilding7 = D3D_Main.CreateFrame(FR_Root)
 Set FR_Stands = D3D_Main.CreateFrame(FR_Root)
 Set FR_Car = D3D_Main.CreateFrame(FR_Root)
 
 DS_Front.DrawText 200, 280, "Loading Textures...", False
 
 'That above code set up the hierarchy of frames where FR_Root is the parent, and the other frames are all
 'owned by it.
 'Set Textures
 Set texStreet = D3D_Main.LoadTexture(App.Path & "\street.bmp")
 Set texBetterBricks = D3D_Main.LoadTexture(App.Path & "\betterbricks.bmp")
 Set texWindows1 = D3D_Main.LoadTexture(App.Path & "\windows1.bmp")
 Set texWindows2 = D3D_Main.LoadTexture(App.Path & "\windows2.bmp")
 Set texWindowMix = D3D_Main.LoadTexture(App.Path & "\windowmix.bmp")
 Set texStands = D3D_Main.LoadTexture(App.Path & "\stands.bmp")
 
 DS_Front.DrawText 200, 310, "Loading Meshes and applying textures...", False
 
 FR_Root.SetSceneBackgroundRGB 0, 0, 0 'Set the background color. Use decimals, not the standerd 255 = max.
                                                         'What I have here will make the background/sky 100% blue.
 FR_Camera.SetPosition Nothing, 1, 9, -55 'Set The Camera position; X=1, Y=4, z=-35.
 Set D3D_ViewPort = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 640, 480) 'Make our viewport and set
                                                                                                                  'it our camera to be it.
 D3D_ViewPort.SetBack 200 'How far back it will draw the image. (Kinda like a visibility limit)
  
 FR_Light.SetPosition Nothing, 0, 10, 0 'Set our 'point' light.
 Set LT_Spot = D3D_Main.CreateLightRGB(D3DRMLIGHT_POINT, 1, 0.7, 0.4) 'Set the light type and it's color.
 FR_Light.AddLight LT_Spot 'Add the light to it's frame.
  
 Set LT_Ambient = D3D_Main.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.3, 0.3, 0.3) 'Create our ambient light.
 FR_Root.AddLight LT_Ambient 'Add the ambient light to the root frame.

 Set MS_Street = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
 MS_Street.LoadFromFile App.Path & "\street.x", 0, 0, Nothing, Nothing 'Load our building mesh from its .X file.
 MS_Street.ScaleMesh 3, 3, 3 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
                                          'it the same size as it was built in whatever program it was built in.
 MS_Street.SetTexture texStreet
 FR_Street.AddVisual MS_Street 'Add the 3D Building mesh to it's frame.
 

 
 
 Set MS_WideBuilding1 = D3D_Main.CreateMeshBuilder()
 MS_WideBuilding1.LoadFromFile App.Path & "\widebuilding.x", 0, 0, Nothing, Nothing
 MS_WideBuilding1.ScaleMesh 0.7, 0.7, 0.7
 MS_WideBuilding1.SetTexture texBetterBricks
 FR_WideBuilding1.AddVisual MS_WideBuilding1
 FR_WideBuilding1.SetPosition Nothing, 13, 1, -14
 FR_WideBuilding1.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 1.5705
 
 Set MS_midBuilding1 = D3D_Main.CreateMeshBuilder()
 MS_midBuilding1.LoadFromFile App.Path & "\midbuilding.x", 0, 0, Nothing, Nothing
 MS_midBuilding1.ScaleMesh 0.7, 0.7, 0.7
 MS_midBuilding1.SetTexture texWindows1
 FR_midBuilding1.AddVisual MS_midBuilding1
 FR_midBuilding1.SetPosition Nothing, 1, 1, -30
 
 Set MS_BigBuilding1 = D3D_Main.CreateMeshBuilder()
 MS_BigBuilding1.LoadFromFile App.Path & "\bigbuilding.x", 0, 0, Nothing, Nothing
 MS_BigBuilding1.ScaleMesh 0.7, 0.7, 0.7
 MS_BigBuilding1.SetTexture texWindows1
 FR_BigBuilding1.AddVisual MS_BigBuilding1
 FR_BigBuilding1.SetPosition Nothing, 5, 1, 6

 Set MS_midBuilding2 = D3D_Main.CreateMeshBuilder()
 MS_midBuilding2.LoadFromFile App.Path & "\midbuilding.x", 0, 0, Nothing, Nothing
 MS_midBuilding2.ScaleMesh 0.7, 0.7, 0.7
 MS_midBuilding2.SetTexture texWindows2
 FR_MidBuilding2.AddVisual MS_midBuilding2
 FR_MidBuilding2.SetPosition Nothing, 11, 1, 27
 
 Set MS_midBuilding3 = D3D_Main.CreateMeshBuilder()
 MS_midBuilding3.LoadFromFile App.Path & "\midbuilding.x", 0, 0, Nothing, Nothing
 MS_midBuilding3.ScaleMesh 0.7, 0.7, 0.7
 MS_midBuilding3.SetTexture texWindows1
 FR_midBuilding3.AddVisual MS_midBuilding3
 FR_midBuilding3.SetPosition Nothing, -10, 1, 27

 Set MS_longbuilding1 = D3D_Main.CreateMeshBuilder()
 MS_longbuilding1.LoadFromFile App.Path & "\longbuilding.x", 0, 0, Nothing, Nothing
 MS_longbuilding1.ScaleMesh 0.7, 0.7, 0.7
 MS_longbuilding1.SetTexture texWindowMix
 FR_longbuilding1.AddVisual MS_longbuilding1
 FR_longbuilding1.SetPosition Nothing, 1, 1, -70
 FR_longbuilding1.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 6.33
 
 Set MS_longbuilding2 = D3D_Main.CreateMeshBuilder()
 MS_longbuilding2.LoadFromFile App.Path & "\longbuilding.x", 0, 0, Nothing, Nothing
 MS_longbuilding2.ScaleMesh 0.7, 0.7, 0.7
 MS_longbuilding2.SetTexture texWindowMix
 FR_longbuilding2.AddVisual MS_longbuilding2
 FR_longbuilding2.SetPosition Nothing, 1, 1, -70
 FR_longbuilding2.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 3.141
 
 Set MS_LongBuilding3 = D3D_Main.CreateMeshBuilder()
 MS_LongBuilding3.LoadFromFile App.Path & "\longbuilding.x", 0, 0, Nothing, Nothing
 MS_LongBuilding3.ScaleMesh 0.7, 0.7, 0.7
 MS_LongBuilding3.SetTexture texWindows1
 FR_LongBuilding3.AddVisual MS_LongBuilding3
 FR_LongBuilding3.SetPosition Nothing, -60, 1, -20
 FR_LongBuilding3.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 1.5705

' Back long Buildings

 Set MS_LongBuilding4 = D3D_Main.CreateMeshBuilder()
 MS_LongBuilding4.LoadFromFile App.Path & "\longbuilding.x", 0, 0, Nothing, Nothing
 MS_LongBuilding4.ScaleMesh 0.7, 0.7, 0.7
 MS_LongBuilding4.SetTexture texWindows2
 FR_LongBuilding4.AddVisual MS_LongBuilding4
 FR_LongBuilding4.SetPosition Nothing, 1, 1, 70
 FR_LongBuilding4.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 6.33
 
 Set MS_LongBuilding5 = D3D_Main.CreateMeshBuilder()
 MS_LongBuilding5.LoadFromFile App.Path & "\longbuilding.x", 0, 0, Nothing, Nothing
 MS_LongBuilding5.ScaleMesh 0.7, 0.7, 0.7
 MS_LongBuilding5.SetTexture texBetterBricks
 FR_LOngBuilding5.AddVisual MS_LongBuilding5
 FR_LOngBuilding5.SetPosition Nothing, -5, 1, 70
 FR_LOngBuilding5.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 3.141
 
 Set MS_WideBuilding2 = D3D_Main.CreateMeshBuilder()
 MS_WideBuilding2.LoadFromFile App.Path & "\widebuilding.x", 0, 0, Nothing, Nothing
 MS_WideBuilding2.ScaleMesh 1, 1, 1
 MS_WideBuilding2.SetTexture texBetterBricks
 FR_WideBuilding2.AddVisual MS_WideBuilding2
 FR_WideBuilding2.SetPosition Nothing, -69, 1, 17
 FR_WideBuilding2.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 1.5705
 
 Set MS_WideBuilding3 = D3D_Main.CreateMeshBuilder()
 MS_WideBuilding3.LoadFromFile App.Path & "\widebuilding.x", 0, 0, Nothing, Nothing
 MS_WideBuilding3.ScaleMesh 1.3, 1.3, 1.3
 MS_WideBuilding3.SetTexture texWindows2
 FR_WideBuilding3.AddVisual MS_WideBuilding3
 FR_WideBuilding3.SetPosition Nothing, -72, 1, 70
 FR_WideBuilding3.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 1.5705
 
 Set MS_MidBuilding4 = D3D_Main.CreateMeshBuilder()
 MS_MidBuilding4.LoadFromFile App.Path & "\midbuilding.x", 0, 0, Nothing, Nothing
 MS_MidBuilding4.ScaleMesh 1, 1, 1
 MS_MidBuilding4.SetTexture texWindows1
 FR_MidBuilding4.AddVisual MS_MidBuilding4
 FR_MidBuilding4.SetPosition Nothing, 1, 1, 75
 
 FR_WideBuilding4.AddVisual MS_WideBuilding1
 FR_WideBuilding4.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 1.5705
 FR_WideBuilding4.SetPosition Nothing, 57, 1, 60
 
 FR_midBuilding5.AddVisual MS_MidBuilding4
 FR_midBuilding5.SetPosition Nothing, 62, 1, 26
 
 Set MS_Wearhouse1 = D3D_Main.CreateMeshBuilder()
 MS_Wearhouse1.LoadFromFile App.Path & "\wearhouse.x", 0, 0, Nothing, Nothing
 MS_Wearhouse1.ScaleMesh 1, 1, 1
 MS_Wearhouse1.SetTexture texBetterBricks
 FR_Wearhouse1.AddVisual MS_Wearhouse1
 FR_Wearhouse1.SetPosition Nothing, 90, 1, -15

 FR_MidBuilding6.AddVisual MS_MidBuilding4
 FR_MidBuilding6.SetPosition Nothing, 62, 1, -55
 
 FR_MidBuilding7.AddVisual MS_midBuilding2
 FR_MidBuilding7.SetPosition Nothing, 62, 1, -70
 
 Set MS_Stands = D3D_Main.CreateMeshBuilder()
 MS_Stands.LoadFromFile App.Path & "\stands.x", 0, 0, Nothing, Nothing
 MS_Stands.ScaleMesh 0.4, 0.4, 0.4
 MS_Stands.SetTexture texStands
 FR_Stands.AddVisual MS_Stands
 FR_Stands.SetPosition Nothing, -25, 6, -17
 
 'Set MS_Car = D3D_Main.CreateMeshBuilder()
 'MS_Car.LoadFromFile App.Path & "\car.x", 0, 0, Nothing, Nothing
 'MS_Car.ScaleMesh 0.5, 0.5, 0.5
 'FR_Car.AddVisual MS_Car
 'FR_Car.SetPosition Nothing, 1, 10, -40
 
 FR_Root.SetSceneFogMode D3DRMFOG_LINEAR
 FR_Root.SetSceneFogColor RGB(0, 0, 0)
 FR_Root.SetSceneFogEnable True
 
 FR_Root.SetZbufferMode D3DRMZBUFFER_ENABLE
 
 
End Sub '1,5,-55

'=====================================================================================

Private Sub DX_Render()
 'Lets put our main loop. Make it loop until esc = true (I'll explain later)
 Do While esc = False
   On Local Error Resume Next 'Incase there is an error
   DoEvents 'Give the computer time to do what it needs to do.

   DX_Input 'Call the input sub.


   If Go = False Then speed = speed - 0.05
   
   If speed <= 0 Then: Else FR_Camera.SetPosition FR_Camera, 0, 0, speed
   
   D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'Clear your viewport.
   D3D_Device.Update 'Update the Direct3D Device.
   D3D_ViewPort.Render FR_Root 'Render the 3D Objects (lights, and your building!)
   DS_Back.DrawText 200, 0, tXt, False  'Draw some text!
   DS_Front.Flip Nothing, DDFLIP_WAIT 'Flip the back buffer with the front buffer.
 Loop
End Sub

'=====================================================================================

Private Sub DX_Input()


 
 DI_Device.GetDeviceStateKeyboard DI_State 'Get the array of keyboard keys and their current states
  
 If DI_State.Key(DIK_ESCAPE) <> 0 Then Call DX_Exit 'If user presses [esc] then exit end the program.
 
 'If DI_State.Key(DIK_LEFT) <> 0 Then 'Quick Note: <> means 'does not'
 '  FR_Camera.SetOrientation FR_Camera, -Sin5, 0, Cos5, 0, 1, 0 'Rotate viewport/camera left
 'End If
 
 'If DI_State.Key(DIK_RIGHT) <> 0 Then
 '  FR_Camera.SetOrientation FR_Camera, Sin5, 0, Cos5, 0, 1, 0 'Rotate viewport/camera right
 'End If
 
 'If DI_State.Key(DIK_UP) <> 0 Then
  ' FR_Camera.SetPosition FR_Camera, 0, 0, 1 'Move the viewport forward
 'End If

 'If DI_State.Key(DIK_DOWN) <> 0 Then
 '  FR_Camera.SetPosition FR_Camera, 0, 0, -1 'Move the viewport back
' End If

 If xxxx < 200 Then
   FR_Camera.SetOrientation FR_Camera, -Sin5, 0, Cos5, 0, 1, 0 'Rotate viewport/camera left
 End If
 
 If xxxx > 440 Then
   FR_Camera.SetOrientation FR_Camera, Sin5, 0, Cos5, 0, 1, 0 'Rotate viewport/camera left
 End If
End Sub

'=====================================================================================

Private Sub DX_Exit()
 esc = True
 Call DD_Main.RestoreDisplayMode
 Call DD_Main.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
  Call DI_Device.Unacquire
 'Restore all the devices

 End 'Ends the program.
End Sub

'=====================================================================================

Private Sub Form_Load()
 speed = 0
 Me.Show 'Some computers do weird stuff if you don't show the form.
 DoEvents 'Give the computer time to do what it needs to do
 DX_Init 'Initialize DirectX
 DX_MakeObjects 'Make frames, lights, and mesh(es)
 DX_Render 'The Main Loop
 'One More Line of code after this!!!
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then speed = 2:  Go = True



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
xxxx = x

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then Go = False


End Sub
