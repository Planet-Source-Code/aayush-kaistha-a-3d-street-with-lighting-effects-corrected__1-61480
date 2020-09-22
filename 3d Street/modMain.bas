Attribute VB_Name = "modMain"
'########################################################
'
'       Walk through an exotic 3d locality
'
'       Written By: Aayush Kaistha
'       Place:      UIET, Panjab University, Chandigarh
'       Contact:    aayushk_007@yahoo.com
'
'   Special thanx 2 Jack Hoxley (externalweb.exhedra.com/directx4vb)
'   for his gr8 tutorials
'
'########################################################

'3d objects were created using 3d studio max and exported
'in .3ds format which were then converted into .x format
'using the program conv3ds.exe available with direct x sdk

Option Explicit

Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim D3DX As D3DX8
Dim D3DDevice As Direct3DDevice8

Public bRunning As Boolean
Public UpKey As Boolean, DownKey As Boolean
Public LeftKey As Boolean, RightKey As Boolean
Public WKey As Boolean, SKey As Boolean

Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private Type VERTEX
    P As D3DVECTOR
    N As D3DVECTOR
    T As D3DVECTOR2
End Type

'all these variables r required to load 3d objects in directX
Private Type Object3D
    nMaterials As Long
    Materials() As D3DMATERIAL8
    Textures() As Direct3DTexture8
    TextureFile As String
    Mesh As D3DXMesh
End Type
    
Private Type Plyr_Data
    Pos As D3DVECTOR
    Rotation As Integer
    MoveSpeed As Single
End Type

Private Declare Function GetTickCount Lib "kernel32" () As Long

'this only holds data req to calculate frames per second
Private Type FPS_data
    Count As Long
    Value As Long
    Last As Long
End Type

Dim fps As FPS_data

Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim fnt As New StdFont

'this holds the no of 3d objects we r loading. we have
'no_obj + 1 objects in our prog
Const no_obj = 6

'we store values of sin & cos in array 2 avoid calculation
'at run time
Dim XSin(360) As Single, XCos(360) As Single
Dim Player As Plyr_Data
Dim Obj(no_obj) As Object3D 'array of 3d objects

'we req a matrix 4 each object. we use matrices to modify
'vertices of an object so that it can b rotated, translated
'or scaled
Dim matObj(no_obj) As D3DMATRIX

Dim matProj As D3DMATRIX 'this holds the camera settings
Dim matView As D3DMATRIX 'this tells where the camera is n where it is looking at
Dim matWorld As D3DMATRIX 'this holds the reference coordinates of entire 3d world

Const PI = 3.14159
Const RAD = PI / 180

Private Sub StoreSinCos()

'if we calculate sines and cosines at run time, it slows
'down the prog. we instead store the values of sine n cos
'in an array so that we do not have to cal them at run time
Dim I As Integer, Ang As Double

For I = 0 To 360
    Ang = I * RAD
    XCos(I) = Cos(Ang)
    XSin(I) = Sin(Ang)
Next

End Sub

Private Function Initialize() As Boolean

On Error GoTo Hell:

Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE

'initialize and allocate memory 4 directX variables
Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate
Set D3DX = New D3DX8

DispMode.Format = CheckDisplayMode(640, 480, 32)
If DispMode.Format > D3DFMT_UNKNOWN Then
    Debug.Print "Using 32-Bit format"
Else
    DispMode.Format = CheckDisplayMode(640, 480, 16)
    If DispMode.Format > D3DFMT_UNKNOWN Then
        Debug.Print "32-Bit format not supported. Using 16-Bit format"
    Else
        MsgBox "Neither 16-Bit nor 32-Bit Display Mode Supported", vbInformation, "ERROR"
        Unload frmMain
        End
    End If
End If

'this sets our resolution settings. here we r using 16-bit
'screen format so that it runs on older computers as well
With D3DWindow
    .BackBufferCount = 1
    .BackBufferFormat = DispMode.Format
    .BackBufferWidth = 640
    .BackBufferHeight = 480
    .hDeviceWindow = frmMain.hWnd
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
End With

If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
    D3DWindow.EnableAutoDepthStencil = 1
    MsgBox "32-bit"
Else
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D24X8
        D3DWindow.EnableAutoDepthStencil = 1
    Else
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
            D3DWindow.EnableAutoDepthStencil = 1
        Else
            D3DWindow.EnableAutoDepthStencil = 0
            MsgBox "Depth buffer could not be enabled", vbInformation, "Depth buffer not supported"
        End If
    End If
End If

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

With D3DDevice
    .SetVertexShader FVF_VERTEX
    .SetRenderState D3DRS_LIGHTING, 1
    .SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150)
    .SetRenderState D3DRS_ZENABLE, 1
    .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
End With

'initialize our matrices
D3DXMatrixIdentity matWorld
D3DDevice.SetTransform D3DTS_WORLD, matWorld

D3DXMatrixLookAtLH matView, MakeVector(-2, 2, -2), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView

D3DXMatrixPerspectiveFovLH matProj, PI / 3, 1, 0.1, 500
D3DDevice.SetTransform D3DTS_PROJECTION, matProj

'font settings
fnt.Name = "Verdana"
fnt.Size = 10
fnt.Bold = True
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

LoadModels
SetLights

Initialize = True

Exit Function
Hell:
MsgBox "ERROR initializing D3D ", vbCritical, "ERROR"
Initialize = False

End Function

Private Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.Z = Z
End Function

Private Function MakeRect(Left As Single, Right As Single, Top As Single, Bottom As Single) As RECT

MakeRect.Left = Left
MakeRect.Right = Right
MakeRect.Top = Top
MakeRect.Bottom = Bottom
 
End Function

Private Sub LoadModels()

On Error GoTo Out:

Dim mtrlBuffer As D3DXBuffer
Dim I As Long, j As Integer, fname(6) As String

fname(0) = App.Path + "\0.x"
fname(1) = App.Path + "\1.x"
fname(2) = App.Path + "\2.x"
fname(3) = App.Path + "\3.x"
fname(4) = App.Path + "\4.x"
fname(5) = App.Path + "\5.x"
fname(6) = App.Path + "\6.x"

'this loop loads 3d objects in obj() array
For j = 0 To no_obj
    Set Obj(j).Mesh = D3DX.LoadMeshFromX(fname(j), D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, Obj(j).nMaterials)
    
    ReDim Obj(j).Materials(Obj(j).nMaterials) As D3DMATERIAL8
    ReDim Obj(j).Textures(Obj(j).nMaterials) As Direct3DTexture8

    For I = 0 To Obj(j).nMaterials - 1
        D3DX.BufferGetMaterial mtrlBuffer, I, Obj(j).Materials(I)
        Obj(j).Materials(I).Ambient = Obj(j).Materials(I).diffuse
        Obj(j).TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, I)
        If Obj(j).TextureFile <> "" Then
            Set Obj(j).Textures(I) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\" + Obj(j).TextureFile)
        End If
    Next
Next

Exit Sub
Out:
    MsgBox "Error loading models", vbCritical, "ERROR"
End Sub

Private Sub SetLights()

Dim light(3) As D3DLIGHT8, j As Integer

light(0).Type = D3DLIGHT_POINT
With light(0).diffuse
    .a = 1: .b = 1: .g = 1: .r = 0
End With
light(0).Position = MakeVector(0, 70, 0)
light(0).Range = 100#
light(0).Attenuation0 = 1#
light(0).Attenuation1 = 0#
light(0).Attenuation2 = 0#

light(1).Type = D3DLIGHT_POINT
With light(1).diffuse
    .a = 1: .b = 0: .g = 1: .r = 0
End With
light(1).Position = MakeVector(180, 70, 150)
light(1).Range = 100#
light(1).Attenuation0 = 1#
light(1).Attenuation1 = 0#
light(1).Attenuation2 = 0#

light(2).Type = D3DLIGHT_POINT
With light(2).diffuse
    .a = 1: .b = 1: .g = 0: .r = 1
End With
light(2).Position = MakeVector(180, 70, 20)
light(2).Range = 100#
light(2).Attenuation0 = 1#
light(2).Attenuation1 = 0#
light(2).Attenuation2 = 0#

light(3).Type = D3DLIGHT_POINT
With light(3).diffuse
    .a = 1: .b = 0: .g = 1: .r = 1
End With
light(3).Position = MakeVector(-180, 50, 0)
light(3).Range = 100#
light(3).Attenuation0 = 1#
light(3).Attenuation1 = 0#
light(3).Attenuation2 = 0#

For j = 0 To 3
    D3DDevice.SetLight j, light(j)
    D3DDevice.LightEnable j, 1
Next

End Sub

Private Sub Main()

frmMain.Show

Dim matTemp As D3DMATRIX, j As Integer, LastUpdated As Long
Dim Angle As Integer

Player.Pos.X = -100
Player.Pos.Y = 5
Player.Pos.Z = -200
Player.Rotation = 70

StoreSinCos
bRunning = Initialize
PlaceObjects

fps.Last = GetTickCount
LastUpdated = GetTickCount

Do While bRunning
    'limit the fps to max of 100 so that our prog does not
    'run too fast
    If ((GetTickCount - LastUpdated) >= 10) Then
        LastUpdated = GetTickCount
        CheckKeys
    
        'as we move using arrow keys, make the camera
        'follow our position
        D3DXMatrixLookAtLH matView, Player.Pos, MakeVector(Player.Pos.X + (XCos(Player.Rotation) * 10), 5, Player.Pos.Z + (XSin(Player.Rotation) * 10)), MakeVector(0, 1, 0)
        D3DDevice.SetTransform D3DTS_VIEW, matView
        
        '################################################
        'the globe
        D3DXMatrixIdentity matObj(5)
        D3DXMatrixIdentity matTemp
        D3DXMatrixRotationY matTemp, Angle * RAD
        D3DXMatrixMultiply matObj(5), matObj(5), matTemp
        
        D3DXMatrixIdentity matTemp
        D3DXMatrixTranslation matTemp, 0, 20, 0
        D3DXMatrixMultiply matObj(5), matObj(5), matTemp
        '################################################
        
        D3DXMatrixScaling matWorld, 1, 1, 1
        D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
        Render
    
        fps.Count = fps.Count + 1
        If ((GetTickCount - fps.Last) >= 1000) Then
            fps.Value = fps.Count
            fps.Count = 0
            fps.Last = GetTickCount
        End If
        
        DoEvents
        Player.MoveSpeed = ((GetTickCount - LastUpdated) / 1000) * 150
        Angle = Angle + (((GetTickCount - LastUpdated) / 1000) * 90)
        If Angle > 360 Then Angle = 0
        
    End If
Loop
    
Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing

End

End Sub

Private Sub Render()
Dim I As Long, j As Integer

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H33, 1#, 0

D3DDevice.BeginScene
    For j = 0 To no_obj
        D3DDevice.SetTransform D3DTS_WORLD, matObj(j)
        For I = 0 To Obj(j).nMaterials - 1
            D3DDevice.SetTexture 0, Obj(j).Textures(I)
            D3DDevice.SetMaterial Obj(j).Materials(I)
            Obj(j).Mesh.DrawSubset I
        Next
    Next
    
    D3DX.DrawText MainFont, &HFFFFCC00, "Position : [ " + Str(Player.Pos.X) + " , " + Str(Player.Pos.Z) + " ]", MakeRect(10, 320, 0, 20), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "Rotation : " + Str(Player.Rotation), MakeRect(10, 320, 20, 40), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "FPS : " + Str(fps.Value), MakeRect(10, 320, 40, 60), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "Use Arrow keys to move", MakeRect(330, 640, 0, 20), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "Press W for wire-frame geometry", MakeRect(330, 640, 20, 40), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "Press S for solid geometry", MakeRect(330, 640, 40, 60), DT_TOP Or DT_LEFT
D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub PlaceObjects()
Dim matTemp As D3DMATRIX

'this is the street
D3DXMatrixIdentity matObj(0)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 0, -20, 0
D3DXMatrixMultiply matObj(0), matObj(0), matTemp

'buildings
D3DXMatrixIdentity matObj(1)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 200, -20, 180
D3DXMatrixMultiply matObj(1), matObj(1), matTemp

D3DXMatrixIdentity matObj(2)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 200, -20, 80
D3DXMatrixMultiply matObj(2), matObj(2), matTemp

D3DXMatrixIdentity matObj(3)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 200, -20, -100
D3DXMatrixMultiply matObj(3), matObj(3), matTemp

'base below globe
D3DXMatrixIdentity matObj(4)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 0, -20, 0
D3DXMatrixMultiply matObj(4), matObj(4), matTemp

'the wall
D3DXMatrixIdentity matObj(6)
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, -200, -20, 0
D3DXMatrixMultiply matObj(6), matObj(6), matTemp

End Sub

Private Sub CheckKeys()

If LeftKey Then Player.Rotation = Player.Rotation + 2
If RightKey Then Player.Rotation = Player.Rotation - 2
If Player.Rotation < 0 Then Player.Rotation = 360
If Player.Rotation > 360 Then Player.Rotation = 0

If UpKey Then
    Player.Pos.X = Player.Pos.X + (XCos(Player.Rotation) * Player.MoveSpeed)
    Player.Pos.Z = Player.Pos.Z + (XSin(Player.Rotation) * Player.MoveSpeed)
End If
If DownKey Then
    Player.Pos.X = Player.Pos.X - (XCos(Player.Rotation) * Player.MoveSpeed)
    Player.Pos.Z = Player.Pos.Z - (XSin(Player.Rotation) * Player.MoveSpeed)
End If

If SKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
If WKey Then D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME

If Player.Pos.X > 200 Then Player.Pos.X = 200
If Player.Pos.X < -200 Then Player.Pos.X = -200
If Player.Pos.Z > 200 Then Player.Pos.Z = 200
If Player.Pos.Z < -200 Then Player.Pos.Z = -200

End Sub

Private Function CheckDisplayMode(Width As Long, Height As Long, Depth As Long) As CONST_D3DFORMAT
Dim I As Long
Dim DispMode As D3DDISPLAYMODE
    
For I = 0 To D3D.GetAdapterModeCount(0) - 1
    D3D.EnumAdapterModes 0, I, DispMode
    If DispMode.Width = Width Then
        If DispMode.Height = Height Then
            If (DispMode.Format = D3DFMT_R5G6B5) Or (DispMode.Format = D3DFMT_X1R5G5B5) Or (DispMode.Format = D3DFMT_X4R4G4B4) Then
                '16 bit mode
                If Depth = 16 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            ElseIf (DispMode.Format = D3DFMT_R8G8B8) Or (DispMode.Format = D3DFMT_X8R8G8B8) Then
                '32bit mode
                If Depth = 32 Then
                    CheckDisplayMode = DispMode.Format
                    Exit Function
                End If
            End If
        End If
    End If
Next I
CheckDisplayMode = D3DFMT_UNKNOWN
End Function
