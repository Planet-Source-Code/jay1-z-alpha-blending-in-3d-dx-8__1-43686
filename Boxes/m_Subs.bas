Attribute VB_Name = "m_Subs"
Option Explicit

Public Sub sub_Setup()
    '==========================
    'Add Code For Setup Options
    '==========================
End Sub

Public Sub sub_CreateObjects()
    Set VertexBuffer = D3DDevice.CreateVertexBuffer _
                                            (24 * Len(VERTEX(1)), _
                                            0, D3DFVF_CUSTOMVERTEX, _
                                            D3DPOOL_DEFAULT)
    If VertexBuffer Is Nothing Then MsgBox "VertexBuffer problem!"
    Set IndexBuffer = D3DDevice.CreateIndexBuffer(36 * Len(indIndices(1)), _
                                                0, D3DFMT_INDEX16, _
                                                D3DPOOL_DEFAULT)
    If IndexBuffer Is Nothing Then MsgBox "IndexBuffer problem!"
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, _
                                        App.Path & "\Texture.bmp")
    If Texture Is Nothing Then MsgBox "Texture problem!"
    Call sub_CreateVerts 'Create The Vertices
    '==========================
    'Copy the indices into the index buffer
    D3DIndexBuffer8SetData IndexBuffer, 0, 36 * Len(indIndices(1)), _
                                                0, indIndices(1)
    '==========================
    'Set the vertex format
    D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    '==========================
    'Set the vertex and index buffers as current ones to render from
    D3DDevice.SetStreamSource 0, VertexBuffer, vertexSize
    D3DDevice.SetIndices IndexBuffer, -1
    D3DVertexBuffer8SetData VertexBuffer, 0, Len(verts(1)) * 24, _
                                                0, verts(1)
End Sub

Public Sub sub_Render()
On Error GoTo RenderFailed
    Call sub_Matrix 'Setup the "Camera"
    '==========================
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or _
                                D3DCLEAR_ZBUFFER, _
                                RGB(0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    Call sub_SurfaceOptions
    Call sub_RenderOptions
    Call sub_LightOptions
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    Exit Sub
RenderFailed:
    frmMain.WindowState = vbMinimized
    MsgBox "Error Rendering!"
    End
End Sub

Sub sub_Matrix()
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
    D3DXMatrixLookAtLH matView, _
        D3DVec(Cos(Timer) * 4, Sin(Timer) * 4, 5#), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1, 0.1, 100
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

Function D3DVec(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    With D3DVec
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Sub sub_LightOptions()
    '==========================
    '==========LIGHTS==========
    With Light
        .Type = D3DLIGHT_POINT
        .diffuse.a = 255#
        .diffuse.r = 255#
        .diffuse.g = 255#
        .diffuse.b = 255#
        .Position = D3DVec(0#, 0#, 0#)
        .Direction = D3DVec(0#, 0#, 0#)
        .Attenuation0 = 0.3
        .Range = 100#
        .specular = Material.Ambient
    End With
    D3DDevice.SetLight 0, Light
    D3DDevice.LightEnable 0, 0
    D3DDevice.SetRenderState D3DRS_AMBIENT, vbWhite
    D3DDevice.SetRenderState D3DRS_LIGHTING, frmMain.chkLights.Value
End Sub

Public Sub sub_RenderOptions()
    '==========================
    '====RENDERING OPTIONS=====
    'D3DDevice.SetRenderState D3DRS_SPECULARENABLE, 1
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    'Transparency Options
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, frmMain.chkAlpha.Value
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    '==========================
    If frmMain.chkIndex = Checked Then
        D3DDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 36, 0, frmMain.udcPrimCount.Value 'DrawPrimitive D3DPT_TRIANGLELIST, 0, 8
    Else
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 1, 6
    End If
End Sub

Sub sub_CleanUp()
    frmMain.tmrRender.Enabled = False
    Set Texture = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set DX8 = Nothing
    Unload frmMain
    End
End Sub

Public Sub sub_SurfaceOptions()
    '==========================
    '========MATERIALS=========
    With Material.Ambient
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    Material.diffuse = Material.Ambient
    With Material.specular
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    Material.power = 200
    D3DDevice.SetMaterial Material
    '==========================
    '==========TEXTURE=========
    D3DDevice.SetTexture 0, Texture
    '==========================
    '======LEVEL OF DETAIL=====
    Texture.SetLOD 0 'Lower is better
End Sub
