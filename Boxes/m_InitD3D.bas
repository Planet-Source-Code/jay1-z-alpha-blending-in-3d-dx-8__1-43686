Attribute VB_Name = "m_InitD3D"
Option Explicit

Public Function sub_InitD3D(windowed As Boolean, hWnd As Long) As Boolean
'On Error GoTo initFail
Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
    '===================================
    Set DX8 = New DirectX8
        If DX8 Is Nothing Then
            MsgBox "DirectX8 Object Error!"
            End
        End If
    '===================================
    Set D3D = DX8.Direct3DCreate
        If D3D Is Nothing Then
            MsgBox "Error! Could not create a Direct3d Object."
            End
        End If
    '===================================
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    If windowed = True Then
        '===================================
        With D3DWindow
            .windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            .AutoDepthStencilFormat = D3DFMT_D16
            .EnableAutoDepthStencil = 1
            .BackBufferFormat = DispMode.Format
            .hDeviceWindow = hWnd
        End With
        '===================================
        If D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                D3DWindow.BackBufferFormat, _
                                D3DWindow.BackBufferFormat, _
                                True) = D3D_OK Then
            '===================================
            Set D3DDevice = _
                    D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                    D3DDEVTYPE_HAL, _
                    hWnd, _
                    D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                    D3DWindow)
            '===================================
        Else
            Set D3DDevice = _
               D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, _
               hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
        End If
    Else    'Fullscreen mode
        With D3DWindow
            .windowed = False
            .BackBufferHeight = 768
            .BackBufferWidth = 1024
            .BackBufferFormat = D3DFMT_R5G6B5
            .BackBufferCount = 1
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            .AutoDepthStencilFormat = D3DFMT_D16
            .EnableAutoDepthStencil = 1
            .hDeviceWindow = hWnd
        End With
        If D3D.CheckDeviceType(D3DADAPTER_DEFAULT, _
                                D3DDEVTYPE_HAL, _
                                D3DWindow.BackBufferFormat, _
                                D3DWindow.BackBufferFormat, _
                                False) = D3D_OK Then
            Set D3DDevice = _
            D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                            D3DDEVTYPE_HAL, _
                            hWnd, _
                            D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                            D3DWindow)
        Else
            Set D3DDevice = _
                D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                D3DDEVTYPE_REF, _
                                hWnd, _
                                D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                D3DWindow)
        End If
        End If
        sub_InitD3D = True
        Exit Function
initFail:
    sub_InitD3D = False
End Function


