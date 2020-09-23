Attribute VB_Name = "m_Declare"
Option Explicit

Public DX8 As DirectX8
Public D3D As Direct3D8
Public D3DX As New D3DX8
Public D3DDevice As Direct3DDevice8
Public VertexBuffer As Direct3DVertexBuffer8
Public VERTEX(25) As D3DVERTEX
Public Material As D3DMATERIAL8
Public Texture As Direct3DTexture8
Public Light As D3DLIGHT8
Public IndexBuffer As Direct3DIndexBuffer8
Public indIndices(1 To 42) As Integer
Public intTemp As Integer
Public verts(1 To 24) As VERTEX
Public vertexSize As Single
Public sizeVert As VERTEX
Public Const PI = 3.14159275180032
'Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_VERTEX)

Type VERTEX
    X As Single
    Y As Single
    Z As Single
    Colour As Long
    NX As Single
    NY As Single
    NZ As Single
    tu As Single
    tv As Single
End Type

Private Type CUSTOMVERTEX
    postion As D3DVECTOR    '3-D position for vertex.
    color As Long           'Color of the vertex.
    tu As Single            'Texture map coordinate.
    tv As Single            'Texture map coordinate.
End Type
