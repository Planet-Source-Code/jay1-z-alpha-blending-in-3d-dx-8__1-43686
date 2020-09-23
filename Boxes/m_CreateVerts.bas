Attribute VB_Name = "m_CreateVerts"
Option Explicit

Public Sub sub_CreateVerts()
    vertexSize = Len(sizeVert)
    
    verts(1) = createVertex(-1, 1, -1, D3DColorARGB(255, 255, 255, 0), 0, 0, -1, 0, 0)
    verts(2) = createVertex(1, 1, -1, D3DColorARGB(255, 255, 255, 0), 0, 0, -1, 1, 0)
    verts(3) = createVertex(-1, -1, -1, D3DColorARGB(255, 255, 255, 0), 0, 0, -1, 0, 1)
    verts(4) = createVertex(1, -1, -1, D3DColorARGB(255, 255, 0, 0), 0, 0, -1, 1, 1)
    verts(5) = createVertex(-1, 1, 1, D3DColorARGB(255, 255, 0, 0), 0, 0, 1, 0, 0)
    verts(6) = createVertex(-1, -1, 1, D3DColorARGB(255, 255, 0, 0), 0, 0, 1, 0, 1)
    verts(7) = createVertex(1, 1, 1, D3DColorARGB(255, 0, 255, 0), 0, 0, 1, 1, 0)
    verts(8) = createVertex(1, -1, 1, D3DColorARGB(255, 0, 255, 0), 0, 0, 1, 1, 1)
    verts(9) = createVertex(-1, 1, 1, D3DColorARGB(255, 0, 255, 0), -1, 0, 0, 0, 0)
    verts(10) = createVertex(-1, 1, -1, D3DColorARGB(255, 0, 255, 0), -1, 0, 0, 1, 0)
    verts(11) = createVertex(-1, -1, 1, D3DColorARGB(255, 0, 255, 0), -1, 0, 0, 0, 1)
    verts(12) = createVertex(-1, -1, -1, D3DColorARGB(255, 0, 255, 0), -1, 0, 0, 1, 1)
    verts(13) = createVertex(1, 1, -1, D3DColorARGB(255, 0, 255, 0), 1, 0, 0, 0, 0)
    verts(14) = createVertex(1, 1, 1, D3DColorARGB(255, 0, 255, 0), 1, 0, 0, 1, 0)
    verts(15) = createVertex(1, -1, -1, D3DColorARGB(255, 0, 255, 0), 1, 0, 0, 0, 1)
    verts(16) = createVertex(1, -1, 1, D3DColorARGB(255, 0, 255, 0), 1, 0, 0, 1, 1)
    verts(17) = createVertex(-1, 1, -1, D3DColorARGB(255, 0, 255, 0), 0, 1, 0, 0, 0)
    verts(18) = createVertex(1, 1, -1, D3DColorARGB(255, 0, 255, 0), 0, 1, 0, 1, 0)
    verts(19) = createVertex(-1, 1, 1, D3DColorARGB(255, 0, 255, 0), 0, 1, 0, 0, 1)
    verts(20) = createVertex(1, 1, 1, D3DColorARGB(255, 0, 255, 0), 0, 1, 0, 1, 1)
    verts(21) = createVertex(-1, -1, -1, D3DColorARGB(255, 255, 0, 0), 0, -1, 0, 0, 0)
    verts(22) = createVertex(1, -1, -1, D3DColorARGB(255, 255, 0, 0), 0, -1, 0, 1, 0)
    verts(23) = createVertex(-1, -1, 1, D3DColorARGB(255, 255, 0, 0), 0, -1, 0, 0, 1)
    verts(24) = createVertex(1, -1, 1, D3DColorARGB(255, 255, 0, 0), 0, -1, 0, 1, 1)
    
    MakeIndices 1, 2, 3, _
                3, 2, 4, _
                9, 10, 11, _
                11, 10, 12, _
                13, 14, 15, _
                15, 14, 16, _
                17, 18, 19, _
                19, 18, 20, _
                21, 22, 23, _
                23, 22, 24, _
                5, 6, 7, _
                7, 6, 8

End Sub

Private Function createVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Colour As Single, ByVal NX As Single, ByVal NY As Single, ByVal NZ As Single, ByVal tu As Single, ByVal tv As Single) As VERTEX
    With createVertex
        .X = X
        .Y = Y
        .Z = Z
        .Colour = Colour
        .NX = NX
        .NY = NY
        .NZ = NZ
        .tu = tu
        .tv = tv
    End With
End Function

Function MakeIndices(ParamArray Indices()) As Integer()
    Dim i As Integer
    For i = LBound(Indices) To UBound(Indices)
        indIndices(i + 1) = Indices(i)
    Next
End Function

