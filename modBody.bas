Attribute VB_Name = "modBody"
Option Explicit

Public Type tAABB
    pMin    As tVec2
    pMax    As tVec2
End Type

Public Enum eBodyType
    eCircle = 0    '&H0
    ePolygon = 1    '&H1
End Enum


Public Type tBody
    myType  As eBodyType
    mass    As Double
    invMass As Double
    inertia As Double
    invInertia As Double

    ii      As Double


    Area    As Double


    COM     As tVec2

    '------------------
    'Circle
    radius  As Double
    '------------------
    'Poligon
    Vertex() As tVec2
    tVertex() As tVec2

    normals() As tVec2
    Nvertex As Long
    '------------------


    Pos     As tVec2
    VEL     As tVec2
    FORCE   As tVec2
    angularVelocity As Double
    torque  As Double
    orient  As Double

    'Material
    staticFriction As Double
    dynamicFriction As Double
    Restitution As Double

    U       As tMAT2

    color   As Long

    AABB    As tAABB


End Type


Public Body() As tBody
Public NofBodies As Long




Private Sub CalcCentroid(wB As Long)


'// Calculate centroid and moment of inertia

    Dim triangleArea As Double


    Const k_inv3 As Double = 1 / 3

    Dim I   As Long
    Dim J   As Long

    Dim P1  As tVec2
    Dim p2  As tVec2
    Dim D   As Double

    Dim weight As Double

    Dim intx2 As Double
    Dim inty2 As Double

    With Body(wB)

        .ii = 0
        .COM.X = 0
        .COM.Y = 0
        .Area = 0

        For I = 1 To .Nvertex
            J = I + 1
            If J > .Nvertex Then J = 1
            '            // Triangle vertices, third vertex implied as (0, 0)
            P1 = .Vertex(I)
            p2 = .Vertex(J)

            D = Vec2CROSS(P1, p2)
            triangleArea = 0.5 * D

            .Area = .Area + triangleArea

            '            // Use area to weight the centroid average, not just vertex position
            weight = triangleArea * k_inv3

            '            com.addsi( p1, weight );
            '            com.addsi( p2, weight );

            .COM = Vec2ADD(.COM, Vec2MUL(P1, weight))
            .COM = Vec2ADD(.COM, Vec2MUL(p2, weight))


            intx2 = P1.X * P1.X + p2.X * P1.X + p2.X * p2.X
            inty2 = P1.Y * P1.Y + p2.Y * P1.Y + p2.Y * p2.Y

            .ii = .ii + (0.25 * k_inv3 * D) * (intx2 + inty2)

        Next

        'com.muli( 1.0f / area );
        .COM = Vec2MUL(.COM, 1 / .Area)

        .ii = .ii / .Area    '-----------<<<<<<<<<<<<<<<< main missing line causing polygons not to rotate


    End With

End Sub


Public Sub ComputeMass(wB As Long, Density As Double)
    Dim I   As Long

    With Body(wB)

        If .myType = eCircle Then
            .mass = PI * .radius * .radius * Density
            .invMass = IIf(.mass <> 0#, 1# / .mass, 0)
            .inertia = .mass * .radius * .radius
            .invInertia = IIf(.inertia <> 0#, 1# / .inertia, 0#)
        End If

        If .myType = ePolygon Then

            CalcCentroid wB

            ' Translate vertices to centroid (make the centroid (0, 0)
            ' for the polygon in model space)
            ' Not really necessary, but I like doing this anyway
            For I = 1 To .Nvertex
                .Vertex(I) = Vec2SUB(.Vertex(I), .COM)
            Next

            .mass = Density * .Area
            .invMass = IIf(.mass <> 0#, 1# / .mass, 0)
            .inertia = .ii * Density
            .invInertia = IIf(.inertia <> 0#, 1# / .inertia, 0#)
        End If


    End With

End Sub




Public Sub BodyApplyForce(wB As Long, F As tVec2)
    With Body(wB)
        .FORCE = Vec2ADD(.FORCE, F)
    End With
End Sub

Public Sub BodyApplyImpulse(wB As Long, impulse As tVec2, contactVector As tVec2)
'      velocity.addsi( impulse, invMass );
'      angularVelocity += invInertia * Vec2.cross( contactVector, impulse );
    With Body(wB)
        .VEL = Vec2ADD(.VEL, Vec2MUL(impulse, .invMass))
        .angularVelocity = .angularVelocity + Vec2CROSS(contactVector, impulse) * .invInertia
    End With
End Sub

Public Sub BodySetStatic(wB As Long)
    With Body(wB)
        .inertia = 0
        .invInertia = 0
        .mass = 0#
        .invMass = 0
    End With

End Sub

Public Sub POLYGONComputeFaceNormals(wB As Long)
'
'        // Compute face normals
'        for (int i = 0; i < vertexCount; ++i)
'        {
'            Vec2 face = vertices[(i + 1) % vertexCount].sub( vertices[i] );
'
'            // Calculate normal with 2D cross product between vector and scalar
'            normals[i].set( face.y, -face.x );
'            normals[i].normalize();
'        }

    If Body(wB).myType <> ePolygon Then Exit Sub

    Dim I   As Long
    Dim J   As Long
    Dim face As tVec2
    Dim N   As tVec2

    With Body(wB)
        ReDim .normals(.Nvertex)
        For I = 1 To .Nvertex
            J = I + 1
            If J > .Nvertex Then J = 1
            face = Vec2SUB(.Vertex(J), .Vertex(I))
            N.X = face.Y
            N.Y = -face.X
            .normals(I) = Vec2Normalize(N)
        Next
    End With



End Sub
Public Sub CREATECircle(Pos As tVec2, r As Double, Density As Double)
    NofBodies = NofBodies + 1
    ReDim Preserve Body(NofBodies)
    With Body(NofBodies)
        .myType = eCircle
        .Pos = Pos
        .radius = r
        .staticFriction = GlobalSTATICFRICTION   ' 0.15    '0.3    '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    ' 0.5 '0.07    ' 0.1    '0.3
        .Restitution = GlobalRestitution
        .orient = rndFT(-PI, PI)
         .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

    End With


    ComputeMass NofBodies, Density
End Sub

Public Sub CREATERandomPoly(Pos As tVec2, Density As Double)
    Dim I   As Long


    NofBodies = NofBodies + 1
    ReDim Preserve Body(NofBodies)
    With Body(NofBodies)
        .myType = ePolygon

        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = rndFT(-PI, PI)

        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .Nvertex = 4 + Rnd * 2

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        For I = 1 To .Nvertex
            '        For I = .Nvertex To 1 Step -1

            .Vertex(I) = Vec2ADD(Pos, _
                                 Vec2((10 + Rnd * 30) * Cos(PI2 * (I - 1) / .Nvertex), _
                                      (10 + Rnd * 30) * Sin(PI2 * (I - 1) / .Nvertex)))
        Next


    End With


    POLYGONComputeFaceNormals NofBodies

    ComputeMass NofBodies, Density


 Body(NofBodies).COM = Vec2SUB(Body(NofBodies).Pos, Body(NofBodies).COM)

End Sub


'  // The extreme point along a direction within a polygon
'  Vec2 GetSupport( const Vec2& dir )
'  {
'    real bestProjection = -FLT_MAX;
'    Vec2 bestVertex;
'
'    for(uint32 i = 0; i < m_vertexCount; ++i)
'    {
'      Vec2 v = m_vertices[i];
'      real projection = Dot( v, dir );
'
'      if(projection > bestProjection)
'      {
'        bestVertex = v;
'        bestProjection = projection;
'      }
'    }
'
'    return bestVertex;
'  }
Public Function GetSupport(Body As tBody, dire As tVec2) As tVec2
'// The extreme point along a direction within a polygon
    Dim bestProjection As Double
    Dim bestVertex As tVec2
    Dim V   As tVec2
    Dim I   As Long
    Dim projection As Double

    bestProjection = -FLT_MAX
    For I = 1 To Body.Nvertex
        V = Body.Vertex(I)
        projection = Vec2DOT(V, dire)

        If (projection > bestProjection) Then
            bestVertex = V
            bestProjection = projection
        End If
    Next

    GetSupport = bestVertex
End Function

