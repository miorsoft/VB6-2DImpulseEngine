Attribute VB_Name = "modBody"
Option Explicit

Public Type tAABB
    pMin       As tVec2
    pMax       As tVec2
End Type

Public Enum eBodyType
    eCircle = 0    '&H0
    ePolygon = 1    '&H1
End Enum


Public Type tBody
    myType     As eBodyType
    mass       As Single
    invMass    As Single
    inertia    As Single
    invInertia As Single

    ii         As Single


    Area       As Single


    COM        As tVec2

    '------------------
    'Circle
    Radius     As Single
    '------------------
    'Poligon
    Vertex()   As tVec2
    tVertex()  As tVec2

    normals()  As tVec2
    Nvertex    As Long
    '------------------


    Pos        As tVec2
    VEL        As tVec2
    FORCE      As tVec2
    angularVelocity As Single
    torque     As Single
    orient     As Single

    'Material
    staticFriction As Single
    dynamicFriction As Single
    Restitution As Single

    U          As tMAT2

    color      As Long

    AABB       As tAABB

    CollisionGroup As Long
    CollideWith As Long


End Type


Public Body()  As tBody
Public NBodies As Long




Private Sub CalcCentroid(wB As Long)


'// Calculate centroid and moment of inertia

    Dim triangleArea As Single


    Const k_inv3 As Single = 1 / 3

    Dim I      As Long
    Dim J      As Long

    Dim P1     As tVec2
    Dim p2     As tVec2
    Dim D      As Single

    Dim weight As Single

    Dim intx2  As Single
    Dim inty2  As Single

    With Body(wB)

        .ii = 0
        .COM.x = 0
        .COM.y = 0
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

            P1 = Vec2SUB(P1, .Pos)
            p2 = Vec2SUB(p2, .Pos)

            intx2 = P1.x * P1.x + p2.x * P1.x + p2.x * p2.x
            inty2 = P1.y * P1.y + p2.y * P1.y + p2.y * p2.y

            .ii = .ii + (0.25 * k_inv3 * D) * (intx2 + inty2)

        Next
        'com.muli( 1.0f / area );
        .COM = Vec2MUL(.COM, 1 / .Area)

        '       .ii = .ii / .Area ^ 0.75 '-----------<<<<<<<<<<<<<<<< main missing line causing polygons not to rotate .... But in original source there isnt !!!???
    End With

End Sub


Public Sub ComputeMass(wB As Long, Density As Single)
    Dim I      As Long

    With Body(wB)

        If .myType = eCircle Then
            .mass = PI * .Radius * .Radius * Density
            .invMass = IIf(.mass <> 0#, 1# / .mass, 0)
            .inertia = .mass * .Radius * .Radius
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
        .inertia = MAX_VALUE
        .invInertia = 0
        .mass = 0
        .invMass = 0
    End With

End Sub

Public Sub BodySetGroup(wB As Long, G As Long)
    Body(wB).CollisionGroup = G
   If BiggerGroup < G Then BiggerGroup = G
End Sub
Public Sub BodySetCollideWith(wB As Long, M As Long)
    Body(wB).CollideWith = M
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

    Dim I      As Long
    Dim J      As Long
    Dim face   As tVec2
    Dim N      As tVec2

    If Body(wB).myType <> ePolygon Then Exit Sub

    With Body(wB)
        ReDim .normals(.Nvertex)
        For I = 1 To .Nvertex
            J = I + 1
            If J > .Nvertex Then J = 1
            face = Vec2SUB(.Vertex(J), .Vertex(I))
            N.x = face.y
            N.y = -face.x
            .normals(I) = Vec2Normalize(N)
        Next
    End With


End Sub
Public Sub CREATECircle(Pos As tVec2, r As Single, Density As Single)
    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = eCircle
        .Pos = Pos
        .Radius = r
        .staticFriction = GlobalSTATICFRICTION   ' 0.15    '0.3    '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    ' 0.5 '0.07    ' 0.1    '0.3
        .Restitution = GlobalRestitution
        .orient = rndFT(-PI, PI)
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)
        .U = SetOrient(0)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0
    End With


    ComputeMass NBodies, Density
End Sub

Public Sub CREATERandomPoly(Pos As tVec2, Density As Single)
    Dim I      As Long


    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon

        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = rndFT(-PI, PI)

        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .Nvertex = 4    '+ Rnd * 2

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        '        For I = 1 To .Nvertex
        '            '        For I = .Nvertex To 1 Step -1
        '            .Vertex(I) = Vec2ADD(Pos, _
                     '                                 Vec2((10 + Rnd * 30) * Cos(PI2 * (I - 1) / .Nvertex), _
                     '                                      (10 + Rnd * 30) * Sin(PI2 * (I - 1) / .Nvertex)))
        '        Next

        .Vertex(1) = Vec2(Pos.x - 20, Pos.y - 15)
        .Vertex(2) = Vec2(Pos.x + 40, Pos.y - 15)
        .Vertex(3) = Vec2(Pos.x + 40, Pos.y + 15)
        .Vertex(4) = Vec2(Pos.x - 20, Pos.y + 15)


    End With


    POLYGONComputeFaceNormals NBodies
    ComputeMass NBodies, Density
    Body(NBodies).Pos = Body(NBodies).COM





End Sub


Public Sub CreateBox(Pos As tVec2, W As Single, H As Single, Optional Ang As Single = 0)
    Dim I      As Long


    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon

        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = Ang
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)
        .Nvertex = 4    '+ Rnd * 2

        .U = SetOrient(Ang)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        .Vertex(1) = Vec2(Pos.x - W * 0.5, Pos.y - H * 0.5)
        .Vertex(2) = Vec2(Pos.x + W * 0.5, Pos.y - H * 0.5)
        .Vertex(3) = Vec2(Pos.x + W * 0.5, Pos.y + H * 0.5)
        .Vertex(4) = Vec2(Pos.x - W * 0.5, Pos.y + H * 0.5)

    End With

    POLYGONComputeFaceNormals NBodies
    ComputeMass NBodies, DefDensity
    Body(NBodies).Pos = Body(NBodies).COM

    If Rnd < 0.5 Then Chamfer NBodies, 4 + Rnd * 5

End Sub


Public Sub CreateRegularPoly(Pos As tVec2, Rw As Single, Rh As Single, N As Long, Flat As Long, Density As Single)
    Dim I      As Long
    Dim A      As Single

    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon

        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = 0
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .U = SetOrient(.orient)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0

        .Nvertex = N

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        If Flat Then A = 0.5 * PI2 / N

        For I = 1 To .Nvertex
            '        For I = .Nvertex To 1 Step -1
            .Vertex(I) = Vec2ADD(Pos, _
                                 Vec2((Rw) * Cos(A + PI2 * (I - 1) / .Nvertex), _
                                      (Rh) * Sin(A + PI2 * (I - 1) / .Nvertex)))
        Next

    End With


    POLYGONComputeFaceNormals NBodies
    ComputeMass NBodies, Density
    Body(NBodies).Pos = Body(NBodies).COM

    '   If Rnd < 0.5 Then Chamfer NBodies, 9
    Chamfer NBodies, 8 + Rnd * 8, , 4

End Sub


'  // The extreme point along a direction within a polygon
'  Vec2 GetSupport( const Vec2& dir )
'  {
'    real bestProjection = -MAX_VALUE;
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
    Dim bestProjection As Single
    Dim bestVertex As tVec2
    Dim V      As tVec2
    Dim I      As Long
    Dim projection As Single

    bestProjection = -MAX_VALUE
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


'''
'''    /**
'''     * Chamfers a set of vertices by giving them rounded corners, returns a new set of vertices.
'''     * The radius parameter is a single number or an array to specify the radius for each vertex.
'''     * @method chamfer
'''     * @param {vertices} vertices
'''     * @param {number[]} radius
'''     * @param {number} quality
'''     * @param {number} qualityMin
'''     * @param {number} qualityMax
'''     */
'''    Vertices.chamfer = function(vertices, radius, quality, qualityMin, qualityMax) {
'''        radius = radius || [8];
'''
'''        if (!radius.length)
'''            radius = [radius];
'''
'''        // quality defaults to -1, which is auto
'''        quality = (typeof quality !== 'undefined') ? quality : -1;
'''        qualityMin = qualityMin || 2;
'''        qualityMax = qualityMax || 14;
'''
'''        var newVertices = [];
'''
'''        for (var i = 0; i < vertices.length; i++) {
'''            var prevVertex = vertices[i - 1 >= 0 ? i - 1 : vertices.length - 1],
'''                vertex = vertices[i],
'''                nextVertex = vertices[(i + 1) % vertices.length],
'''                currentRadius = radius[i < radius.length ? i : radius.length - 1];
'''
'''            if (currentRadius === 0) {
'''                newVertices.push(vertex);
'''                continue;
'''            }
'''
'''            var prevNormal = Vector.normalise({
'''                x: vertex.y - prevVertex.y,
'''y:                 prevVertex.x -Vertex.x
'''            });
'''
'''            var nextNormal = Vector.normalise({
'''                x: nextVertex.y - vertex.y,
'''y:                 Vertex.x -nextVertex.x
'''            });
'''
'''            var diagonalRadius = Math.sqrt(2 * Math.pow(currentRadius, 2)),
'''                radiusVector = Vector.mult(Common.clone(prevNormal), currentRadius),
'''                midNormal = Vector.normalise(Vector.mult(Vector.add(prevNormal, nextNormal), 0.5)),
'''                scaledVertex = Vector.sub(vertex, Vector.mult(midNormal, diagonalRadius));
'''
'''            var precision = quality;
'''
'''            if (quality === -1) {
'''                // automatically decide precision
'''                precision = Math.pow(currentRadius, 0.32) * 1.75;
'''            }
'''
'''            precision = Common.clamp(precision, qualityMin, qualityMax);
'''
'''            // use an even value for precision, more likely to reduce axes by using symmetry
'''            if (precision % 2 === 1)
'''                precision += 1;
'''
'''            var alpha = Math.acos(Vector.dot(prevNormal, nextNormal)),
'''                theta = alpha / precision;
'''
'''            for (var j = 0; j < precision; j++) {
'''                newVertices.push(Vector.add(Vector.rotate(radiusVector, theta * j), scaledVertex));
'''            }
'''        }
'''
'''        return newVertices;
'''    };

'
'    /**
'     * Chamfers a set of vertices by giving them rounded corners, returns a new set of vertices.
'     * The radius parameter is a single number or an array to specify the radius for each vertex.
'     * @method chamfer
'     * @param {vertices} vertices
'     * @param {number[]} radius
'     * @param {number} quality
'     * @param {number} qualityMin
'     * @param {number} qualityMax
'     */
Public Sub Chamfer(wB As Long, Radius As Single, _
                   Optional Quality As Single = -1, Optional QualityMin As Single = 2, Optional QualityMax As Single = 14)
'        radius = radius || [8];
'
'        if (!radius.length)
'            radius = [radius];

'// quality defaults to -1, which is auto
'quality = (typeof quality !== 'undefined') ? quality : -1;
'qualityMin = qualityMin || 2;
'qualityMax = qualityMax || 14;
    Dim prevVertex As tVec2
    Dim nextVertex As tVec2
    Dim Vertex As tVec2
    Dim newVertex() As tVec2
    Dim newVertexN As Long

    Dim I      As Long

    Dim Precision As Single
    Dim currentRadius As Single
    Dim prevNormal As tVec2
    Dim nextNormal As tVec2
    Dim diagonalRadius As Single
    Dim radiusVector As tVec2
    Dim midNormal As tVec2
    Dim scaledVertex As tVec2
    Dim alpha  As Single
    Dim theta  As Single
    Dim J      As Long


    With Body(wB)





        For I = 1 To .Nvertex
            Vertex = .Vertex(I)
            If I = .Nvertex Then nextVertex = .Vertex(1) Else: nextVertex = .Vertex(I + 1)
            If Vec2Length(Vec2SUB(Vertex, nextVertex)) <= Radius * 2 Then Exit Sub
        Next

        For I = 1 To .Nvertex
            Vertex = .Vertex(I)
            If I = 1 Then prevVertex = .Vertex(.Nvertex) Else: prevVertex = .Vertex(I - 1)
            If I = .Nvertex Then nextVertex = .Vertex(1) Else: nextVertex = .Vertex(I + 1)




            '        for (var i = 0; i < vertices.length; i++) {
            '            var prevVertex = vertices[i - 1 >= 0 ? i - 1 : vertices.length - 1],
            '                vertex = vertices[i],
            '                nextVertex = vertices[(i + 1) % vertices.length],
            '                currentRadius = radius[i < radius.length ? i : radius.length - 1];
            currentRadius = Radius

            '            if (currentRadius === 0) {
            '                newVertices.push(vertex);
            '                continue;
            '            }

            '            var prevNormal = Vector.normalise({
            '                x: vertex.y - prevVertex.y,
            'y:                 prevVertex.x -Vertex.x
            '            });
            '
            '            var nextNormal = Vector.normalise({
            '                x: nextVertex.y - vertex.y,
            'y:                 Vertex.x -nextVertex.x
            '            });
            prevNormal = Vec2Normalize(Vec2(Vertex.y - prevVertex.y, prevVertex.x - Vertex.x))
            nextNormal = Vec2Normalize(Vec2(nextVertex.y - Vertex.y, Vertex.x - nextVertex.x))

            '            var diagonalRadius = Math.sqrt(2 * Math.pow(currentRadius, 2)),
            '                radiusVector = Vector.mult(Common.clone(prevNormal), currentRadius),
            '                midNormal = Vector.normalise(Vector.mult(Vector.add(prevNormal, nextNormal), 0.5)),
            '                scaledVertex = Vector.sub(vertex, Vector.mult(midNormal, diagonalRadius));
            diagonalRadius = Sqr(2 * (currentRadius ^ 2))
            radiusVector = Vec2MUL(prevNormal, currentRadius)
            midNormal = Vec2Normalize(Vec2MUL(Vec2ADD(prevNormal, nextNormal), 0.5))
            scaledVertex = Vec2SUB(Vertex, Vec2MUL(midNormal, diagonalRadius))

            Precision = Quality
            '            var precision = quality;

            '            if (quality === -1) {
            '                // automatically decide precision
            '                precision = Math.pow(currentRadius, 0.32) * 1.75;
            '            }
            If Quality = -1 Then Precision = (currentRadius ^ 0.32) * 1.75



            '            precision = Common.clamp(precision, qualityMin, qualityMax);
            Precision = Clamp(Precision, QualityMin, QualityMax)

            '            // use an even value for precision, more likely to reduce axes by using symmetry
            '            if (precision % 2 === 1)
            '                precision += 1;
            Precision = Round(Precision)
            If Precision Mod 2 = 0 Then Precision = Precision + 1

            '            var alpha = Math.acos(Vector.dot(prevNormal, nextNormal)),
            '                theta = alpha / precision;
            alpha = ACOS(Vec2DOT(prevNormal, nextNormal))
            theta = alpha / Precision

            '  for (var j = 0; j < precision; j++) {
            '  newVertices.push(Vector.add(Vector.rotate(radiusVector, theta * j), scaledVertex));
            '  }

            For J = 0 To Precision
                newVertexN = newVertexN + 1
                ReDim Preserve newVertex(newVertexN)
                newVertex(newVertexN) = Vec2ADD(Vec2Rotate(radiusVector, theta * J), scaledVertex)
            Next


        Next


        .Vertex = newVertex
        .Nvertex = newVertexN
        ReDim .tVertex(newVertexN)

    End With

    POLYGONComputeFaceNormals wB
    '     ComputeMass wB, DefDensity  '''' Causes error to much inertia!!!
    '    Body(wB).Pos = Body(NBodies).COM


End Sub

