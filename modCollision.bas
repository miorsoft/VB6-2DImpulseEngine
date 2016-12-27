Attribute VB_Name = "modCollision"
Option Explicit

Public Enum eCollision
    eCollisionCircleCircle = 0    '&H0
    eCollisionCirclePolygon = 1    '&H1
    eCollisionPolygonCircle = 2    '&H10
    eCollisionPolygonPolygon = 3    '&H11
End Enum


Public Type tManifold

    ContactType As eCollision

    bodyA   As Long
    bodyB   As Long

    penetration As Double
    normal  As tVec2

    contactsPTS(1 To 2) As tVec2
    contactCount As Long
    MAXcontactCount As Long

    e       As Double
    DF      As Double
    SF      As Double

End Type


Public Contacts() As tManifold
Public NofContactMainFolds As Long
Public MAXNofContactMainFolds As Long



Public Sub ContactsINIT(wC As Long)
    Dim I   As Long
    Dim rA  As tVec2
    Dim rB  As tVec2
    Dim rV  As tVec2
    Dim A   As Long
    Dim B   As Long


    With Contacts(wC)
        A = .bodyA
        B = .bodyB


        '   // Calculate average restitution
        '   // e = std::min( A->restitution, B->restitution );
        '   e = StrictMath.min( A.restitution, B.restitution );
        .e = Min(Body(A).Restitution, Body(B).Restitution)

        '// Calculate static and dynamic friction
        '// sf = std::sqrt( A->staticFriction * A->staticFriction );
        '// df = std::sqrt( A->dynamicFriction * A->dynamicFriction );
        'sf = (float)StrictMath.sqrt( A.staticFriction * A.staticFriction );
        'df = (float)StrictMath.sqrt( A.dynamicFriction * A.dynamicFriction );

        .SF = Sqr(Body(A).staticFriction * Body(B).staticFriction)
        .DF = Sqr(Body(A).dynamicFriction * Body(B).dynamicFriction)


        For I = 1 To .contactCount

            '// Calculate radii from COM to contact
            '// Vec2 ra = contacts[i] - A->position;
            '// Vec2 rb = contacts[i] - B->position;
            'Vec2 ra = contacts[i].sub( A.position );
            'Vec2 rb = contacts[i].sub( B.position );

            rA = Vec2SUB(.contactsPTS(I), Body(A).Pos)
            rB = Vec2SUB(.contactsPTS(I), Body(B).Pos)

            ' // Vec2 rv = B->velocity + Cross( B->angularVelocity, rb ) -
            ' // A->velocity - Cross( A->angularVelocity, ra );
            ' Vec2 rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );

            rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))
            rV = Vec2SUB(rV, Body(A).VEL)
            rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))



            '            // Determine if we should perform a resting collision or not
            '            // The idea is if the only thing moving this object is gravity,
            '            // then the collision should be performed without any restitution
            '            // if(rv.LenSqr( ) < (dt * gravity).LenSqr( ) + EPSILON)
            '            if (rv.lengthSq() < ImpulseMath.RESTING)
            '            {
            '            e = 0.0f;
            '            }
            If Vec2LengthSq(rV) < RESTING Then .e = 0
        Next

    End With


End Sub
Public Sub infiniteMassCorrection(wC As Long)
    Body(Contacts(wC).bodyA).VEL = Vec2(0, 0)
    Body(Contacts(wC).bodyB).VEL = Vec2(0, 0)
End Sub
Public Sub contactsPositionalCorrection(wC As Long)
    Dim correction As Double

    '// const real k_slop = 0.05f; // Penetration allowance
    '// const real percent = 0.4f; // Penetration percentage to correct
    '// Vec2 correction = (std::max( penetration - k_slop, 0.0f ) / (A->im +
    '// B->im)) * normal * percent;
    '// A->position -= correction * A->im;
    '// B->position += correction * B->im;

    'float correction = StrictMath.max( penetration - ImpulseMath.PENETRATION_ALLOWANCE, 0.0f ) / (A.invMass + B.invMass) * ImpulseMath.PENETRATION_CORRETION;

    'A.position.addsi( normal, -A.invMass * correction );
    'B.position.addsi( normal, B.invMass * correction );


    With Contacts(wC)
        correction = (Max(.penetration - PENETRATION_ALLOWANCE, 0) / (Body(.bodyA).invMass + Body(.bodyB).invMass)) _
                     * PENETRATION_CORRETION

        Body(.bodyA).Pos = Vec2ADDS(Body(.bodyA).Pos, .normal, -Body(.bodyA).invMass * correction)
        Body(.bodyB).Pos = Vec2ADDS(Body(.bodyB).Pos, .normal, Body(.bodyB).invMass * correction)

    End With

End Sub
Public Sub contactsApplyImpulse(wC As Long)
    Dim A   As Long
    Dim B   As Long

    Dim I   As Long

    Dim rA  As tVec2
    Dim rB  As tVec2
    Dim rV  As tVec2
    Dim contactVel As Double


    Dim rACrossN As Double
    Dim rBCrossN As Double
    Dim InvMassSUM As Double
    Dim J   As Double
    Dim Jt  As Double


    Dim impulse As tVec2
    Dim T   As tVec2

    Dim tangentImpulse As tVec2


    With Contacts(wC)
        A = .bodyA
        B = .bodyB

        '// Early out and positional correct if both objects have infinite mass
        '// if(Equal( A->im + B->im, 0 ))
        'if (ImpulseMath.equal( A.invMass + B.invMass, 0 ))
        '{
        '    infiniteMassCorrection();
        '    return;
        '}


        If Equal(Body(A).invMass + Body(B).invMass, 0) Then
            infiniteMassCorrection wC
        Else

            For I = 1 To .contactCount

                '   // Calculate radii from COM to contact
                '   // Vec2 ra = contacts[i] - A->position;
                '   // Vec2 rb = contacts[i] - B->position;
                '   Vec2 ra = contacts[i].sub( A.position );
                '   Vec2 rb = contacts[i].sub( B.position );
                rA = Vec2SUB(.contactsPTS(I), Body(A).Pos)
                rB = Vec2SUB(.contactsPTS(I), Body(B).Pos)

                ' // Relative velocity
                ' // Vec2 rv = B->velocity + Cross( B->angularVelocity, rb ) -
                ' // A->velocity - Cross( A->angularVelocity, ra );
                ' Vec2 rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );

                rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))
                rV = Vec2SUB(rV, Body(A).VEL)
                rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))


                '// Relative velocity along the normal
                '// real contactVel = Dot( rv, normal );
                'float contactVel = Vec2.dot( rv, normal );
                contactVel = Vec2DOT(rV, .normal)


                '// Do not resolve if velocities are separating
                'if (contactVel > 0)
                '{
                'return;
                '}
                If contactVel < 0 Then
                    '// real raCrossN = Cross( ra, normal );
                    '// real rbCrossN = Cross( rb, normal );
                    '// real invMassSum = A->im + B->im + FASTsqr( raCrossN ) * A->iI + FASTsqr(
                    '// rbCrossN ) * B->iI;
                    'float raCrossN = Vec2.cross( ra, normal );
                    'float rbCrossN = Vec2.cross( rb, normal );
                    'float invMassSum = A.invMass + B.invMass + (raCrossN * raCrossN) * A.invInertia + (rbCrossN * rbCrossN) * B.invInertia;
                    rACrossN = Vec2CROSS(rA, .normal)
                    rBCrossN = Vec2CROSS(rB, .normal)
                    InvMassSUM = Body(A).invMass + Body(B).invMass + (rACrossN * rACrossN) * Body(A).invInertia + (rBCrossN * rBCrossN) * Body(B).invInertia



                    ' // Calculate impulse scalar
                    ' float j = -(1.0f + e) * contactVel;
                    ' j /= invMassSum;
                    ' j /= contactCount;
                    J = -(1 + .e) * contactVel
                    J = J / InvMassSUM
                    J = J / .contactCount

                    ' // Apply impulse
                    ' Vec2 impulse = normal.mul( j );
                    ' A.applyImpulse( impulse.neg(), ra );
                    ' B.applyImpulse( impulse, rb );
                    impulse = Vec2MUL(.normal, J)
                    BodyApplyImpulse A, Vec2Negative(impulse), rA
                    BodyApplyImpulse B, impulse, rB


                    '          // Friction impulse
                    '          // rv = B->velocity + Cross( B->angularVelocity, rb ) -
                    '          // A->velocity - Cross( A->angularVelocity, ra );
                    '          rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );
                    rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))

                    rV = Vec2SUB(rV, Body(A).VEL)
                    rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))

                    '// Vec2 t = rv - (normal * Dot( rv, normal ));
                    '// t.Normalize( );
                    'Vec2 t = new Vec2( rv );
                    't.addsi( normal, -Vec2.dot( rv, normal ) );
                    't.normalize();
                    T = Vec2SUB(rV, Vec2MUL(.normal, Vec2DOT(rV, .normal)))
                    T = Vec2Normalize(T)

                    '// j tangent magnitude
                    'float jt = -Vec2.dot( rv, t );
                    'jt /= invMassSum;
                    'jt /= contactCount;
                    Jt = -Vec2DOT(rV, T)
                    Jt = Jt / InvMassSUM
                    Jt = Jt / .contactCount


                    '// Don    't apply tiny friction impulses
                    'if (ImpulseMath.equal( jt, 0.0f ))
                    '{
                    'return;
                    '}

                    If Not (Equal(Jt, 0)) Then

                        '// Coulumb    's law
                        'Vec2 tangentImpulse;
                        '// if(std::abs( jt ) < j * sf)
                        'if (StrictMath.abs( jt ) < j * sf)
                        '{
                        '// tangentImpulse = t * jt;
                        'tangentImpulse = t.mul( jt );
                        '}
                        'Else
                        '{
                        '// tangentImpulse = t * -j * df;
                        'tangentImpulse = t.mul( j ).muli( -df );
                        '}
                        If Abs(Jt) < J * .SF Then
                            tangentImpulse = Vec2MUL(T, Jt)
                        Else
                            tangentImpulse = Vec2MUL(T, -J * .DF)
                        End If

                        '// Apply friction impulse
                        '// A->ApplyImpulse( -tangentImpulse, ra );
                        '// B->ApplyImpulse( tangentImpulse, rb );
                        'A.applyImpulse( tangentImpulse.neg(), ra );
                        'B.applyImpulse( tangentImpulse, rb );
                        BodyApplyImpulse A, Vec2Negative(tangentImpulse), rA
                        BodyApplyImpulse B, tangentImpulse, rB
                    Else
                        Exit Sub    'Don't apply tiny friction impulses
                    End If
                Else
                    Exit Sub    ' Do not resolve if velocities are separating
                End If
            Next

        End If


    End With

End Sub

'Public Function CollisionCirclePolygon(wbA As Long, wBB As Long) As tManifold
'
'    Dim A      As tBody
'    Dim B      As tBody
'
'    Dim normal As tVec2
'    Dim DistDist As Double
'    Dim radius As Double
'    Dim distance As Double
'
'    Dim center As tVec2
'
'
'    Dim separation As Double
'    Dim faceNormal As Long
'
'    Dim I      As Long
'    Dim I2     As Long
'
'    Dim S      As Double
'
'    Dim v1     As tVec2
'    Dim v2     As tVec2
'
'    Dim Dot1   As Double
'    Dim Dot2   As Double
'
'    Dim N      As tVec2
'
'
'
'    CollisionCirclePolygon.bodyA = wbA
'    CollisionCirclePolygon.bodyB = wBB
'    A = Body(wbA)
'    B = Body(wBB)
'
'    '*******************************************************************************
'    '*******************************************************************************
'    CollisionCirclePolygon.contactCount = 0
'
'    '// Transform circle center to Polygon model space
'    '// Vec2 center = a->position;
'    '// center = B->u.Transpose( ) * (center - b->position);
'    'Vec2 center = B.u.transpose().muli( a.position.sub( b.position ) );
'    center = matMULv(matTranspose(B.U), Vec2SUB(A.Pos, B.Pos))
'
'
'    '        // Find edge with minimum penetration
'    '        // Exact concept as using support points in Polygon vs Polygon
'    '        float separation = -Float.MAX_VALUE;
'    '        int faceNormal = 0;
'    separation = -MAX_VALUE
'    faceNormal = 0
'
'
'
'
'    For I = 1 To B.Nvertex
'
'        '// real s = Dot( B->m_normals[i], center - B->m_vertices[i] );
'        'float s = Vec2.dot( B.normals[i], center.sub( B.vertices[i] ) );
'
'        S = Vec2DOT(B.normals(I), Vec2SUB(center, B.Vertex(I)))
'
'        If (S < A.radius) Then
'            If (S > separation) Then
'                separation = S
'                faceNormal = I
'            End If
'        End If
'
'    Next
'
'
'
'    '// Grab face's vertices
'    'Vec2 v1 = B.vertices[faceNormal];
'    'int i2 = faceNormal + 1 < B.vertexCount ? faceNormal + 1 : 0;
'    'Vec2 v2 = B.vertices[i2];
'
'    v1 = B.Vertex(faceNormal)
'    I2 = faceNormal + 1
'    If I2 > B.Nvertex Then I2 = 1
'    v2 = B.Vertex(I2)
'
'
'
'    '// Check to see if center is within polygon
'    'if (separation < ImpulseMath.EPSILON)
'    '{
'    '    // m->contact_count = 1;
'    '    // m->normal = -(B->u * B->m_normals[faceNormal]);
'    '    // m->contacts[0] = m->normal * A->radius + a->position;
'    '    // m->penetration = A->radius;
'
'    '    m.contactCount = 1;
'    '    B.u.mul( B.normals[faceNormal], m.normal ).negi();
'    '    m.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
'    '    m.penetration = A.radius;
'    '    return;
'    '}
'    If separation < EPSILON Then
'        CollisionCirclePolygon.contactCount = 1
'        ReDim CollisionCirclePolygon.contactsPTS(1)
'        CollisionCirclePolygon.normal = Vec2Negative(matMULv(B.U, B.normals(faceNormal)))
'        CollisionCirclePolygon.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionCirclePolygon.normal, A.radius), A.Pos)
'        CollisionCirclePolygon.penetration = A.radius
'    Else    'because of return
'
'
'        '// Determine which voronoi region of the edge center of circle lies within
'        '// real dot1 = Dot( center - v1, v2 - v1 );
'        '// real dot2 = Dot( center - v2, v1 - v2 );
'        '// m->penetration = A->radius - separation;
'        'float dot1 = Vec2.dot( center.sub( v1 ), v2.sub( v1 ) );
'        'float dot2 = Vec2.dot( center.sub( v2 ), v1.sub( v2 ) );
'        'm.penetration = A.radius - separation;
'
'        Dot1 = Vec2DOT(Vec2SUB(center, v1), Vec2SUB(v2, v1))
'        Dot2 = Vec2DOT(Vec2SUB(center, v2), Vec2SUB(v1, v2))
'        CollisionCirclePolygon.penetration = A.radius - separation
'
'
'        '        // Closest to v1
'        '        if (dot1 <= 0.0f)
'        '        {
'        '            if (Vec2.distanceSq( center, v1 ) > A.radius * A.radius)
'        '            {
'        '                return;
'        '            }
'        '
'        '            // m->contact_count = 1;
'        '            // Vec2 n = v1 - center;
'        '            // n = B->u * n;
'        '            // n.Normalize( );
'        '            // m->normal = n;
'        '            // v1 = B->u * v1 + b->position;
'        '            // m->contacts[0] = v1;
'        '
'        '            m.contactCount = 1;
'        '            B.u.muli( m.normal.set( v1 ).subi( center ) ).normalize();
'        '            B.u.mul( v1, m.contacts[0] ).addi( b.position );
'        '        }
'        If Dot1 <= 0 Then
'
'            If Vec2DISTANCEsq(center, v1) < A.radius * A.radius Then
'
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                N = Vec2SUB(v1, center)
'                CollisionCirclePolygon.normal = Vec2Normalize(matMULv(B.U, N))
'                v1 = Vec2ADD(matMULv(B.U, v1), B.Pos)
'                CollisionCirclePolygon.contactsPTS(1) = v1
'
'            End If
'            '        else if (dot2 <= 0.0f)
'        ElseIf Dot2 <= 0 Then
'
'            '                {
'            '                if (Vec2.distanceSq( center, v2 ) > A.radius * A.radius)
'            '                {
'            '                return;
'            '                }
'            '
'            '                // m->contact_count = 1;
'            '                // Vec2 n = v2 - center;
'            '                // v2 = B->u * v2 + b->position;
'            '                // m->contacts[0] = v2;
'            '                // n = B->u * n;
'            '                // n.Normalize( );
'            '                // m->normal = n;
'            '
'            '                m.contactCount = 1;
'            '                B.u.muli( m.normal.set( v2 ).subi( center ) ).normalize();
'            '                B.u.mul( v2, m.contacts[0] ).addi( b.position );
'            '                }
'            If Vec2DISTANCEsq(center, v2) < A.radius * A.radius Then
'
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                N = Vec2SUB(v2, center)
'                v2 = Vec2ADD(matMULv(B.U, v2), B.Pos)
'                CollisionCirclePolygon.contactsPTS(1) = v2
'                CollisionCirclePolygon.normal = Vec2Normalize(matMULv(B.U, N))
'            End If
'
'
'
'            '// Closest to face
'        Else
'            '                Vec2 n = B.normals[faceNormal];
'            '
'            '                if (Vec2.dot( center.sub( v1 ), n ) > A.radius)
'            '                {
'            '                return;
'            '                }
'            '
'            '                // n = B->u * n;
'            '                // m->normal = -n;
'            '                // m->contacts[0] = m->normal * A->radius + a->position;
'            '                // m->contact_count = 1;
'            '
'            '                m.contactCount = 1;
'            '                B.u.mul( n, m.normal ).negi();
'            '                m.contacts[0].set( a.position ).addsi( m.normal, A.radius );
'
'
'            N = B.normals(faceNormal)
'            If Vec2DOT(Vec2SUB(center, v1), N) < A.radius Then
'
'
'                N = matMULv(B.U, N)
'                CollisionCirclePolygon.normal = Vec2Negative(N)
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                CollisionCirclePolygon.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionCirclePolygon.normal, A.radius), A.Pos)
'
'            End If
'
'        End If
'
'    End If
'
'End Function


Public Function CollisionSOLVE(wbA As Long, wbB As Long) As tManifold
    Dim A   As tBody
    Dim B   As tBody

    Dim normal As tVec2
    Dim DistDist As Double
    Dim radius As Double
    Dim distance As Double

    Dim Center As tVec2


    Dim separation As Double
    Dim faceNormal As Long

    Dim I   As Long
    Dim I2  As Long

    Dim S   As Double

    Dim v1  As tVec2
    Dim v2  As tVec2

    Dim Dot1 As Double
    Dim Dot2 As Double

    Dim N   As tVec2

    Dim ContactType As Long

    Dim faceA As Long    '---------------polypoly
    Dim faceB As Long
    Dim penetrationA As Double
    Dim penetrationB As Double
    Dim flip As Boolean
    Dim RefPoly As tBody
    Dim IncPoly As tBody
    Dim incidentFace(0 To 1) As tVec2
    Dim referenceIndex As Long
    Dim sidePlaneNormal As tVec2
    Dim refFaceNormal As tVec2
    Dim refC As Double
    Dim posSide As Double
    Dim negSide As Double
    Dim cp  As Long


    '---------------------------------------------------------------


    ContactType = Body(wbA).myType * 2 + Body(wbB).myType

    Select Case ContactType

        Case 0    'CircleCircle

            '*******************************************************************************
            '*******************************************************************************
            CollisionSOLVE.bodyA = wbA
            CollisionSOLVE.bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)

            '        // Calculate translational vector, which is normal
            '        // Vec2 normal = b->position - a->position;
            '        Vec2 normal = b.position.sub( a.position );
            normal = Vec2SUB(B.Pos, A.Pos)
            '        // real DistDist = normal.LenSqr( );
            '        // real radius = A->radius + B->radius;
            '        float DistDist = normal.lengthSq();
            '        float radius = A.radius + B.radius;
            DistDist = Vec2LengthSq(normal)
            radius = A.radius + B.radius

            '        // Not in contact
            If (DistDist >= radius * radius) Then
                CollisionSOLVE.contactCount = 0
            Else

                distance = Sqr(DistDist)

                CollisionSOLVE.contactCount = 1

                If (distance = 0) Then
                    ' // m->penetration = A->radius;
                    ' // m->normal = Vec2( 1, 0 );
                    ' // m->contacts [0] = a->position;
                    ' m.penetration = A.radius;
                    ' m.normal.set( 1.0f, 0.0f );
                    ' m.contacts[0].set( a.position );
                    CollisionSOLVE.penetration = A.radius
                    CollisionSOLVE.normal.X = 1
                    CollisionSOLVE.normal.y = 0
                    'ReDim CollisionSOLVE.contactsPTS(1)
                    CollisionSOLVE.contactsPTS(1) = A.Pos

                Else
                    '// m->penetration = radius - distance;
                    '// m->normal = normal / distance; // Faster than using Normalized since
                    '// we already performed sqrt
                    '// m->contacts[0] = m->normal * A->radius + a->position;
                    'm.penetration = radius - distance;
                    'm.normal.set( normal ).divi( distance );
                    'm.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
                    CollisionSOLVE.penetration = radius - distance
                    CollisionSOLVE.normal = Vec2MUL(normal, 1 / distance)
                    'ReDim CollisionSOLVE.contactsPTS(1)
                    CollisionSOLVE.contactsPTS(1) = Vec2ADD(A.Pos, Vec2MUL(CollisionSOLVE.normal, A.radius))
                End If

            End If


        Case 1    ' CirclePolygon
            CollisionSOLVE.bodyA = wbA
            CollisionSOLVE.bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)
            GoSub LABELCirclePolygon


        Case 2    'PolygonCircle

            CollisionSOLVE.bodyA = wbB
            CollisionSOLVE.bodyB = wbA
            A = Body(wbB)
            B = Body(wbA)
            GoSub LABELCirclePolygon
            'CollisionSOLVE.normal = Vec2Negative(CollisionSOLVE.normal)


        Case 3    ' PolygonPolygon
            '*******************************************************************************
            '*******************************************************************************
            CollisionSOLVE.bodyA = wbA
            CollisionSOLVE.bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)
            GoSub LabelPolygonPolygon

    End Select




    Exit Function

LABELCirclePolygon:

    '*******************************************************************************
    '*******************************************************************************
    CollisionSOLVE.contactCount = 0

    '// Transform circle center to Polygon model space
    '// Vec2 center = a->position;
    '// center = B->u.Transpose( ) * (center - b->position);
    'Vec2 center = B.u.transpose().muli( a.position.sub( b.position ) );

    Center = matMULv(matTranspose(B.U), Vec2SUB(A.Pos, B.Pos))



    '        // Find edge with minimum penetration
    '        // Exact concept as using support points in Polygon vs Polygon
    '        float separation = -Float.MAX_VALUE;
    '        int faceNormal = 0;
    separation = -MAX_VALUE
    faceNormal = 0

    For I = 1 To B.Nvertex

        '// real s = Dot( B->m_normals[i], center - B->m_vertices[i] );
        'float s = Vec2.dot( B.normals[i], center.sub( B.vertices[i] ) );

        S = Vec2DOT(B.normals(I), Vec2SUB(Center, B.Vertex(I)))
        '    if(s > A->radius)
        '      return;
        '    if(s > separation)
        '    {
        '      separation = s;
        '      faceNormal = i;
        '    }
        If (S <= A.radius) Then
            If (S > separation) Then
                separation = S
                faceNormal = I
            End If
        Else

            Return

        End If
    Next




    '// Grab face's vertices
    'Vec2 v1 = B.vertices[faceNormal];
    'int i2 = faceNormal + 1 < B.vertexCount ? faceNormal + 1 : 0;
    'Vec2 v2 = B.vertices[i2];

    v1 = B.Vertex(faceNormal)
    I2 = faceNormal + 1: If I2 > B.Nvertex Then I2 = 1
    v2 = B.Vertex(I2)



    '// Check to see if center is within polygon
    'if (separation < ImpulseMath.EPSILON)
    '{
    '    // m->contact_count = 1;
    '    // m->normal = -(B->u * B->m_normals[faceNormal]);
    '    // m->contacts[0] = m->normal * A->radius + a->position;
    '    // m->penetration = A->radius;

    '    m.contactCount = 1;
    '    B.u.mul( B.normals[faceNormal], m.normal ).negi();
    '    m.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
    '    m.penetration = A.radius;
    '    return;
    '}
    If separation < EPSILON Then
        CollisionSOLVE.contactCount = 1
        'ReDim CollisionSOLVE.contactsPTS(1)
        CollisionSOLVE.normal = Vec2Negative(matMULv(B.U, B.normals(faceNormal)))
        CollisionSOLVE.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionSOLVE.normal, A.radius), A.Pos)
        CollisionSOLVE.penetration = A.radius
        Return
        'Else    'because of return
    End If

    '// Determine which voronoi region of the edge center of circle lies within
    '// real dot1 = Dot( center - v1, v2 - v1 );
    '// real dot2 = Dot( center - v2, v1 - v2 );
    '// m->penetration = A->radius - separation;
    'float dot1 = Vec2.dot( center.sub( v1 ), v2.sub( v1 ) );
    'float dot2 = Vec2.dot( center.sub( v2 ), v1.sub( v2 ) );
    'm.penetration = A.radius - separation;

    Dot1 = Vec2DOT(Vec2SUB(Center, v1), Vec2SUB(v2, v1))
    Dot2 = Vec2DOT(Vec2SUB(Center, v2), Vec2SUB(v1, v2))
    CollisionSOLVE.penetration = A.radius - separation


    '        // Closest to v1
    '        if (dot1 <= 0.0f)
    '        {
    '            if (Vec2.distanceSq( center, v1 ) > A.radius * A.radius)
    '            {
    '                return;
    '            }
    '
    '            // m->contact_count = 1;
    '            // Vec2 n = v1 - center;
    '            // n = B->u * n;
    '            // n.Normalize( );
    '            // m->normal = n;
    '            // v1 = B->u * v1 + b->position;
    '            // m->contacts[0] = v1;
    '
    '            m.contactCount = 1;
    '            B.u.muli( m.normal.set( v1 ).subi( center ) ).normalize();
    '            B.u.mul( v1, m.contacts[0] ).addi( b.position );
    '        }
    If Dot1 <= 0 Then

        If Vec2DISTANCEsq(Center, v1) < A.radius * A.radius Then

            CollisionSOLVE.contactCount = 1
            'ReDim CollisionSOLVE.contactsPTS(1)
            N = Vec2SUB(v1, Center)
            CollisionSOLVE.normal = Vec2Normalize(matMULv(B.U, N))
            v1 = Vec2ADD(matMULv(B.U, v1), B.Pos)
            CollisionSOLVE.contactsPTS(1) = v1
        Else
            Return
        End If
        '        else if (dot2 <= 0.0f)
    ElseIf Dot2 <= 0 Then

        '                {
        '                if (Vec2.distanceSq( center, v2 ) > A.radius * A.radius)
        '                {
        '                return;
        '                }
        '
        '                // m->contact_count = 1;
        '                // Vec2 n = v2 - center;
        '                // v2 = B->u * v2 + b->position;
        '                // m->contacts[0] = v2;
        '                // n = B->u * n;
        '                // n.Normalize( );
        '                // m->normal = n;
        '
        '                m.contactCount = 1;
        '                B.u.muli( m.normal.set( v2 ).subi( center ) ).normalize();
        '                B.u.mul( v2, m.contacts[0] ).addi( b.position );
        '                }
        If Vec2DISTANCEsq(Center, v2) < A.radius * A.radius Then

            CollisionSOLVE.contactCount = 1
            ' ReDim CollisionSOLVE.contactsPTS(1)
            N = Vec2SUB(v2, Center)
            v2 = Vec2ADD(matMULv(B.U, v2), B.Pos)
            CollisionSOLVE.contactsPTS(1) = v2
            CollisionSOLVE.normal = Vec2Normalize(matMULv(B.U, N))
        Else
            Return
        End If



        '// Closest to face
    Else

        '                Vec2 n = B.normals[faceNormal];
        '
        '                if (Vec2.dot( center.sub( v1 ), n ) > A.radius)
        '                {
        '                return;
        '                }
        '
        '                // n = B->u * n;
        '                // m->normal = -n;
        '                // m->contacts[0] = m->normal * A->radius + a->position;
        '                // m->contact_count = 1;
        '
        '                m.contactCount = 1;
        '                B.u.mul( n, m.normal ).negi();
        '                m.contacts[0].set( a.position ).addsi( m.normal, A.radius );


        N = B.normals(faceNormal)
        If Vec2DOT(Vec2SUB(Center, v1), N) < A.radius Then

            N = matMULv(B.U, N)
            CollisionSOLVE.normal = Vec2Negative(N)
            CollisionSOLVE.contactCount = 1
            ' ReDim CollisionSOLVE.contactsPTS(1)
            CollisionSOLVE.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionSOLVE.normal, A.radius), A.Pos)
        Else
            Return

        End If

    End If

    'End If ''''(Return up)
    Return




LabelPolygonPolygon:
    '*******************************************************************************
    '*******************************************************************************
    CollisionSOLVE.contactCount = 0


    'penetrationA = FindAxisLeastPenetration( &faceA, A, B );
    penetrationA = FindAxisLeastPenetration(faceA, A, B)
    If penetrationA >= 0 Then Return
    penetrationB = FindAxisLeastPenetration(faceB, B, A)
    If penetrationB >= 0 Then Return

    '// Determine which shape contains reference face
    If (BiasGreaterThan(penetrationA, penetrationB)) Then
        RefPoly = A
        IncPoly = B
        referenceIndex = faceA
        flip = False
    Else
        RefPoly = B
        IncPoly = A
        referenceIndex = faceB
        flip = True
    End If

    '  // World space incident face


    FindIncidentFace incidentFace(), RefPoly, IncPoly, referenceIndex

    ''  //        y
    ''  //        ^  ->n       ^
    ''  //      +---c ------posPlane--
    ''  //  x < | i |\
    ''  //      +---+ c-----negPlane--
    ''  //             \       v
    ''  //              r
    ''  //
    ''  //  r : reference face
    ''  //  i : incident poly
    ''  //  c : clipped point
    ''  //  n : incident normal



    '// Setup reference face vertices
    'Vec2 v1 = RefPoly->m_vertices[referenceIndex];
    'referenceIndex = referenceIndex + 1 == RefPoly->m_vertexCount ? 0 : referenceIndex + 1;
    'Vec2 v2 = RefPoly->m_vertices[referenceIndex];
    '// Transform vertices to world space
    'v1 = RefPoly->u * v1 + RefPoly->body->position;
    'v2 = RefPoly->u * v2 + RefPoly->body->position;
    '// Setup reference face vertices
    v1 = RefPoly.Vertex(referenceIndex)
    referenceIndex = referenceIndex + 1: If referenceIndex > RefPoly.Nvertex Then referenceIndex = 1
    v2 = RefPoly.Vertex(referenceIndex)

    '// Transform vertices to world space
    'v1 = RefPoly->u * v1 + RefPoly->body->position;
    'v2 = RefPoly->u * v2 + RefPoly->body->position;
    v1 = Vec2ADD(matMULv(RefPoly.U, v1), RefPoly.Pos)
    v2 = Vec2ADD(matMULv(RefPoly.U, v2), RefPoly.Pos)

    '  // Calculate reference face side normal in world space
    '  Vec2 sidePlaneNormal = (v2 - v1);
    '  sidePlaneNormal.Normalize( );
    '  // Orthogonalize
    '  Vec2 refFaceNormal( sidePlaneNormal.y, -sidePlaneNormal.x );
    sidePlaneNormal = Vec2Normalize(Vec2SUB(v2, v1))
    refFaceNormal.X = sidePlaneNormal.y
    refFaceNormal.y = -sidePlaneNormal.X



    '  // ax + by = c
    '  // c is distance from origin
    '  real refC = Dot( refFaceNormal, v1 );
    '  real negSide = -Dot( sidePlaneNormal, v1 );
    '  real posSide =  Dot( sidePlaneNormal, v2 );
    refC = Vec2DOT(refFaceNormal, v1)
    negSide = -Vec2DOT(sidePlaneNormal, v1)
    posSide = Vec2DOT(sidePlaneNormal, v2)



    '  // Clip2 incident face to reference face side planes
    '  if(Clip2( -sidePlaneNormal, negSide, incidentFace ) < 2)
    '    return; // Due to floating point error, possible to not have required points
    '  if(Clip2(  sidePlaneNormal, posSide, incidentFace ) < 2)
    '    return; // Due to floating point error, possible to not have required points
    '  // Flip
    '  m->normal = flip ? -refFaceNormal : refFaceNormal;


    '************** HERE
    If Clip2(Vec2Negative(sidePlaneNormal), negSide, incidentFace) < 2 Then
        Return
    End If
    If Clip2(sidePlaneNormal, posSide, incidentFace) < 2 Then
        Return
    End If

    If flip Then
        CollisionSOLVE.normal = Vec2Negative(refFaceNormal)
    Else
        CollisionSOLVE.normal = refFaceNormal
    End If





    '// Keep points behind reference face
    '  uint32 cp = 0; // clipped points behind reference face
    '  real separation = Dot( refFaceNormal, incidentFace[0] ) - refC;
    '  if(separation <= 0.0f)
    '  {
    '    m->contacts[cp] = incidentFace[0];
    '    m->penetration = -separation;
    '    ++cp;
    '  }
    '  Else
    '    m->penetration = 0;
    '
    '  separation = Dot( refFaceNormal, incidentFace[1] ) - refC;
    '  if(separation <= 0.0f)
    '  {
    '    m->contacts[cp] = incidentFace[1];
    '
    '    m->penetration += -separation;
    '    ++cp;
    '
    '    // Average penetration
    '    m->penetration /= (real)cp;
    '  }
    '
    '  m->contact_count = cp;



    'ReDim CollisionSOLVE.contactsPTS(2)

    '// Keep points behind reference face
    cp = 0
    separation = Vec2DOT(refFaceNormal, incidentFace(0)) - refC
    If (separation <= 0#) Then
        cp = cp + 1

        CollisionSOLVE.contactsPTS(cp) = incidentFace(0)
        CollisionSOLVE.penetration = -separation
        CollisionSOLVE.contactCount = cp
    Else
        CollisionSOLVE.penetration = 0
    End If


    'separation = Dot( refFaceNormal, incidentFace[1] ) - refC;
    separation = Vec2DOT(refFaceNormal, incidentFace(1)) - refC
    If (separation <= 0) Then
        cp = cp + 1
        CollisionSOLVE.contactsPTS(cp) = incidentFace(1)
        CollisionSOLVE.penetration = CollisionSOLVE.penetration - separation
        CollisionSOLVE.penetration = CollisionSOLVE.penetration / cp
        CollisionSOLVE.contactCount = cp
    End If

    CollisionSOLVE.contactCount = cp

    Return

End Function









'Real FindAxisLeastPenetration(uint32 * faceIndex, PolygonShape * A, PolygonShape * B)
'{
'  real bestDistance = -FLT_MAX;
'  uint32 bestIndex;
'
'  for(uint32 i = 0; i < A->m_vertexCount; ++i)
'  {
'    // Retrieve a face normal from A
'    Vec2 n = A->m_normals[i];
'    Vec2 nw = A->u * n;
'
'    // Transform face normal into B's model space
'    Mat2 buT = B->u.Transpose( );
'    n = buT * nw;
'
'    // Retrieve support point from B along -n
'    Vec2 s = B->GetSupport( -n );
'
'    // Retrieve vertex on face from A, transform into
'    // B's model space
'    Vec2 v = A->m_vertices[i];
'    v = A->u * v + A->body->position;
'    v -= B->body->position;
'    v = buT * v;
'
'    // Compute penetration distance (in B's model space)
'    real d = Dot( n, s - v );
'
'    // Store greatest distance
'    if(d > bestDistance)
'    {
'      bestDistance = d;
'      bestIndex = i;
'    }
'  }
Private Function FindAxisLeastPenetration(faceIndex As Long, A As tBody, B As tBody) As Double
    Dim I   As Long

    Dim bestIndex As Long
    Dim bestDistance As Double


    Dim N   As tVec2
    Dim nw  As tVec2
    Dim buT As tMAT2
    Dim S   As tVec2
    Dim V   As tVec2
    Dim D   As Double


    bestDistance = -FLT_MAX

    For I = 1 To A.Nvertex

        '// Retrieve a face normal from A
        ' Vec2 n = A->m_normals[i];
        ' nw = A->u * n;
        N = A.normals(I)

        nw = matMULv(A.U, N)

        '// Transform face normal into B    's model space
        'Mat2 buT = B->u.Transpose( );
        'n = buT * nw;
        buT = matTranspose(B.U)
        N = matMULv(buT, nw)


        '// Retrieve support point from B along -n
        'Vec2 s = B->GetSupport( -n );
        S = GetSupport(B, Vec2Negative(N))

        '// Retrieve vertex on face from A, transform into
        '// B    's model space
        'Vec2 v = A->m_vertices[i];
        'v = A->u * v + A->body->position;
        'v -= B->body->position;
        'v = buT * v;
        V = A.Vertex(I)
        V = Vec2ADD(matMULv(A.U, V), A.Pos)
        V = Vec2SUB(V, B.Pos)
        V = matMULv(buT, V)


        '// Compute penetration distance (in B    's model space)
        'real d = Dot( n, s - v );
        D = Vec2DOT(N, Vec2SUB(S, V))

        '// Store greatest distance
        If (D > bestDistance) Then
            bestDistance = D
            bestIndex = I
        End If
    Next
    FindAxisLeastPenetration = bestDistance
    faceIndex = bestIndex

End Function


'void FindIncidentFace( Vec2 *v, PolygonShape *RefPoly, PolygonShape *IncPoly, uint32 referenceIndex )
'{
'  Vec2 referenceNormal = RefPoly->m_normals[referenceIndex];
'
'  // Calculate normal in incident's frame of reference
'  referenceNormal = RefPoly->u * referenceNormal; // To world space
'  referenceNormal = IncPoly->u.Transpose( ) * referenceNormal; // To incident's model space
'
'  // Find most anti-normal face on incident polygon
'  int32 incidentFace = 0;
'  real minDot = FLT_MAX;
'  for(uint32 i = 0; i < IncPoly->m_vertexCount; ++i)
'  {
'    real dot = Dot( referenceNormal, IncPoly->m_normals[i] );
'    if(dot < minDot)
'    {
'      minDot = dot;
'      incidentFace = i;
'    }
'  }
'
'  // Assign face vertices for incidentFace
'  v[0] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
'  incidentFace = incidentFace + 1 >= (int32)IncPoly->m_vertexCount ? 0 : incidentFace + 1;
'  v[1] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
'}
Private Sub FindIncidentFace(V() As tVec2, RefPoly As tBody, IncPoly As tBody, referenceIndex As Long)

    Dim referenceNormal As tVec2
    Dim incidentFace As Long
    Dim I   As Long
    Dim dot As Double
    Dim minDot As Double

    'Vec2 referenceNormal = RefPoly->m_normals[referenceIndex];
    referenceNormal = RefPoly.normals(referenceIndex)

    '// Calculate normal in incident's frame of reference
    'referenceNormal = RefPoly->u * referenceNormal; // To world space
    'referenceNormal = IncPoly->u.Transpose( ) * referenceNormal; // To incident's model space
    referenceNormal = matMULv(RefPoly.U, referenceNormal)
    referenceNormal = matMULv(matTranspose(IncPoly.U), referenceNormal)


    '// Find most anti-normal face on incident polygon
    '  for(uint32 i = 0; i < IncPoly->m_vertexCount; ++i)
    '  {
    '    real dot = Dot( referenceNormal, IncPoly->m_normals[i] );
    '    if(dot < minDot)
    '    {
    '      minDot = dot;
    '      incidentFace = i;
    '    }
    '  }
    incidentFace = 0
    minDot = FLT_MAX
    For I = 1 To IncPoly.Nvertex
        dot = Vec2DOT(referenceNormal, IncPoly.normals(I))
        If (dot < minDot) Then
            minDot = dot
            incidentFace = I
        End If

    Next

    '// Assign face vertices for incidentFace
    'v[0] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
    'incidentFace = incidentFace + 1 >= (int32)IncPoly->m_vertexCount ? 0 : incidentFace + 1;
    'v[1] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;

    V(0) = Vec2ADD(matMULv(IncPoly.U, IncPoly.Vertex(incidentFace)), IncPoly.Pos)
    incidentFace = incidentFace + 1: If incidentFace > IncPoly.Nvertex Then incidentFace = 1
    V(1) = Vec2ADD(matMULv(IncPoly.U, IncPoly.Vertex(incidentFace)), IncPoly.Pos)


End Sub



'int32 Clip2( Vec2 n, real c, Vec2 *face )
'{
'  uint32 sp = 0;
'  Vec2 out[2] = {
'    face[0],
'    face [1]
'  };
'
'  // Retrieve distances from each endpoint to the line
'  // d = ax + by - c
'  real d1 = Dot( n, face[0] ) - c;
'  real d2 = Dot( n, face[1] ) - c;
'
'  // If negative (behind plane) Clip2
'  if(d1 <= 0.0f) out[sp++] = face[0];
'  if(d2 <= 0.0f) out[sp++] = face[1];
'
'  // If the points are on different sides of the plane
'  if(d1 * d2 < 0.0f) // less than to ignore -0.0f
'  {
'    // Push interesection point
'    real alpha = d1 / (d1 - d2);
'    out[sp] = face[0] + alpha * (face[1] - face[0]);
'    ++sp;
'  }
'
'  // Assign our new converted values
'  face[0] = out[0];
'  face[1] = out[1];
'
'  assert( sp != 3 );
'
'  return sp;
'}



Private Function Clip2(N As tVec2, C As Double, face() As tVec2) As Long
    Dim sp  As Long
    Dim out(0 To 1) As tVec2
    Dim d1  As Double
    Dim d2  As Double
    Dim ALPHA As Double

    '  uint32 sp = 0;
    '  Vec2 out[2] = {
    '    face[0],
    '    face [1]
    '  };
    out(0) = face(0)
    out(1) = face(1)


    ' // Retrieve distances from each endpoint to the line
    ' // d = ax + by - c
    ' real d1 = Dot( n, face[0] ) - c;
    ' real d2 = Dot( n, face[1] ) - c;
    d1 = Vec2DOT(N, face(0)) - C
    d2 = Vec2DOT(N, face(1)) - C


    '// If negative (behind plane) Clip2
    'if(d1 <= 0.0f) out[sp++] = face[0];
    'if(d2 <= 0.0f) out[sp++] = face[1];
    If d1 <= 0 Then out(sp) = face(0): sp = sp + 1
    If d2 <= 0 Then out(sp) = face(1): sp = sp + 1


    '// If the points are on different sides of the plane
    'if(d1 * d2 < 0.0f) // less than to ignore -0.0f
    '{
    '  // Push interesection point
    '  real alpha = d1 / (d1 - d2);
    '  out[sp] = face[0] + alpha * (face[1] - face[0]);
    '  ++sp;
    '}
    If d1 * d2 < 0 Then
        ALPHA = d1 / (d1 - d2)
        out(sp) = Vec2ADD(face(0), Vec2MUL(Vec2SUB(face(1), face(0)), ALPHA))
        sp = sp + 1
    End If

    '// Assign our new converted values
    face(0) = out(0)
    face(1) = out(1)

    '    assert( sp != 3 );
    ''    If sp = 3 Then MsgBox "sp=3"
    Clip2 = sp

End Function
