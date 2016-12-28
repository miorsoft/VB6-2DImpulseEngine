Attribute VB_Name = "modMAIN"
Option Explicit


'Private Declare Function GetTickCount Lib "kernel32" () As Long

Public INVdt As Double


Public pHDC As Long
Public PicW As Long
Public PicH As Long
Public Frame As Long
Public SaveFrames As Long

Public TotalNContacts As Long

Public Version As String

Public DisplayRefreshPeriod As Long


'Public DT As Double

'Public Sub doSTEP(DeltaTime As Double)
Public Sub doSTEP()

    Dim I   As Long
    Dim J   As Long
    Dim A   As tBody

    Dim MinX As Double
    Dim MinY As Double
    Dim MaxX As Double
    Dim MaxY As Double

    'DT = DeltaTime
    'RESTING = Vec2LengthSq(Vec2MUL(GRAVITY, DT)) + EPSILON


    Dim tmpContacts As tManifold

    Dim ContactType As Long

    Dim V   As tVec2

    ''   // Generate new collision info
    ''       contacts.clear();
    NofContactMainFolds = 0

    For I = 1 To NBodies - 1
        A = Body(I)

        For J = I + 1 To NBodies

            If (A.invMass <> 0 Or Body(J).invMass <> 0) Then

                '                Manifold m = new Manifold( A, B );
                '                m.solve();
                '
                '                if (m.contactCount > 0)
                '                {
                '                contacts.add( m );
                '                }

                If AABBvsAABB(I, J) Then
                    tmpContacts = CollisionSOLVE(I, J)

                    If tmpContacts.contactCount > 0 Then
                        NofContactMainFolds = NofContactMainFolds + 1
                        If NofContactMainFolds > MAXNofContactMainFolds Then
                            MAXNofContactMainFolds = NofContactMainFolds + 20
                            ReDim Preserve Contacts(MAXNofContactMainFolds)
                        End If

                        Contacts(NofContactMainFolds) = tmpContacts
                    End If

                End If


            End If

        Next
    Next

    '  resolveJoints

    '// Integrate forces
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'integrateForces( bodies.get( i ), dt );
    '}
    For I = 1 To NBodies
        integrateForces I    ', DT
    Next


    ''// Initialize collision
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).initialize();
    '}

    For I = 1 To NofContactMainFolds
        ContactsINIT I
    Next



    '// Solve collisions
    'for (int j = 0; j < iterations; ++j)
    '{
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).applyImpulse();
    '}
    '}

    For J = 1 To Iterations
        For I = 1 To NofContactMainFolds
            contactsApplyImpulse I
        Next
    Next


    '// Integrate velocities
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'integrateVelocity( bodies.get( i ), dt );
    '}
    For I = 1 To NBodies
        integrateVelocity I    ', DT
    Next

    '// Correct positions
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).positionalCorrection();
    '}

    For I = 1 To NofContactMainFolds
        contactsPositionalCorrection I
    Next


    '// Clear all forces
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'Body b = bodies.get( i );
    'b.force.set( 0, 0 );
    'b.torque = 0;
    '}



    For I = 1 To NBodies
        With Body(I)
            .FORCE.X = 0
            .FORCE.y = 0
            .torque = 0


            'SET AABB
            If .myType = eCircle Then

                .AABB.pMin = Vec2ADD(.Pos, Vec2(-.radius, -.radius))
                .AABB.pMax = Vec2ADD(.Pos, Vec2(.radius, .radius))

            Else

                MinX = MAX_VALUE
                MinY = MAX_VALUE
                MaxX = -MAX_VALUE
                MaxY = -MAX_VALUE

                For J = 1 To .Nvertex
                    V = matMULv(.U, .Vertex(J))
                    If V.X < MinX Then MinX = V.X
                    If V.y < MinY Then MinY = V.y
                    If V.X > MaxX Then MaxX = V.X
                    If V.y > MaxY Then MaxY = V.y
                    .tVertex(J) = V
                Next

                .AABB.pMin = Vec2ADD(Vec2(MinX, MinY), .Pos)
                .AABB.pMax = Vec2ADD(Vec2(MaxX, MaxY), .Pos)

            End If


            '-------------------------------


        End With


    Next


    resolveJoints



End Sub
Private Function AABBvsAABB(wA As Long, wB As Long) As Boolean
    Dim ab1 As tAABB
    Dim ab2 As tAABB

    ab1 = Body(wA).AABB
    ab2 = Body(wB).AABB

    If ab1.pMin.y > ab2.pMax.y Then Exit Function
    If ab2.pMin.y > ab1.pMax.y Then Exit Function
    If ab1.pMin.X > ab2.pMax.X Then Exit Function
    If ab2.pMin.X > ab1.pMax.X Then Exit Function

    AABBvsAABB = True

End Function





Private Sub integrateForces(wB As Long)    ', DT As Double)
'''    // see http://www.niksula.hut.fi/~hkankaan/Homepages/gravity.html
'''    public void integrateForces( Body b, float dt )
'''    {
'''//      if(b->im == 0.0f)
'''//          return;
'''//      b->velocity += (b->force * b->im + gravity) * (dt / 2.0f);
'''//      b->angularVelocity += b->torque * b->iI * (dt / 2.0f);
'''
'''        if (b.invMass == 0.0f)
'''        {
'''            return;
'''        }
'''
'''        float dts = dt * 0.5f;
'''
'''        b.velocity.addsi( b.force, b.invMass * dts );
'''        b.velocity.addsi( ImpulseMath.GRAVITY, dts );
'''        b.angularVelocity += b.torque * b.invInertia * dts;
'''    }
    Dim dts As Double
    dts = DT * 0.5

    With Body(wB)
        If .invMass <> 0 Then


            .VEL = Vec2ADD(.VEL, Vec2MUL(.FORCE, .invMass * dts))
            '.VEL = Vec2ADD(.VEL, Vec2MUL(GRAVITY, .invMass * dts))
            .VEL = Vec2ADD(.VEL, Vec2MUL(GRAVITY, dts))

            .angularVelocity = .angularVelocity + .torque * .invInertia * dts

            .angularVelocity = .angularVelocity * 0.9999    'Air Resistence
            .VEL = Vec2MUL(.VEL, 0.9999)


            'If .Pos.Y + .radius > PicH And .Pos.X - .radius < 0 Then
            If .Pos.y + .radius > PicH Then

                While .Pos.X > PicW
                    .Pos.X = .Pos.X - PicW
                Wend
                While .Pos.X < 0
                    .Pos.X = .Pos.X + PicW
                Wend
                .Pos.y = 0
                .VEL.y = 0
            End If


            '    If .Pos.X + .radius > PicW Then
            '
            '    BodyApplyImpulse wB, Vec2(-.VEL.X * 2, 0), Vec2(.Pos.X + .radius, .Pos.Y)
            '    End If


            '            If .Pos.X + .radius > PicW Then
            '                .VEL.X = -.VEL.X * .restitution
            '                .Pos.X = PicW - .radius
            '            End If

            '   If .Pos.Y + .radius > PicH Then
            '                .VEL.Y = -.VEL.Y * .restitution
            '                .Pos.Y = PicH - .radius
            '            End If

        End If

    End With
End Sub

Private Sub integrateVelocity(wB As Long)    ', DT As Double)
'        if (b.invMass == 0.0f)
'        {
'            return;
'        }
'
'        b.position.addsi( b.velocity, dt );
'        b.orient += b.angularVelocity * dt;
'        b.setOrient( b.orient );
'
'        integrateForces( b, dt );
    With Body(wB)
        If .invMass <> 0 Then

            .Pos = Vec2ADD(.Pos, Vec2MUL(.VEL, DT))
            .orient = .orient + .angularVelocity * DT

            'If .myType = ePolygon Then
            .U = SetOrient(.orient)
            'End If

            'integrateForces wB, DT

        End If

    End With
End Sub




Public Sub MAINLOOP()
    Dim CNT As Long



    Dim I   As Long

    Dim A   As Long
    Dim B   As Long

    '    Dim Accumulator As Long
    '    Dim currTime As Long
    '    Dim frameStart As Long
    '    frameStart = GetTickCount



    Do

        'currTime = GetTickCount
        'Accumulator = Accumulator + currTime - frameStart
        'frameStart = currTime
        'If Accumulator > 200 Then Accumulator = 200
        'While Accumulator > 10
        '    doSTEP Accumulator * 0.01
        '    Accumulator = Accumulator - 10
        'Wend

        doSTEP


        If CNT Mod DisplayRefreshPeriod = 0 Then

            RENDERrc
            frmMain.PIC.Refresh


            TotalNContacts = 0
            For I = 1 To NofContactMainFolds
                TotalNContacts = TotalNContacts + Contacts(I).contactCount
            Next


            If SaveFrames Then
                vbDRAW.Srf.WriteContentToJpgFile App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", JPGQuality
                Frame = Frame + 1
            End If

        End If

        CNT = CNT + 1



        'If Rnd < 0.0001 Then
        'Do
        'A = 1 + Rnd * (NBodies - 1)
        'Loop While Body(A).invMass = 0
        'Do
        'B = 1 + Rnd * (NBodies - 1)
        'Loop While Body(B).invMass = 0 Or (A = B)
        '
        'AddDistanceJoint A, B, 60
        'End If

    Loop While True




End Sub
