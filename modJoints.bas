Attribute VB_Name = "modJoints"
Option Explicit

' by reexre miorsoft


Public Enum eJointType
    JointDistance
    Joint2Pins
    JointPin
    Rotor1
    Rotor2
End Enum



Public Type tJoint
    JointType As eJointType
    bA      As Long
    bB      As Long
    L       As Single
    AnchA   As tVec2
    AnchB   As tVec2
    tAnchA  As tVec2
    tAnchB  As tVec2

    StifPULL As Single
    StifPUSH As Single
End Type

Public Joints() As tJoint
Public NJ   As Long

Private Const KMASS As Single = 0.1
Private Const KDamp As Single = 1


Public Sub AddDistanceJoint(bA As Long, bB As Long, D As Single, Optional StiffPull As Single = 1, Optional stiffPush As Single = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointDistance
        .bA = bA
        .bB = bB
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Public Sub Add2PinsJoint(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, D As Single, _
                         Optional StiffPull As Single = 1, Optional stiffPush As Single = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Joint2Pins
        .bA = bA
        .bB = bB
        .AnchA = AnchA
        .AnchB = AnchB
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Public Sub AddPinJoint(wBody As Long, Anch As tVec2, Optional D As Single = 0, _
                       Optional StiffPull As Single = 1, Optional stiffPush As Single = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointPin
        .bA = wBody
        .bB = 0
        .AnchA = Anch
        .AnchB = Vec2ADD(Body(wBody).Pos, Anch)
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Public Sub AddRotorJoint(wBody As Long, Leva As tVec2, Speed As Single)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Rotor1
        .bA = wBody
        .bB = 0
        .AnchA = Leva
        .L = Speed
    End With
End Sub

Public Sub AddRotor2Joint(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, _
                         Optional StiffPull As Single = 1, Optional stiffPush As Single = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Rotor2
        .bA = bA
        .bB = bB
        .AnchA = AnchA
        .AnchB = AnchB
        .L = 0
'        .StifPULL = StiffPull
'        .StifPUSH = stiffPush
    End With
End Sub

Public Sub ResolveJoints()
    Dim I      As Long
    Dim J      As Long

    Dim D      As Single
    Dim N      As tVec2
    Dim Center As tVec2
    Dim FIXa   As tVec2
    Dim FIXb   As tVec2

    Dim vAprj  As tVec2
    Dim vBprj  As tVec2
    Dim RVA    As tVec2
    Dim RVB    As tVec2
    Dim Axe    As tVec2

    Dim Tmass  As Single
    Dim ORI    As tMAT2
    Dim RotDiff As Single


    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType

            Case JointDistance

                N = Vec2SUB(Body(.bA).Pos, Body(.bB).Pos)
                N = Vec2ADD(N, Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), 1 * KDamp))    ' INVdt))
                D = Vec2Length(N)
                N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If

                '                  Body(.bA).FORCE = Vec2ADD(Body(.bA).FORCE, Vec2MUL(N, -D))
                '                  Body(.bB).FORCE = Vec2ADD(Body(.bB).FORCE, Vec2MUL(N, D))


                'For J = 1 To Iterations
                BodyApplyImpulse .bA, Vec2MUL(N, -D), Vec2(0, 0)
                BodyApplyImpulse .bB, Vec2MUL(N, D), Vec2(0, 0)
                'Next

            Case Joint2Pins

                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                '    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                .tAnchA = matMULv(ORI, .AnchA)
                FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)

                ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity)
                '.tAnchB = matMULv(Body(.bB).U, .AnchB)
                .tAnchB = matMULv(ORI, .AnchB)

                FIXb = Vec2ADD(Body(.bB).Pos, .tAnchB)

                'vAprj = VectorProject(Body(.bA).VEL, .tAnchA)
                'vBprj = VectorProject(Body(.bB).VEL, .tAnchB)
                'Axe = Vec2SUB(.tAnchB, .tAnchB)
                Axe = Vec2SUB(FIXb, FIXa)

                '                                    vAprj = VectorProject(Body(.bA).VEL, Axe)
                '                                    vBprj = VectorProject(Body(.bB).VEL, Axe)
                vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * KDamp), Axe)
                vBprj = VectorProject(Vec2MUL(Body(.bB).VEL, 1 * KDamp), Axe)

                'Consider Angular velocity too.
                ''                    RVA = Vec2CROSSav(Body(.bA).angularVelocity, .tAnchA)
                ''                    RVB = Vec2CROSSav(Body(.bB).angularVelocity, .tAnchB)
                ''                    RVA = VectorProject(Vec2MUL(RVA, 1 * KDamp * Body(.bA).invInertia), Axe)
                ''                    RVB = VectorProject(Vec2MUL(RVB, -1 * KDamp * Body(.bB).invInertia), Axe)
                ''                    FIXa = Vec2ADD(FIXa, RVA)
                ''                    FIXb = Vec2ADD(FIXb, RVB)
                '-------------------------

                N = Vec2SUB(FIXa, FIXb)
                '                N = Vec2ADD(N, Vec2MUL(Vec2SUB(vAprj, vBprj), DT))
                N = Vec2ADD(N, Vec2SUB(vAprj, vBprj))

                D = Vec2Length(N)
                If D Then N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If


                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB


            Case JointPin



                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)

                '                    .tAnchA = matMULv(Body(.bA).U + Body(.bA).angularVelocity, .AnchA)
                .tAnchA = matMULv(ORI, .AnchA)

                Axe = .tAnchA

                'vAprj = VectorProject(Body(.bA).VEL, Axe)
                vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * KDamp), Axe)
                '
                '
                '                    'Consider Angular velocity too.
                '                    RVA = Vec2CROSSav(Body(.bA).angularVelocity, .tAnchA)
                '                    RVA = Vec2MUL(RVA, 1 * KDamp * Body(.bA).invInertia)
                '                    '-------------------------


                'FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                FIXa = Vec2ADD(Body(.bA).Pos, Vec2ADD(.tAnchA, RVA))
                N = Vec2SUB(FIXa, .AnchB)

                'N = Vec2ADD(N, Vec2MUL(vAprj, DT))
                N = Vec2ADD(N, Vec2MUL(vAprj, 1))

                D = Vec2Length(N)
                If D Then N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass) * KMASS * 2

                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If

                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA

            Case Rotor1


                'ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                .tAnchA = matMULv(ORI, .AnchA)
                'FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                N.x = -.tAnchA.y
                N.y = .tAnchA.x
                RotDiff = .L - Body(.bA).angularVelocity
                N = Vec2MUL(N, RotDiff)

                If Sgn(RotDiff) = Sgn(.L) Then
                    BodyApplyImpulse .bA, N, Vec2ADD(.tAnchA, Body(.bA).VEL)
                End If



            Case Rotor2
'''' Like Joint2pins

                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                '    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                .tAnchA = matMULv(ORI, .AnchA)
                FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)

                ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity)
                '.tAnchB = matMULv(Body(.bB).U, .AnchB)
                .tAnchB = matMULv(ORI, .AnchB)

                FIXb = Vec2ADD(Body(.bB).Pos, .tAnchB)

                Axe = Vec2SUB(FIXb, FIXa)

                vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * KDamp), Axe)
                vBprj = VectorProject(Vec2MUL(Body(.bB).VEL, 1 * KDamp), Axe)

                N = Vec2SUB(FIXa, FIXb)
                N = Vec2ADD(N, Vec2SUB(vAprj, vBprj))

                D = Vec2Length(N)
                If D Then N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * 0.5
                Else
                    D = D * 0.5
                End If

                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB
'---------------------------------

                RotDiff = AngleDIFF(PI * 0.8 * Cos(CNT * 0.0008) + _
                Body(.bB).orient + Body(.bB).angularVelocity, _
                Body(.bA).orient + Body(.bA).angularVelocity)
                
               If Abs(RotDiff) > 1 Then RotDiff = Sgn(RotDiff) * 1
                
                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                .tAnchA = matMULv(ORI, .AnchA)
                .tAnchA.x = -.tAnchA.x
                .tAnchA.y = -.tAnchA.y
                N.x = -.tAnchA.y
                N.y = .tAnchA.x
                N = Vec2MUL(N, RotDiff)
                BodyApplyImpulse .bA, N, Vec2ADD(.tAnchA, Body(.bA).VEL)
               

ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity)
                .tAnchB = matMULv(ORI, .AnchB)
                .tAnchB.x = -.tAnchB.x
                .tAnchB.y = -.tAnchB.y
                N.x = -.tAnchB.y
                N.y = .tAnchB.x
                N = Vec2MUL(N, -RotDiff)
                BodyApplyImpulse .bB, N, Vec2ADD(.tAnchB, Body(.bB).VEL)
               






            End Select

        End With
    Next

End Sub
