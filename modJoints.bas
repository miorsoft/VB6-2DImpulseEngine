Attribute VB_Name = "modJoints"
Option Explicit

' by Reexre


Public Enum eJointType
    JointDistance
    Joint2Pins
    JointPin
End Enum



Public Type tJoint
    JointType As eJointType
    bA      As Long
    bB      As Long
    L       As Double
    AnchA   As tVec2
    AnchB   As tVec2
    tAnchA  As tVec2
    tAnchB  As tVec2

    StifPULL As Double
    StifPUSH As Double
End Type

Public Joints() As tJoint
Public NJ   As Long


Public Sub AddDistanceJoint(bA As Long, bB As Long, D As Double, Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
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

Public Sub Add2PinsJoint(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, D As Double, _
                         Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
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

Public Sub AddPinJoint(wBody As Long, Anch As tVec2, Optional D As Double = 0, _
                       Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointPin
        .bA = wBody
        .AnchA = Anch
        .AnchB = Vec2ADD(Body(wBody).Pos, Anch)
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub




Public Sub resolveJoints()
    Dim I   As Long
    Dim D   As Double
    Dim N   As tVec2
    Dim Center As tVec2
    Dim FIXa As tVec2
    Dim FIXb As tVec2

    Dim vAprj As tVec2
    Dim vBprj As tVec2
    Dim RVA As tVec2
    Dim RVB As tVec2
    Dim Axe As tVec2


    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType

                Case JointDistance

                    N = Vec2SUB(Body(.bA).Pos, Body(.bB).Pos)
                    N = Vec2ADD(N, Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), DT))    ' INVdt))
                    D = Vec2Length(N)
                    N = Vec2MUL(N, 1 / D)
                    D = (D - .L) * (Body(.bA).mass + Body(.bB).mass)  '* DT
                    If D > 0 Then
                        D = D * .StifPULL
                    Else
                        D = D * .StifPUSH
                    End If


                    '                    Body(.bA).FORCE = Vec2ADD(Body(.bA).FORCE, Vec2MUL(N, -D))
                    '                    Body(.bB).FORCE = Vec2ADD(Body(.bB).FORCE, Vec2MUL(N, D))

                    BodyApplyImpulse .bA, Vec2MUL(N, -D), Vec2(0, 0)
                    BodyApplyImpulse .bB, Vec2MUL(N, D), Vec2(0, 0)



                Case Joint2Pins

                    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                    FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                    .tAnchB = matMULv(Body(.bB).U, .AnchB)
                    FIXb = Vec2ADD(Body(.bB).Pos, .tAnchB)

                    'vAprj = VectorProject(Body(.bA).VEL, .tAnchA)
                    'vBprj = VectorProject(Body(.bB).VEL, .tAnchB)
                    Axe = Vec2SUB(.tAnchB, .tAnchB)
                    vAprj = VectorProject(Body(.bA).VEL, Axe)
                    vBprj = VectorProject(Body(.bB).VEL, Axe)


                    'Consider Angular velocity too.
                    RVA = Vec2CROSSav(Body(.bA).angularVelocity, .tAnchA)
                    RVB = Vec2CROSSav(Body(.bB).angularVelocity, .tAnchB)
                    FIXa = Vec2ADD(FIXa, Vec2MUL(RVA, DT))
                    FIXb = Vec2ADD(FIXb, Vec2MUL(RVB, DT))
                    '-------------------------

                    N = Vec2SUB(FIXa, FIXb)
                    'N = Vec2ADD(N, Vec2MUL(Vec2SUB(vAprj, vBprj), INVdt))
                    N = Vec2ADD(N, Vec2MUL(Vec2SUB(vAprj, vBprj), DT))


                    D = Vec2Length(N)
                    N = Vec2MUL(N, 1 / D)
                    D = (D - .L) * (Body(.bA).mass + Body(.bB).mass)  '*dt
                    If D > 0 Then
                        D = D * .StifPULL
                    Else
                        D = D * .StifPUSH
                    End If

                    BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                    BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB
                Case JointPin

                    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                    vAprj = VectorProject(Body(.bA).VEL, .tAnchA)



                    'Consider Angular velocity too.
                    RVA = Vec2CROSSav(Body(.bA).angularVelocity, .tAnchA)
                    RVA = Vec2MUL(RVA, DT)
                    '-------------------------
                    'FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                    FIXa = Vec2ADD(Body(.bA).Pos, Vec2ADD(.tAnchA, RVA))
                    N = Vec2SUB(FIXa, .AnchB)
                    N = Vec2ADD(N, Vec2MUL(vAprj, DT))

                    D = Vec2Length(N)
                    N = Vec2MUL(N, 1 / D)
                    D = (D - .L) * (Body(.bA).mass)    ' * DT

                    If D > 0 Then
                        D = D * .StifPULL
                    Else
                        D = D * .StifPUSH
                    End If

                    BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA

            End Select

        End With
    Next

End Sub
