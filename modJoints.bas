Attribute VB_Name = "modJoints"
Option Explicit

' by Reexre
Public Enum eJointType
    JointDistance
    JointPINS
End Enum



Public Type tJoint
    bA      As Long
    bB      As Long
    L       As Double
    AnchA   As tVec2
    AnchB   As tVec2
    tAnchA  As tVec2
    tAnchB  As tVec2
    JointType As eJointType
End Type

Public Joints() As tJoint
Public NJ   As Long


Public Sub AddDistanceJoint(bA As Long, bB As Long, D As Double)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointDistance
        .bA = bA
        .bB = bB
        .L = D
    End With
End Sub

Public Sub AddPinsJoint(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, D As Double)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointPINS
        .bA = bA
        .bB = bB
        .AnchA = AnchA
        .AnchB = AnchB
        .L = D
    End With
End Sub




Public Sub resolveJoints()
    Dim I   As Long
    Dim D   As Double
    Dim N   As tVec2
    Dim Center As tVec2
    Dim pA  As tVec2
    Dim pB  As tVec2

    Dim vAprj As tVec2
    Dim vBprj As tVec2


    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType
                Case JointDistance

                    N = Vec2SUB(Body(.bA).Pos, Body(.bB).Pos)
                    N = Vec2ADD(N, Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), 1))    ' INVdt))
                    D = Vec2Length(N)
                    N = Vec2MUL(N, 1 / D)
                    D = (D - .L) * DT * (Body(.bA).mass + Body(.bB).mass)

                    '                    Body(.bA).FORCE = Vec2ADD(Body(.bA).FORCE, Vec2MUL(N, -D))
                    '                    Body(.bB).FORCE = Vec2ADD(Body(.bB).FORCE, Vec2MUL(N, D))

                    BodyApplyImpulse .bA, Vec2MUL(N, -D), Vec2(0, 0)
                    BodyApplyImpulse .bB, Vec2MUL(N, D), Vec2(0, 0)



                Case JointPINS

                    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                    pA = Vec2ADD(Body(.bA).Pos, .tAnchA)
                    .tAnchB = matMULv(Body(.bB).U, .AnchB)
                    pB = Vec2ADD(Body(.bB).Pos, .tAnchB)

                    vAprj = VectorProject(Body(.bA).VEL, .tAnchA)
                    vBprj = VectorProject(Body(.bB).VEL, .tAnchB)

                    N = Vec2SUB(pA, pB)
                    N = Vec2ADD(N, Vec2MUL(Vec2SUB(vAprj, vBprj), 1))    ' INVdt))

                    D = Vec2Length(N)
                    N = Vec2MUL(N, 1 / D)
                    D = (D - .L) * DT * (Body(.bA).mass + Body(.bB).mass)

                    BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                    BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB


            End Select

        End With
    Next

End Sub
