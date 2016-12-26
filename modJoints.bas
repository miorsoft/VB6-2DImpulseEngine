Attribute VB_Name = "modJoints"
Option Explicit

' by Reexre



Public Type tJoint
    bA      As Long
    bB      As Long
    L       As Double
End Type

Public Joints() As tJoint
Public NJ   As Long


Public Sub AddJoint(bA As Long, bB As Long, D As Double)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    Joints(NJ).bA = bA
    Joints(NJ).bB = bB
    Joints(NJ).L = D
End Sub


Public Sub resolveJoints()
    Dim I   As Long
    Dim D   As Double
    Dim N   As tVec2
Dim Center As tVec2


    For I = 1 To NJ

        With Joints(I)
            N = Vec2SUB(Body(.bA).Pos, Body(.bB).Pos)
            N = Vec2ADD(N, Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), INVdt))
            D = Vec2Length(N)
N = Vec2MUL(N, 1 / D)
'            N.X = N.X / D
'            N.Y = N.Y / D

            D = (D - .L) * DT * (Body(.bA).mass + Body(.bB).mass)

            Body(.bA).FORCE = Vec2ADD(Body(.bA).FORCE, Vec2MUL(N, -D))
            Body(.bB).FORCE = Vec2ADD(Body(.bB).FORCE, Vec2MUL(N, D))

        End With
    Next



End Sub
