Attribute VB_Name = "modVEC2"
Option Explicit

Public Type tVec2
    X       As Double
    y       As Double
End Type

Public Type tMAT2
    m00     As Double
    m01     As Double
    m10     As Double
    m11     As Double
End Type

Public Function Vec2(X As Double, y As Double) As tVec2

    Vec2.X = X
    Vec2.y = y

End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.X = -V.X
    Vec2Negative.y = -V.y
End Function



Public Function Vec2ADD(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2ADD.X = v1.X + v2.X
    Vec2ADD.y = v1.y + v2.y
End Function

Public Function Vec2SUB(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2SUB.X = v1.X - v2.X
    Vec2SUB.y = v1.y - v2.y
End Function

Public Function Vec2MULV(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2MULV.X = v1.X * v2.X
    Vec2MULV.y = v1.y * v2.y
End Function
Public Function Vec2MUL(V As tVec2, S As Double) As tVec2
    Vec2MUL.X = V.X * S
    Vec2MUL.y = V.y * S
End Function

Public Function Vec2ADDScaled(v1 As tVec2, v2 As tVec2, S As Double) As tVec2
    Vec2ADDScaled.X = v1.X + v2.X * S
    Vec2ADDScaled.y = v1.y + v2.y * S
End Function

Public Function Vec2LengthSq(V As tVec2) As Double
    Vec2LengthSq = V.X * V.X + V.y * V.y
End Function

Public Function Vec2Length(V As tVec2) As Double
'   Vec2Length = FASTsqr(V.X * V.X + V.Y * V.Y)
    Vec2Length = Sqr(V.X * V.X + V.y * V.y)

End Function


Public Function Vec2Rotate(V As tVec2, radians As Double) As tVec2
'real c = std::cos( radians );
'real s = std::sin( radians );

'real xp = x * c - y * s;
'real yp = x * s + y * c;

    Dim S   As Double
    Dim C   As Double
    C = Cos(radians)
    S = Sin(radians)

    Vec2Rotate.X = V.X * C - V.y * S
    Vec2Rotate.y = V.X * S + V.y * C
End Function

Public Function Vec2Normalize(V As tVec2) As tVec2
    Dim D   As Double
    D = Vec2Length(V)
    If D Then
        D = 1 / D
        Vec2Normalize.X = V.X * D
        Vec2Normalize.y = V.y * D
    End If

End Function

Public Function Vec2MIN(A As tVec2, B As tVec2) As tVec2
    Vec2MIN.X = IIf(A.X < B.X, A.X, B.X)
    Vec2MIN.y = IIf(A.y < B.y, A.y, B.y)
End Function

Public Function Vec2MAX(A As tVec2, B As tVec2) As tVec2
    Vec2MAX.X = IIf(A.X > B.X, A.X, B.X)
    Vec2MAX.y = IIf(A.y > B.y, A.y, B.y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function Vec2DOT(A As tVec2, B As tVec2) As Double
    Vec2DOT = A.X * B.X + A.y * B.y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, A As Double) As tVec2
    Vec2CROSSva.X = A * V.y
    Vec2CROSSva.y = -A * V.X
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(A As Double, V As tVec2) As tVec2
    Vec2CROSSav.X = -A * V.y
    Vec2CROSSav.y = A * V.X
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(A As tVec2, B As tVec2) As Double
    Vec2CROSS = A.X * B.y - A.y * B.X
End Function


Public Function Vec2DISTANCEsq(A As tVec2, B As tVec2) As Double
    Dim dX  As Double
    Dim DY  As Double
    dX = A.X - B.X
    DY = A.y - B.y
    Vec2DISTANCEsq = dX * dX + DY * DY
End Function


'************************************************************************************



Public Function matTranspose(m As tMAT2) As tMAT2
    matTranspose.m00 = m.m00
    matTranspose.m01 = m.m10    '
    matTranspose.m10 = m.m01    '
    matTranspose.m11 = m.m11

End Function

Public Function matMULv(m As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );

    matMULv.X = m.m00 * V.X + m.m01 * V.y
    matMULv.y = m.m10 * V.X + m.m11 * V.y

End Function

Public Function SetOrient(radians As Double) As tMAT2
'    real c = std::cos( radians );
'    real s = std::sin( radians );
'
'    m00 = c; m01 = -s;
'    m10 = s; m11 =  c;

    Dim C   As Double
    Dim S   As Double

    C = Cos(radians)
    S = Sin(radians)

    SetOrient.m00 = C
    SetOrient.m01 = -S
    SetOrient.m10 = S
    SetOrient.m11 = C


End Function


Public Function VectorProject(ByRef V As tVec2, ByRef Vto As tVec2) As tVec2
'Poject Vector V to vector Vto
    Dim K   As Double
    Dim D   As Double



    D = Vto.X * Vto.X + Vto.y * Vto.y
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    K = (V.X * Vto.X + V.y * Vto.y) * D

    VectorProject.X = (Vto.X * D) * K
    VectorProject.y = (Vto.y * D) * K

End Function

Public Function VectorReflect(ByRef V As tVec2, ByRef wall As tVec2) As tVec2
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot As Double
    Dim D   As Double
    Dim NwX As Double
    Dim NwY As Double

    D = (wall.X * wall.X + wall.y * wall.y)
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    NwX = wall.X * D
    NwY = wall.y * D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.X * NwX + V.y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.X = -V.X + NwX
    VectorReflect.y = -V.y + NwY


End Function
