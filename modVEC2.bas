Attribute VB_Name = "modVEC2"
Option Explicit

Public Type tVec2
    x          As Single
    y          As Single
End Type

Public Type tMAT2
    m00        As Single
    m01        As Single
    m10        As Single
    m11        As Single
End Type

Public Function Vec2(x As Single, y As Single) As tVec2

    Vec2.x = x
    Vec2.y = y

End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.x = -V.x
    Vec2Negative.y = -V.y
End Function



Public Function Vec2ADD(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2ADD.x = v1.x + v2.x
    Vec2ADD.y = v1.y + v2.y
End Function

Public Function Vec2SUB(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2SUB.x = v1.x - v2.x
    Vec2SUB.y = v1.y - v2.y
End Function

Public Function Vec2MULV(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2MULV.x = v1.x * v2.x
    Vec2MULV.y = v1.y * v2.y
End Function
Public Function Vec2MUL(V As tVec2, S As Single) As tVec2
    Vec2MUL.x = V.x * S
    Vec2MUL.y = V.y * S
End Function

Public Function Vec2ADDScaled(v1 As tVec2, v2 As tVec2, S As Single) As tVec2
    Vec2ADDScaled.x = v1.x + v2.x * S
    Vec2ADDScaled.y = v1.y + v2.y * S
End Function

Public Function Vec2LengthSq(V As tVec2) As Single
    Vec2LengthSq = V.x * V.x + V.y * V.y
End Function

Public Function Vec2Length(V As tVec2) As Single
'   Vec2Length = FASTsqr(V.X * V.X + V.Y * V.Y)
    Vec2Length = Sqr(V.x * V.x + V.y * V.y)

End Function


Public Function Vec2Rotate(V As tVec2, radians As Single) As tVec2
'real c = std::cos( radians );
'real s = std::sin( radians );

'real xp = x * c - y * s;
'real yp = x * s + y * c;

    Dim S      As Single
    Dim C      As Single
    C = Cos(radians)
    S = Sin(radians)

    Vec2Rotate.x = V.x * C - V.y * S
    Vec2Rotate.y = V.x * S + V.y * C
End Function

Public Function Vec2Normalize(V As tVec2) As tVec2
    Dim D      As Single
    D = Vec2Length(V)
    If D Then
        D = 1 / D
        Vec2Normalize.x = V.x * D
        Vec2Normalize.y = V.y * D
    End If

End Function

Public Function Vec2MIN(A As tVec2, B As tVec2) As tVec2
    Vec2MIN.x = IIf(A.x < B.x, A.x, B.x)
    Vec2MIN.y = IIf(A.y < B.y, A.y, B.y)
End Function

Public Function Vec2MAX(A As tVec2, B As tVec2) As tVec2
    Vec2MAX.x = IIf(A.x > B.x, A.x, B.x)
    Vec2MAX.y = IIf(A.y > B.y, A.y, B.y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function Vec2DOT(A As tVec2, B As tVec2) As Single
    Vec2DOT = A.x * B.x + A.y * B.y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, A As Single) As tVec2
    Vec2CROSSva.x = A * V.y
    Vec2CROSSva.y = -A * V.x
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(A As Single, V As tVec2) As tVec2
    Vec2CROSSav.x = -A * V.y
    Vec2CROSSav.y = A * V.x
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(A As tVec2, B As tVec2) As Single
    Vec2CROSS = A.x * B.y - A.y * B.x
End Function


Public Function Vec2DISTANCEsq(A As tVec2, B As tVec2) As Single
    Dim Dx     As Single
    Dim DY     As Single
    Dx = A.x - B.x
    DY = A.y - B.y
    Vec2DISTANCEsq = Dx * Dx + DY * DY
End Function


'************************************************************************************



Public Function matTranspose(M As tMAT2) As tMAT2
    matTranspose.m00 = M.m00
    matTranspose.m01 = M.m10    '
    matTranspose.m10 = M.m01    '
    matTranspose.m11 = M.m11

End Function

Public Function matMULv(M As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );

    matMULv.x = M.m00 * V.x + M.m01 * V.y
    matMULv.y = M.m10 * V.x + M.m11 * V.y

End Function

Public Function SetOrient(radians As Single) As tMAT2
'    real c = std::cos( radians );
'    real s = std::sin( radians );
'
'    m00 = c; m01 = -s;
'    m10 = s; m11 =  c;

    Dim C      As Single
    Dim S      As Single

    C = Cos(radians)
    S = Sin(radians)

    SetOrient.m00 = C
    SetOrient.m01 = -S
    SetOrient.m10 = S
    SetOrient.m11 = C


End Function


Public Function VectorProject(ByRef V As tVec2, ByRef Vto As tVec2) As tVec2
'Poject Vector V to vector Vto
    Dim K      As Single
    Dim D      As Single



    D = Vto.x * Vto.x + Vto.y * Vto.y
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    K = (V.x * Vto.x + V.y * Vto.y) * D

    VectorProject.x = (Vto.x * D) * K
    VectorProject.y = (Vto.y * D) * K

End Function

Public Function VectorReflect(ByRef V As tVec2, ByRef wall As tVec2) As tVec2
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot   As Single
    Dim D      As Single
    Dim NwX    As Single
    Dim NwY    As Single

    D = (wall.x * wall.x + wall.y * wall.y)
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    NwX = wall.x * D
    NwY = wall.y * D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.x * NwX + V.y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.x = -V.x + NwX
    VectorReflect.y = -V.y + NwY


End Function


Public Function ACOS(x As Single) As Single

    ACOS = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)

End Function

Public Function AngleDIFF(A1 As Single, A2 As Single) As Single

    AngleDIFF = A1 - A2
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend

End Function
