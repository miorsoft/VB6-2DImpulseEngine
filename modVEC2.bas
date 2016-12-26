Attribute VB_Name = "modVEC2"
Option Explicit

Public Type tVec2
    X       As Double
    Y       As Double
End Type


Public Function Vec2(X As Double, Y As Double) As tVec2

    Vec2.X = X
    Vec2.Y = Y

End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.X = -V.X
    Vec2Negative.Y = -V.Y
End Function



Public Function Vec2ADD(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2ADD.X = v1.X + v2.X
    Vec2ADD.Y = v1.Y + v2.Y
End Function

Public Function Vec2SUB(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2SUB.X = v1.X - v2.X
    Vec2SUB.Y = v1.Y - v2.Y
End Function

Public Function Vec2MULV(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2MULV.X = v1.X * v2.X
    Vec2MULV.Y = v1.Y * v2.Y
End Function
Public Function Vec2MUL(V As tVec2, S As Double) As tVec2
    Vec2MUL.X = V.X * S
    Vec2MUL.Y = V.Y * S
End Function

Public Function Vec2ADDS(v1 As tVec2, v2 As tVec2, S As Double) As tVec2
    Vec2ADDS.X = v1.X + v2.X * S
    Vec2ADDS.Y = v1.Y + v2.Y * S
End Function

Public Function Vec2LengthSq(V As tVec2) As Double
    Vec2LengthSq = V.X * V.X + V.Y * V.Y
End Function

Public Function Vec2Length(V As tVec2) As Double
'   Vec2Length = FASTsqr(V.X * V.X + V.Y * V.Y)
    Vec2Length = Sqr(V.X * V.X + V.Y * V.Y)

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

    Vec2Rotate.X = V.X * C - V.Y * S
    Vec2Rotate.Y = V.X * S + V.Y * C
End Function

Public Function Vec2Normalize(V As tVec2) As tVec2
    Dim D   As Double
    D = Vec2Length(V)
    If D Then
        D = 1 / D
        Vec2Normalize.X = V.X * D
        Vec2Normalize.Y = V.Y * D
    End If

End Function

Public Function Vec2MIN(A As tVec2, B As tVec2) As tVec2
    Vec2MIN.X = IIf(A.X < B.X, A.X, B.X)
    Vec2MIN.Y = IIf(A.Y < B.Y, A.Y, B.Y)
End Function

Public Function Vec2MAX(A As tVec2, B As tVec2) As tVec2
    Vec2MAX.X = IIf(A.X > B.X, A.X, B.X)
    Vec2MAX.Y = IIf(A.Y > B.Y, A.Y, B.Y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function Vec2DOT(A As tVec2, B As tVec2) As Double
    Vec2DOT = A.X * B.X + A.Y * B.Y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, A As Double) As tVec2
    Vec2CROSSva.X = A * V.Y
    Vec2CROSSva.Y = -A * V.X
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(A As Double, V As tVec2) As tVec2
    Vec2CROSSav.X = -A * V.Y
    Vec2CROSSav.Y = A * V.X
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(A As tVec2, B As tVec2) As Double
    Vec2CROSS = A.X * B.Y - A.Y * B.X
End Function


Public Function Vec2DISTANCEsq(A As tVec2, B As tVec2) As Double
    Dim dX  As Double
    Dim DY  As Double
    dX = A.X - B.X
    DY = A.Y - B.Y
    Vec2DISTANCEsq = dX * dX + DY * DY
End Function
