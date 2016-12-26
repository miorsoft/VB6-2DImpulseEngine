Attribute VB_Name = "modMAT2"
Option Explicit

Public Type tMAT2
    m00     As Double
    m01     As Double
    m10     As Double
    m11     As Double
End Type


Public Function matTranspose(m As tMAT2) As tMAT2
    matTranspose.m00 = m.m00
    matTranspose.m01 = m.m10    '
    matTranspose.m10 = m.m01    '
    matTranspose.m11 = m.m11

End Function

Public Function matMULv(m As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );

    matMULv.X = m.m00 * V.X + m.m01 * V.Y
    matMULv.Y = m.m10 * V.X + m.m11 * V.Y

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
