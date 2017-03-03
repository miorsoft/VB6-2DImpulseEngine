Attribute VB_Name = "modImpulseMath"
Option Explicit


Public Const DT As Single = 1 / 20 '20    '1 / 24  '1/20  '1 / 10   '1 / 60
Public Const Iterations As Long = 1 ' 2    ' 5 ' 20 '5   '10  '2    ' 4
Public Const DefDensity As Single = 1

Public Const PI As Single = 3.14159265358979
Public Const PI2 As Single = 6.28318530717959
Public Const PIh As Single = 1.5707963267949

Public Const EPSILON As Single = 0.000001    '0.0001
Public Const EPSILON_SQ As Single = EPSILON * EPSILON
Public Const BIAS_RELATIVE As Single = 0.9    '0.95
Public Const BIAS_ABSOLUTE As Single = 0.02    '0.01


Public Const PENETRATION_ALLOWANCE As Single = 0.01    ' 0.05    '0.1   ' 0.05
Public Const PENETRATION_CORRETION As Single = 0.4   '0.125   '0.4

Public Const MAX_VALUE As Single = 1E+32


Public Const GlobalSTATICFRICTION As Single = 0.25   '0.5
Public Const GlobalDYNAMICFRICTION As Single = 0.25   '0.3
Public Const GlobalRestitution As Single = 0.9    '0.8


Public GRAVITY As tVec2
Public RESTING As Single


Public Sub InitMATH()
    GRAVITY.x = 0
    GRAVITY.y = 0.01 / DT


    RESTING = Vec2LengthSq(Vec2MUL(GRAVITY, DT)) + EPSILON

    INVdt = 1 / DT
    INVdt2 = 1 / (DT * DT)
    
    DisplayRefreshPeriod = 2.5 / DT

End Sub

Public Function Equal(A As Single, B As Single) As Boolean
    If Abs(A - B) <= EPSILON Then Equal = True
End Function

Public Function Clamp(F As Single, T As Single, A As Single) As Single
    Clamp = A
    If Clamp < F Then
        Clamp = F
    ElseIf Clamp > T Then
        Clamp = T
    End If
End Function

Public Function rndFT(F As Single, T As Single) As Single
    rndFT = (T - F) * Rnd + F
End Function

'inline bool BiasGreaterThan( real a, real b )
'{
'  const real k_biasRelative = 0.95f;
'  const real k_biasAbsolute = 0.01f;
'  return a >= b * k_biasRelative + a * k_biasAbsolute;
'}
Public Function BiasGreaterThan(A As Single, B As Single) As Boolean
    BiasGreaterThan = (A >= (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function

Public Function gt(A As Single, B As Single) As Boolean
'return a >= b * BIAS_RELATIVE + a * BIAS_ABSOLUTE;
    gt = (A >= (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function


'********************** MATHS: ********************************


Public Function Min(A As Single, B As Single) As Single
    If A < B Then
        Min = A
    Else
        Min = B
    End If
End Function
Public Function Max(A As Single, B As Single) As Single
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function



