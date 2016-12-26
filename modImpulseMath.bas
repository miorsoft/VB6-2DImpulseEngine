Attribute VB_Name = "modImpulseMath"
Option Explicit


Public Const DT As Double = 1 / 5   '1 / 60
Public Const Iterations As Long = 2 ' 4


Public Const PI As Double = 3.14159265358979
Public Const PI2 As Double = 6.28318530717959
Public Const PIh As Double = 1.5707963267949

Public Const EPSILON As Double = 0.0001
Public Const EPSILON_SQ As Double = EPSILON * EPSILON
Public Const BIAS_RELATIVE As Double = 0.95
Public Const BIAS_ABSOLUTE As Double = 0.01


Public Const PENETRATION_ALLOWANCE As Double = 0.05   ' 0.05
Public Const PENETRATION_CORRETION As Double = 0.3    '0.5

Public Const MAX_VALUE As Double = 1E+32

Public Const FLT_MAX As Double = 1E+32

Public Const GlobalSTATICFRICTION As Double = 0.75    '0.5
Public Const GlobalDYNAMICFRICTION As Double = 0.5    '0.3
Public Const GlobalRestitution As Double = 0.85    '0.8



Public GRAVITY As tVec2
Public RESTING As Double


Public Sub InitMATH()
    GRAVITY.X = 0
    GRAVITY.Y = 0.0625

    RESTING = Vec2LengthSq(Vec2MUL(GRAVITY, DT)) + EPSILON

    INVdt = 1 / DT

End Sub

Public Function Equal(A As Double, B As Double) As Boolean
    If Abs(A - B) <= EPSILON Then Equal = True
End Function

Public Function Clamp(F As Double, T As Double, A As Double) As Double
    Clamp = A
    If Clamp < F Then
        Clamp = F
    ElseIf Clamp > T Then
        Clamp = T
    End If
End Function

Public Function rndFT(F As Double, T As Double) As Double
    rndFT = (T - F) * Rnd + F
End Function



Public Function gt(A As Double, B As Double) As Boolean
'return a >= b * BIAS_RELATIVE + a * BIAS_ABSOLUTE;
    gt = (A >= (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function


'********************** MATHS: ********************************


Public Function Min(A As Double, B As Double) As Double
    Min = IIf(A < B, A, B)
End Function
Public Function Max(A As Double, B As Double) As Double
    Max = IIf(A > B, A, B)
End Function

'inline bool BiasGreaterThan( real a, real b )
'{
'  const real k_biasRelative = 0.95f;
'  const real k_biasAbsolute = 0.01f;
'  return a >= b * k_biasRelative + a * k_biasAbsolute;
'}
Public Function BiasGreaterThan(A As Double, B As Double) As Boolean
    BiasGreaterThan = (A > (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function

