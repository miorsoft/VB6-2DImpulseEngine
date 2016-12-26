VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVEC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X       As Single
Public Y       As Single


Public Sub SETxy(vX As Single, vY As Single)
    X = vX
    Y = vY
End Sub

Public Sub SETv(V As clsVEC2)
    X = V.X
    Y = V.Y
End Sub

Public Sub Neg(ByRef V As clsVEC2)
    V.X = -V.X
    V.Y = -V.Y
End Sub

Public Function Negi(ByRef V As clsVEC2) As clsVEC2

    Negi = Neg(V)

End Function
