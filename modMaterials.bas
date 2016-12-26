Attribute VB_Name = "modMaterials"
Option Explicit

Public Enum eMat
    mStone
    mMetal
    mGlass
    mWood
    mPlastic
    mRubber
    mFlesh
End Enum

Public Type tMaterial
    SF      As Double
    DF      As Double
    Restitution As Double
    Density As Double
End Type


Public MATERIAL() As tMaterial

Public Sub InitMaterials()

    ReDim MATERIAL(10)

    With MATERIAL(mStone)    'STONE
        .SF = 0.8
        .DF = 0.8
        .Restitution = 0.4
        .Density = 1
    End With

    With MATERIAL(mMetal)    'METAL
        .SF = 0.3
        .DF = 0.3
        .Restitution = 0.4
        .Density = 1
    End With

    With MATERIAL(mGlass)    'GLASS
        .SF = 0.6
        .DF = 0.6
        .Restitution = 0.5
        .Density = 1
    End With

    With MATERIAL(mWood)    'WOOD
        .SF = 0.6
        .DF = 0.6
        .Restitution = 0.5
    End With

    With MATERIAL(mFlesh)    'FLESH
        .SF = 0.9
        .DF = 0.9
        .Restitution = 0.3
        .Density = 1
    End With

    With MATERIAL(mPlastic)    'plastic
        .SF = 0.4
        .DF = 0.4
        .Restitution = 0.7
        .Density = 1
    End With

    With MATERIAL(mRubber)    'RUBBER
        .SF = 0.9
        .DF = 0.9
        .Restitution = 0.9
        .Density = 1
    End With



End Sub
