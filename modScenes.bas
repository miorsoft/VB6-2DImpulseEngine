Attribute VB_Name = "modScenes"
Option Explicit

Public Sub CreateScene(Scene As Long)

    Dim I      As Long
    NBodies = 0
    NJ = 0

    Select Case Scene
    Case 0
        For I = 1 To 20
            CREATECircle Vec2(I * 55, 50), 5 + Rnd * (20), DefDensity
        Next

        'AddDistanceJoint 2, 3, 50
        'AddDistanceJoint 4, 5, 50
        'AddDistanceJoint 6, 7, 50
        'AddDistanceJoint 8, 9, 50
        'AddDistanceJoint 10, 11, 50

        CREATERandomPoly Vec2(300, 150), DefDensity
        CREATERandomPoly Vec2(350, 150), DefDensity

        Add2PinsJoint NBodies - 1, Vec2(30, 0), NBodies, Vec2(-30, 0), 80, 0.5


        For I = 20 + 1 To 20 + 9

            CREATECircle Vec2((I - 20 - 1) * 75, PicH + 40), 65, DefDensity

            BodySetStatic NBodies
        Next

        AddDistanceJoint 20 + 6, 5, 200



        '-----------ROPE
        CREATECircle Vec2(100, 50), 10, DefDensity
        BodySetStatic NBodies
        CREATECircle Vec2(100, 100), 10, DefDensity
        CREATECircle Vec2(100, 150), 10, DefDensity
        CREATECircle Vec2(100, 200), 10, DefDensity
        CREATECircle Vec2(100, 250), 10, DefDensity
        AddDistanceJoint NBodies, NBodies - 1, 50, 1, 0
        AddDistanceJoint NBodies - 1, NBodies - 2, 50, 1, 0
        AddDistanceJoint NBodies - 2, NBodies - 3, 50, 1, 0
        AddDistanceJoint NBodies - 3, NBodies - 4, 50, 1, 0


        CREATERandomPoly Vec2(500, 150), DefDensity

        AddPinJoint NBodies, Vec2(30, 0), 50, 0.1, 0.1
        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next

    Case 1

        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
        BodySetStatic 1

        '-----------ROPE
        CREATECircle Vec2(100, 50), 7, DefDensity
        BodySetStatic NBodies
        CREATECircle Vec2(100, 100), 7, DefDensity
        CREATECircle Vec2(100, 150), 7, DefDensity
        CREATECircle Vec2(100, 200), 7, DefDensity
        CREATECircle Vec2(100, 250), 7, DefDensity
        AddDistanceJoint NBodies, NBodies - 1, 50, 1, 0
        AddDistanceJoint NBodies - 1, NBodies - 2, 50, 1, 0
        AddDistanceJoint NBodies - 2, NBodies - 3, 50, 1, 0
        AddDistanceJoint NBodies - 3, NBodies - 4, 50, 1, 0

        CREATECircle Vec2(PicW * 0.75, PicH * 0.1), 7, DefDensity
        BodySetStatic NBodies
        CreateBox Vec2(PicW * 0.75, PicH * 0.2), 50, 20
        AddDistanceJoint NBodies, NBodies - 1, 100, 1, 0
        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next
    Case 2    '1PIN


        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
        BodySetStatic 1

        CreateBox Vec2(PicW * 0.5, PicH * 0.5), 100, 20
        AddPinJoint NBodies, Vec2(0, 0), 0

        CreateBox Vec2(PicW * 0.25, PicH * 0.5), 100, 20
        AddPinJoint NBodies, Vec2(-25, 0), 0

        CreateBox Vec2(PicW * 0.75, PicH * 0.4), 100, 20
        AddPinJoint NBodies, Vec2(-25, 0), 50


        CreateBox Vec2(PicW * 0.9, PicH * 0.1), 100, 20
        AddPinJoint NBodies, Vec2(-25, 0), 50, 0.005, 0.005
        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next

    Case 3    '"2 Pins Joints"

        'Floor
        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
        BodySetStatic 1

        CreateBox Vec2(PicW * 0.1 + 20, PicH * 0.5), 50, 20
        AddPinJoint NBodies, Vec2(-20, 0), 40, 0.01, 0

        For I = 1 To 5
            CreateBox Vec2(PicW * 0.1 + 20 + 70 * I, PicH * 0.5), 50, 20
            'AddPinJoint NBodies, Vec2(-20, 0), 40, 0.01, 0
            Add2PinsJoint NBodies - 1, Vec2(20, 0), _
                          NBodies, Vec2(-20, 0), 30, 1, 0
        Next

        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next



    Case 4    '"2 Pins Joints II

        'Floor
        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
        BodySetStatic 1

        'CreateBox 50, 20, Vec2(PicW * 0.1 + 20, PicH * 0.4)
        'AddPinJoint NBodies, Vec2(-20, 0), 40, 0.5, 0
        CreateBox Vec2(PicW * 0.05 + 70 * 0, PicH * 0.4), 50, 20
        BodySetStatic NBodies

        For I = 1 To 8
            CreateBox Vec2(PicW * 0.05 + 70 * I, PicH * 0.4), 50, 20
            Add2PinsJoint NBodies - 1, Vec2(20, 0), _
                          NBodies, Vec2(-20, 0), 30, 0.25, 0
        Next

        BodySetStatic NBodies
        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next
    Case 5    'Slope

        CreateBox Vec2(PicW * 0.2, PicH * 0.45), PicW * 0.5, 25, PI * 0.25
        BodySetStatic 1
        CreateBox Vec2(PicW * 0.8, PicH * 0.45), PicW * 0.5, 25, PI * 0.75
        BodySetStatic NBodies

        CreateBox Vec2(PicW * 0.5, PicH * 0.75), 58, 22
        AddPinJoint NBodies, Vec2(-25, 0), 0, 0.006, 0.006
        AddPinJoint NBodies, Vec2(25, 0), 0, 0.006, 0.006

        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next
    Case 6    'Gum Bridge

        For I = 1 To 8
            CreateBox Vec2((I - 0.5) * 82, PicH * 0.7), 58, 22
            AddPinJoint NBodies, Vec2(-25, 0), 0, 0.006, 0.006
            AddPinJoint NBodies, Vec2(25, 0), 0, 0.006, 0.006
        Next
        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next
    Case 7    '''' CAR

        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 1, 25
        BodySetStatic 1

        '        CreateBox Vec2(13, PicH * 0.5), 25, PicH * 1
        '        BodySetStatic 2
        '        CreateBox Vec2(PicW - 13, PicH * 0.5), 25, PicH * 1
        '        BodySetStatic 3


        CreateBox Vec2(100, 200), 80, 4
        CreateBox Vec2(140, 200), 80, 4
        Add2PinsJoint NBodies - 1, Vec2(35, 0), NBodies, Vec2(-35, 0), 0, 0.01, 0.01

        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next

        BodySetGroup NBodies - 1, BiggerGroup * 2    '=2
        BodySetGroup NBodies, BiggerGroup * 2    '=4
        BodySetCollideWith NBodies - 1, ALL - BiggerGroup    '=4
        BodySetCollideWith NBodies, ALL - BiggerGroup \ 2    '=2

        CreateBox Vec2(100, 200), 150, 40
        CREATECircle Vec2(100 - 50, 230), 20, DefDensity    'WHEEL
        BodySetGroup NBodies - 1, BiggerGroup * 2
        BodySetGroup NBodies, BiggerGroup * 2
        BodySetCollideWith NBodies - 1, ALL - BiggerGroup
        BodySetCollideWith NBodies, ALL - BiggerGroup \ 2
        CREATECircle Vec2(100 + 50, 230), 20, DefDensity    'WHEEL
        BodySetGroup NBodies, BiggerGroup
        BodySetCollideWith NBodies, ALL - BiggerGroup \ 4


        Add2PinsJoint NBodies - 1, Vec2(0, 0), NBodies - 2, Vec2(0, 0), Sqr(50 * 50 + 30 * 30)
        Add2PinsJoint NBodies, Vec2(0, 0), NBodies - 2, Vec2(0, 0), Sqr(50 * 50 + 30 * 30)

        Add2PinsJoint NBodies - 1, Vec2(0, 0), NBodies - 2, Vec2(-50, 0), 30, 0.01, 0.01
        Add2PinsJoint NBodies, Vec2(0, 0), NBodies - 2, Vec2(50, 0), 30, 0.01, 0.01



        AddRotorJoint NBodies - 1, Vec2(20, 0), 0.3


    Case 8    'newton cardle


        'Floor
        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
        BodySetStatic 1


        For I = 1 To 5

            CREATECircle Vec2(200 + I * 50, 50), 25, DefDensity
            AddPinJoint NBodies, Vec2(0, 0), 140, 0.1, 0.1

        Next


        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next

    Case 9    '''' CAR

        CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 1, 25
        BodySetStatic 1

        CreateBox Vec2(100, 200), 80, 33
        CreateBox Vec2(140, 200), 80, 33
        AddRotor2Joint NBodies - 1, Vec2(35, 0), NBodies, Vec2(-35, 0), 0.01, 0.01

        CreateBox Vec2(400, 200), 80, 33
        CreateBox Vec2(440, 200), 80, 33
        AddRotor2Joint NBodies - 1, Vec2(35, 0), NBodies, Vec2(-35, 0), 0.01, 0.01



        For I = 1 To NBodies
            BodySetGroup I, 1
            BodySetCollideWith I, ALL
        Next

        BodySetGroup NBodies - 1, BiggerGroup * 2    '=2
        BodySetGroup NBodies, BiggerGroup * 2    '=4
        BodySetCollideWith NBodies - 1, ALL - BiggerGroup    '=4
        BodySetCollideWith NBodies, ALL - BiggerGroup \ 2    '=2


        BodySetGroup NBodies - 3, BiggerGroup * 2    '=2
        BodySetGroup NBodies - 2, BiggerGroup * 2  '=4
        BodySetCollideWith NBodies - 3, ALL - BiggerGroup    '=4
        BodySetCollideWith NBodies - 2, ALL - BiggerGroup \ 2  '=2


    End Select

End Sub

