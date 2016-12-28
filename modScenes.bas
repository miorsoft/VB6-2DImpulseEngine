Attribute VB_Name = "modScenes"
Option Explicit

Public Sub CreateScene(Scene As Long)

    Dim I   As Long
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

            Add2PinsJoint 21, Vec2(30, 0), 22, Vec2(-30, 0), 80, 0.5


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
        Case 2


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


        Case 3

            'Floor
            CreateBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            BodySetStatic 1

            CreateBox Vec2(PicW * 0.1 + 20, PicH * 0.5), 50, 20
            AddPinJoint NBodies, Vec2(-20, 0), 40, 0.01, 0

            For I = 1 To 5
                CreateBox Vec2(PicW * 0.1 + 20 + 70 * I, PicH * 0.5), 50, 20
                'AddPinJoint NBodies, Vec2(-20, 0), 40, 0.01, 0
                Add2PinsJoint NBodies - 1, Vec2(20, 0), _
                              NBodies, Vec2(-20, 0), 30, 0.01, 0
            Next





        Case 4

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
                              NBodies, Vec2(-20, 0), 30, 0.01, 0
            Next

            BodySetStatic NBodies

        Case 5



            CreateBox Vec2(PicW * 0.2, PicH * 0.45), PicW * 0.5, 25, PI * 0.25
            BodySetStatic 1
            CreateBox Vec2(PicW * 0.8, PicH * 0.45), PicW * 0.5, 25, PI * 0.75
            BodySetStatic NBodies

            CreateBox Vec2(PicW * 0.5, PicH * 0.75), 58, 22
            AddPinJoint NBodies, Vec2(-25, 0), 0, 0.0006, 0.0006
            AddPinJoint NBodies, Vec2(25, 0), 0, 0.0006, 0.0006


        Case 6




For I = 1 To 9
            CreateBox Vec2((I - 0.5) * 82, PicH * 0.7), 58, 22
            AddPinJoint NBodies, Vec2(-25, 0), 0, 0.0006, 0.0006
            AddPinJoint NBodies, Vec2(25, 0), 0, 0.0006, 0.0006
Next


    End Select

End Sub

