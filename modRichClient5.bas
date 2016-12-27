Attribute VB_Name = "modRichClient5"
Option Explicit


' After Draw ---> REFRESH:
'vbDRAW.Srf.DrawToDC PicHDC
'DoEvents



Public Srf As cCairoSurface, CC As cCairoContext    'Srf is similar to a DIB, the derived CC similar to a hDC
Attribute CC.VB_VarUserMemId = 1073741824

Public vbDRAW As cVBDraw
Attribute vbDRAW.VB_VarUserMemId = 1073741826
Public CONS As cConstructor
Attribute CONS.VB_VarUserMemId = 1610809344

Public PicHDC As Long
Attribute PicHDC.VB_VarUserMemId = 1073741828

Public Const JPGQuality As Long = 95


Public Sub InitRC()
' Set Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
' Set CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above



    Set vbDRAW = Cairo.CreateVBDrawingObject
    '    Set vbDRAW.Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
    Set vbDRAW.Srf = Cairo.CreateSurface(frmMain.PIC.Width, frmMain.PIC.Height, ImageSurface)         'size of our rendering-area in Pixels

    Set vbDRAW.CC = vbDRAW.Srf.CreateContext    'create a Drawing-Context from the PixelSurface above


    'vbDRAW.BindTo frmmain.PIC

    With vbDRAW

        .CC.AntiAlias = CAIRO_ANTIALIAS_GRAY

        '.CC.SetSourceSurface Srf
        .CC.SetLineCap CAIRO_LINE_CAP_ROUND
        .CC.SetLineJoin CAIRO_LINE_JOIN_ROUND


        .CC.SetLineWidth 1, True


        .CC.SelectFont "Courier New", 9, vbGreen

    End With

    PicHDC = frmMain.PIC.hDC

    '    frmmain.PIC.Cls
    '    frmmain.PIC.Height = 640    '480    '360    ' 480
    '    frmmain.PIC.Width = Int(frmmain.PIC.Height * 4 / 3)


End Sub

Public Sub UnloadRC()
    Set CC = Nothing
    Set Srf = Nothing
    Set vbDRAW = Nothing

    Set CONS = New cConstructor
    CONS.CleanupRichClientDll
End Sub

Public Sub CreateIntroFrames()
    vbDRAW.CC.SetSourceColor 0
    vbDRAW.CC.Paint

    StringOut "                               "
    StringOut "2D-Impulse-Engine:   (V" & Version & ")"
    StringOut "                               "
    StringOut "VB6 Port of Randy Gaul Impulse Engine.  "
    StringOut "                               "
    StringOut "Parameters:"
    StringOut "Delta time: " & DT
    StringOut "Iterations: " & Iterations
    StringOut "G. Static  Friction: " & GlobalSTATICFRICTION
    StringOut "G. Dynamic Friction: " & GlobalDYNAMICFRICTION
    StringOut "Restitution: " & GlobalRestitution
    StringOut "                               "
    StringOut "port to VB6 & Joints by MiorSoft"
    StringOut "                               "
    StringOut "                               "

End Sub
Public Sub CreateOuttroFrames()
    vbDRAW.CC.SetSourceColor 0
    vbDRAW.CC.Paint

    StringOut "                               "
    StringOut "                                Thanks for watching !"
    StringOut "                               "
    StringOut "                               "

End Sub
Private Sub StringOut(S As String)
    Dim I   As Double
    Const sstep As Double = 2
    Static y As Double
    Dim S2  As String

    Do
        S = S & " "
    Loop While (Len(S) - 1) Mod sstep <> 0

    S2 = S

    For I = 1 To Len(S2) Step sstep
        vbDRAW.CC.TextOut 10 + I * 7, y, Mid$(S, I, sstep)
        vbDRAW.Srf.WriteContentToJpgFile App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", JPGQuality
        Frame = Frame + 1
    Next
    y = y + 16



End Sub
Public Sub RENDERrc()
    Dim x1  As Long
    Dim y1  As Long
    Dim x2  As Long
    Dim y2  As Long

    Dim x1d As Double
    Dim y1d As Double
    Dim x2d As Double
    Dim y2d As Double


    Dim I   As Long
    Dim J   As Long
    Dim JJ  As Long


    vbDRAW.CC.SetSourceColor 0
    vbDRAW.CC.Paint
    vbDRAW.CC.SetLineWidth 1.25


    For I = 1 To NBodies

        With Body(I)

            If .myType = eCircle Then
                x1 = .Pos.X
                y1 = .Pos.y

                vbDRAW.CC.SetSourceColor .color
                vbDRAW.CC.Ellipse x1, y1, .radius * 2, .radius * 2
                vbDRAW.CC.Fill

                x2 = x1 + .radius * Cos(.orient)
                y2 = y1 + .radius * Sin(.orient)

                vbDRAW.CC.DrawLine x1, y1, x2, y2, , , 0, 0.5    '.color


            Else


                '                For J = 1 To .Nvertex
                '                    x1 = .tVertex(J).X + .Pos.X
                '                    y1 = .tVertex(J).Y + .Pos.Y
                '                    JJ = J + 1: If JJ > .Nvertex Then JJ = 1
                '                    x2 = .tVertex(JJ).X + .Pos.X
                '                    y2 = .tVertex(JJ).Y + .Pos.Y
                '                    '  FastLine pHDC, x1, y1, x2, y2, 1, .color
                '                    vbDRAW.CC.DrawLine x1, y1, x2, y2, , , .color
                '                Next
                '''' FILL
                vbDRAW.CC.SetSourceColor .color

                x1 = .tVertex(1).X + .Pos.X
                y1 = .tVertex(1).y + .Pos.y
                vbDRAW.CC.MoveTo x1, y1
                For J = 2 To .Nvertex
                    x1 = .tVertex(J).X + .Pos.X
                    y1 = .tVertex(J).y + .Pos.y
                    '                    JJ = J + 1: If JJ > .Nvertex Then JJ = 1
                    '                    x2 = .tVertex(JJ).X + .Pos.X
                    '                    y2 = .tVertex(JJ).Y + .Pos.Y
                    vbDRAW.CC.LineTo x1, y1
                Next
                vbDRAW.CC.Fill

                vbDRAW.CC.SetSourceColor 0, 0.5
                vbDRAW.CC.Ellipse .Pos.X, .Pos.y, 3, 3
                vbDRAW.CC.Fill




            End If

        End With

    Next


    '    ' DRAW Contact Points----------------------------------------------------------------------
    '    For I = 1 To NofContactMainFolds
    '        With Contacts(I)
    '            For J = 1 To .contactCount
    '                x1 = .contactsPTS(J).X
    '                y1 = .contactsPTS(J).Y
    '
    '
    '                x2 = x1 + .normal.X * (1 + .penetration * 25)
    '                y2 = y1 + .normal.Y * (1 + .penetration * 25)
    '
    '                vbDRAW.CC.DrawLine x1, y1, x2, y2, , 2, vbBlue, 0.5
    '            Next
    '        End With
    '    Next
    '    -------------------------------------------------------------------------------------------



    '    For I = 1 To NJ
    '        With Joints(I)
    '            x1 = Body(.bA).Pos.x
    '            y1 = Body(.bA).Pos.y
    '            x2 = Body(.bB).Pos.x
    '            y2 = Body(.bB).Pos.y
    '            FastLine pHDC, x1, y1, x2, y2, 1, vbWhite
    '        End With
    '    Next
    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType

                Case JointDistance
                    x1 = Body(.bA).Pos.X
                    y1 = Body(.bA).Pos.y
                    x2 = Body(.bB).Pos.X
                    y2 = Body(.bB).Pos.y
                    vbDRAW.CC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5


                Case Joint2Pins
                    x1 = Body(.bA).Pos.X + .tAnchA.X
                    y1 = Body(.bA).Pos.y + .tAnchA.y
                    x2 = Body(.bB).Pos.X + .tAnchB.X
                    y2 = Body(.bB).Pos.y + .tAnchB.y
                    vbDRAW.CC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5

                Case JointPin
                    x1 = .AnchB.X
                    y1 = .AnchB.y
                    x2 = Body(.bA).Pos.X + .tAnchA.X
                    y2 = Body(.bA).Pos.y + .tAnchA.y
                    vbDRAW.CC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5
                    vbDRAW.CC.SetSourceColor vbRed, 0.5
                    vbDRAW.CC.Ellipse x1, y1, 5, 5
                    vbDRAW.CC.Fill
            End Select

        End With
    Next

    vbDRAW.CC.TextOut 5, 5, "2D-Impulse-Engine by MiorSoft V" & Version & "   Objects: " & NBodies & "   Contacts: " & TotalNContacts

    DoEvents

    vbDRAW.Srf.DrawToDC PicHDC

End Sub

