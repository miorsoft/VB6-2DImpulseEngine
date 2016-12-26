Attribute VB_Name = "modRender2"
' NOTE: Enums evaluate to a Long
Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

' Quality mode constants
Public Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       ' Best performance
   QualityModeHigh = 2       ' Best rendering quality
End Enum

Public Enum GpUnit  ' aka Unit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum

Public Enum SmoothingMode
   SmoothingModeInvalid = QualityModeInvalid
   SmoothingModeDefault = QualityModeDefault
   SmoothingModeHighSpeed = QualityModeLow
   SmoothingModeHighQuality = QualityModeHigh
   SmoothingModeNone
   SmoothingModeAntiAlias
End Enum

Public Enum FillMode
   FillModeAlternate        ' 0
   FillModeWinding           ' 1
End Enum



Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal brush As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus

Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal brush As Long, ByVal argb As Long) As GpStatus


Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus

Public Declare Function GdipFillPolygon Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTF, ByVal count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipDrawLines Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus

'Private Const SmoothingModeHighQuality = 2
'Private Const UnitPixel = 2
Private Const PI As Double = 3.1415926


Dim initGDIP As Long
Dim objectFill As Long
Dim objectTraceWidth As Long
Dim gdiPLUSAntialising As Long

Dim bcolor1() As Long
Dim bcolor2() As Long
Dim poly() As POINTF

Private P1 As Object

Public Function gdipCOLOR(ByVal colRGB As Long) As Long

Dim gdipco As Long
ReDim co(7) As Byte

CopyMemory co(0), colRGB, 4
co(4) = co(2)
co(5) = co(1)
co(6) = co(0)
co(7) = 255
CopyMemory gdipco, co(4), 4

gdipCOLOR = gdipco


End Function


Public Sub gdipPolygon(P1 As Object, xx1, yy1, radius, ByVal pcolor As Long, ByVal pcolor2 As Long)

Dim GpInput As GdiplusStartupInput
Dim token As Long
Dim graphics As Long
Dim brush As Long
Dim pen As Long
Dim stat As Long

GpInput.GdiplusVersion = 1
If GdiplusStartup(token, GpInput, 0) <> Ok Then Exit Sub

GdipCreateFromHDC P1.hDC, graphics

If objectFill Then
    GdipCreateSolidFill gdipCOLOR(pcolor), brush
    If brush Then
        GdipFillPolygon graphics, brush, poly(0), UBound(poly) + 1, tracePolyFillMode
        GdipDeleteBrush brush
    End If
End If

If gdiPLUSAntialising Then GdipSetSmoothingMode graphics, SmoothingModeHighQuality
GdipCreatePen1 gdipCOLOR(pcolor2), objectTraceWidth, UnitPixel, pen
If pen Then
    GdipDrawLines graphics, pen, poly(0), UBound(poly) + 1
    GdipDeletePen pen
End If

GdipDeleteGraphics graphics

GdiplusShutdown (token)

End Sub
Public Sub gdipCircleFilled3(P1 As Object, xx1, yy1, xx2, yy2, radius, ByVal pcolor As Long, ByVal pcolor2 As Long)


Dim GpInput As GdiplusStartupInput
Dim token As Long
Dim graphics As Long
Dim brush As Long
Dim pen As Long
Dim stat As Long

GpInput.GdiplusVersion = 1
If GdiplusStartup(token, GpInput, 0) <> Ok Then Exit Sub


GdipCreateFromHDC P1.hDC, graphics


nbsteps = CLng(radius * 2) ' rough number.. should work not too bad large or small
ReDim ptl(nbsteps) As POINTF
For I = 0 To nbsteps
    anglerad = I / nbsteps * PI * 2
    ptl(I).X = xx1 + Cos(anglerad) * radius
    ptl(I).Y = yy1 - Sin(anglerad) * radius
Next

If objectFill Then
    GdipCreateSolidFill gdipCOLOR(pcolor), brush
    If brush Then
        GdipFillPolygon graphics, brush, ptl(0), UBound(ptl) + 1, tracePolyFillMode
        GdipDeleteBrush brush
    End If
End If

'If gdiPLUSAntialising Then GdipSetSmoothingMode graphics, SmoothingModeHighQuality
GdipCreatePen1 gdipCOLOR(pcolor2), objectTraceWidth, UnitPixel, pen
If pen Then
    GdipDrawLines graphics, pen, ptl(0), UBound(ptl) + 1
    ptl(0).X = xx1
    ptl(0).Y = yy1
    ptl(1).X = xx2
    ptl(1).Y = yy2
    GdipDrawLines graphics, pen, ptl(0), 2
    GdipDeletePen pen
End If

GdipDeleteGraphics graphics

GdiplusShutdown (token)

End Sub

Public Sub RENDER2(Mode As Long)
    
    Dim x1     As Long
    Dim y1     As Long
    Dim x2     As Long
    Dim y2     As Long

    Dim x1d    As Double
    Dim y1d    As Double
    Dim x2d    As Double
    Dim y2d    As Double


    Dim I      As Long
    Dim J      As Long

    'Dim p1 As Object
    'Set p1 = frmMain.PIC
    
    If initGDIP = 0 Then
    Set P1 = frmMain.PIC
    
        ReDim bcolor1(0) As Long
        ReDim bcolor2(0) As Long
        P1.ScaleMode = vbPixels
        objectFill = 1
        objectTraceWidth = 1 'use only 1 or 2 here
        initGDIP = 1
        
        GdipSetSmoothingMode graphics, SmoothingModeHighSpeed
        
    End If
    gdiPLUSAntialising = Mode
    
    
    If NofBodies > UBound(bcolor1) Then addMoreColors
    
    For I = 1 To NofBodies
        With Body(I)
             If .myType = eCircle Then
                 x1 = .Pos.X
                 y1 = .Pos.Y
                 x2 = x1 + .radius * Cos(-.orient)
                 y2 = y1 + .radius * Sin(-.orient)
                 gdipCircleFilled3 P1, x1, y1, x2, y2, .radius, bcolor1(I), bcolor2(I)
             Else
                 
                 ReDim poly(.nv) As POINTF
                 For J = 1 To .nv
                     x1 = .V(J).X + .Pos.X
                     y1 = .V(J).Y + .Pos.Y
                     poly(J).X = x1
                     poly(J).Y = y1
                 Next
                 poly(0).X = poly(.nv).X
                 poly(0).Y = poly(.nv).Y
                 gdipPolygon P1, x1, y1, .radius, bcolor1(I), bcolor2(I)
             End If

         End With

     Next
P1.CurrentX = 4
End Sub


Public Sub addMoreColors()

    lb = UBound(bcolor1) + 1
    ub = UBound(bcolor1) + 100
    ReDim Preserve bcolor1(ub) As Long
    ReDim Preserve bcolor2(ub) As Long
    For I = lb To ub
         bcolor1(I) = QBColor(1 + CInt(Rnd * 13))
         Do
            bcolor2(I) = QBColor(1 + CInt(Rnd * 14))
            If bcolor2(I) <> bcolor1(I) Then Exit Do
         Loop
         If I >= 4 And I <= 12 Then
            bcolor1(I) = vbBlue
            bcolor2(I) = vbWhite
            End If
         
    Next

End Sub

