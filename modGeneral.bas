Attribute VB_Name = "basGen"
Option Explicit

Public Type rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type RGNDATAHEADER
        dwSize As Long
        iType As Long
        nCount As Long
        nRgnSize As Long
        rcBound As rect
End Type

Public Type rgnData
        rdh As RGNDATAHEADER
        Buffer As Byte
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Type PointDbl    'Point structure (in Doubles)
    x   As Double       'X-coordinate of point.
    y   As Double       'Y-coordinate of point.
End Type

Public Type LineDbl     'Line structure (in Doubles)
    ptStart As PointDbl 'Starting point (X, Y) on line.
    ptEnd   As PointDbl 'Ending point (X, Y) on line.
End Type


Public Const SRCCOPY = &HCC0020  ' (DWORD) dest = source
Public Const WINDING = 2
Public Const ALTERNATE = 1
Public Const RGN_XOR = 3
Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const COMPLEXREGION = 3
Public Const PT_MOVETO = &H6
Public Const PT_LINETO = &H2
Public Const PT_CLOSEFIGURE = &H1
Public Const PT_BEZIERTO = &H4
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const BS_SOLID = 0

Public Declare Function CloseFigure Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Public Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Public Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long 'POINTAPI) As Long
Public Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As rect) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

Public BrushWidth As Integer

Public Const gdPi   As Double = 3.14159265358979    'Pi
Public Const rads As Double = gdPi / 180

Public Sub RotatePoint(ptAxis As PointDbl, ptEnd As PointDbl, dDegrees As Double)

' Rotates ptEnd dDegrees from its
' current position, around ptAxis.

Dim dDX     As Double
Dim dDY     As Double
Dim dRads   As Double

    dRads = dDegrees * (gdPi / 180)
    dDX = ptEnd.x - ptAxis.x
    dDY = ptEnd.y - ptAxis.y
    ptEnd.x = ptAxis.x + ((dDX * Cos(dRads)) + (dDY * Sin(dRads)))
    ptEnd.y = ptAxis.y + -((dDX * Sin(dRads)) - (dDY * Cos(dRads)))
    
End Sub

Public Function LineAngleDegrees(Line1 As LineDbl) As Double

'Returns the angle of a line in degrees (see LineAngleRadians).

    LineAngleDegrees = RadiansToDegrees(LineAngleRadians(Line1))
    
End Function


Public Function LineAngleRadians(Line1 As LineDbl) As Double

'Calculates the angle(in radians) of a line from ptStart to ptEnd.

Dim dDeltaX As Double
Dim dDeltaY As Double
Dim dAngle  As Double

    dDeltaX = Line1.ptEnd.x - Line1.ptStart.x
    dDeltaY = Line1.ptEnd.y - Line1.ptStart.y
    
    If dDeltaX = 0 Then      'Vertical
        If dDeltaY < 0 Then
            dAngle = gdPi / 2
        Else
            dAngle = gdPi * 1.5
        End If
    
    ElseIf dDeltaY = 0 Then  'Horizontal
        If dDeltaX >= 0 Then
            dAngle = 0
        Else
            dAngle = gdPi
        End If
    
    Else    'Angled
        'Note: ++ = positive X, positive Y; +- = positive X, negative Y; etc.
        'On a true coordinate plane, Y increases as it move upward.
        'In VB coordinates, Y is reversed. It increases as it moves downward.
        
        'Calc for true Upper Right Quadrant (++) (For VB this is +-)
        dAngle = Atn(Abs(dDeltaY / dDeltaX))        'VB Upper Right (+-)
        
        'Correct for other 3 quadrants in VB coordinates (Reversed Y)
        If dDeltaX >= 0 And dDeltaY >= 0 Then       'VB Lower Right (++)
            dAngle = (gdPi * 2) - dAngle
            
        ElseIf dDeltaX < 0 And dDeltaY >= 0 Then    'VB Lower Left (-+)
            dAngle = gdPi + dAngle
            
        ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left (--)
            dAngle = gdPi - dAngle
            
        End If
        
    End If
    
    LineAngleRadians = dAngle
    
End Function

Public Function PerpLineCenter(Line1 As LineDbl) As LineDbl

'Returns a line perpendicular (90°) to Line1 using
'the center of Line1 as the first point.

Dim dDeltaX As Double
Dim dDeltaY As Double
Dim Line2   As LineDbl

    Line2.ptStart.x = (Line1.ptStart.x + Line1.ptEnd.x) / 2#
    Line2.ptStart.y = (Line1.ptStart.y + Line1.ptEnd.y) / 2#
    dDeltaX = Line2.ptStart.x - Line1.ptStart.x
    dDeltaY = Line2.ptStart.y - Line1.ptStart.y
    Line2.ptEnd.x = Line2.ptStart.x + -dDeltaY
    Line2.ptEnd.y = Line2.ptStart.y + dDeltaX
    
    PerpLineCenter = Line2
    
End Function

Public Function RadiansToDegrees(ByVal dRadians As Double) As Double

'Converts Radians to Degrees.

    RadiansToDegrees = dRadians * (180# / gdPi)
    
End Function

Public Function DegreesToRadians(ByVal dDegrees As Double) As Double

'Converts Degrees to Radians.

    DegreesToRadians = dDegrees * (gdPi / 180#)
    
End Function

Public Function Distance(ptStart As PointDbl, ptEnd As PointDbl) As Double

'Calculates the distance between 2 points.

    'Standard hypotenuse equation (c = Sqr(a^2 + b^2))
    Distance = Sqr(((ptEnd.x - ptStart.x) ^ 2) + ((ptEnd.y - ptStart.y) ^ 2))
    
End Function

Public Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
' Divides 2 numbers avoiding a "Division by zero" error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Public Function PointOnLine(ptStart As PointDbl, ptEnd As PointDbl, ByVal dDistance As Double) As PointDbl

'Returns a point on a line at dDistance from ptStart.
'This point need not be between ptStart and ptEnd.

Dim dDX     As Single
Dim dDY     As Single
Dim dLen    As Single
Dim dPct    As Single
    
    If dDistance > 1000000 Then
        dDistance = 1000000
    End If
        
    dLen = Distance(ptStart, ptEnd)
    
    If dLen > 0 Then
        dDX = ptEnd.x - ptStart.x
        dDY = ptEnd.y - ptStart.y
        dPct = Div(dDistance, dLen)
        PointOnLine.x = ptStart.x + (dDX * dPct)
        PointOnLine.y = ptStart.y + (dDY * dPct)
    Else
        PointOnLine.x = ptStart.x
        PointOnLine.y = ptStart.y
    End If
    
End Function

