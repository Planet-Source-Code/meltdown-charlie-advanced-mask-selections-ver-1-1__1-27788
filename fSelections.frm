VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSelections 
   BackColor       =   &H00C5A774&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection Tools"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog file 
      Left            =   9285
      Top             =   7740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Open Picture ..."
      Filter          =   "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg"
   End
   Begin VB.PictureBox canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   4  'Dash-Dot-Dot
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00800000&
      Height          =   7080
      Left            =   105
      Picture         =   "fSelections.frx":0000
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   631
      TabIndex        =   0
      Top             =   660
      Width           =   9525
      Begin VB.Timer tmrSelection 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   5235
         Top             =   4785
      End
      Begin VB.Timer tmrAnts 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   4515
         Top             =   4890
      End
      Begin VB.Shape shpBounds 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   3300
         Left            =   3195
         Top             =   870
         Visible         =   0   'False
         Width           =   4725
      End
      Begin VB.Line lnPoly 
         BorderStyle     =   5  'Dash-Dot-Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   49
         X2              =   171
         Y1              =   81
         Y2              =   141
      End
      Begin VB.Shape shpSelection 
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   1110
         Left            =   555
         Top             =   2550
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin VB.PictureBox picInvis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Left            =   390
      Picture         =   "fSelections.frx":96B7
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   410
      TabIndex        =   2
      Top             =   1245
      Visible         =   0   'False
      Width           =   6150
   End
   Begin VB.Image imgOpen 
      Height          =   480
      Index           =   2
      Left            =   8625
      Picture         =   "fSelections.frx":9721
      Top             =   5175
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgOpen 
      Height          =   465
      Index           =   1
      Left            =   8625
      Picture         =   "fSelections.frx":9C27
      Top             =   4620
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgOpen 
      Height          =   465
      Index           =   0
      Left            =   8625
      Picture         =   "fSelections.frx":A218
      Top             =   4080
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgFlood 
      Height          =   480
      Index           =   2
      Left            =   8640
      Picture         =   "fSelections.frx":A80F
      Top             =   3450
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgFlood 
      Height          =   480
      Index           =   1
      Left            =   8625
      Picture         =   "fSelections.frx":ACEB
      Top             =   2925
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgFlood 
      Height          =   480
      Index           =   0
      Left            =   8610
      Picture         =   "fSelections.frx":B270
      Top             =   2400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   20
      Left            =   8025
      Picture         =   "fSelections.frx":B7F5
      Top             =   5685
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   19
      Left            =   8010
      Picture         =   "fSelections.frx":BCAC
      Top             =   5130
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   18
      Left            =   8010
      Picture         =   "fSelections.frx":C19D
      Top             =   4575
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   16
      Left            =   7995
      Picture         =   "fSelections.frx":C61A
      Top             =   4065
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   14
      Left            =   7995
      Picture         =   "fSelections.frx":CA93
      Top             =   3525
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgNegative 
      Height          =   480
      Index           =   2
      Left            =   7905
      Picture         =   "fSelections.frx":CF03
      Top             =   2895
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgAdditive 
      Height          =   480
      Index           =   2
      Left            =   7920
      Picture         =   "fSelections.frx":D36A
      Top             =   2355
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgRedo 
      Height          =   480
      Index           =   3
      Left            =   8505
      Picture         =   "fSelections.frx":D7E3
      Top             =   1635
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgUndoPics 
      Height          =   480
      Index           =   3
      Left            =   8475
      Picture         =   "fSelections.frx":DCF8
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgMove 
      Height          =   480
      Index           =   2
      Left            =   8070
      Picture         =   "fSelections.frx":E244
      Top             =   6255
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgMove 
      Height          =   480
      Index           =   1
      Left            =   7515
      Picture         =   "fSelections.frx":E6E4
      Top             =   6285
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgMove 
      Height          =   480
      Index           =   0
      Left            =   6915
      Picture         =   "fSelections.frx":EC27
      Top             =   6255
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image btnMove 
      Height          =   480
      Left            =   9045
      Picture         =   "fSelections.frx":F16A
      ToolTipText     =   "Click to set for moving the selection"
      Top             =   75
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   13
      Left            =   7470
      Picture         =   "fSelections.frx":F6AD
      Top             =   5685
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   12
      Left            =   7440
      Picture         =   "fSelections.frx":FC14
      Top             =   5115
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   11
      Left            =   7440
      Picture         =   "fSelections.frx":101B1
      Top             =   4545
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   9
      Left            =   7440
      Picture         =   "fSelections.frx":106E8
      Top             =   4035
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   7
      Left            =   7455
      Picture         =   "fSelections.frx":10C11
      Top             =   3510
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   6
      Left            =   6900
      Picture         =   "fSelections.frx":1113D
      Top             =   5685
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   5
      Left            =   6855
      Picture         =   "fSelections.frx":116A4
      Top             =   5100
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   4
      Left            =   6870
      Picture         =   "fSelections.frx":11C41
      Top             =   4545
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   2
      Left            =   6855
      Picture         =   "fSelections.frx":12178
      Top             =   4005
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgTools 
      Height          =   480
      Index           =   0
      Left            =   6840
      Picture         =   "fSelections.frx":126A1
      Top             =   3495
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgNegative 
      Height          =   480
      Index           =   1
      Left            =   7320
      Picture         =   "fSelections.frx":12BCD
      Top             =   2865
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgNegative 
      Height          =   480
      Index           =   0
      Left            =   6780
      Picture         =   "fSelections.frx":130E2
      Top             =   2865
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgAdditive 
      Height          =   480
      Index           =   1
      Left            =   7290
      Picture         =   "fSelections.frx":135F7
      Top             =   2340
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgAdditive 
      Height          =   480
      Index           =   0
      Left            =   6720
      Picture         =   "fSelections.frx":13B19
      Top             =   2340
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgRedo 
      Height          =   465
      Index           =   2
      Left            =   7860
      Picture         =   "fSelections.frx":1403B
      Top             =   1635
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgRedo 
      Height          =   465
      Index           =   1
      Left            =   7260
      Picture         =   "fSelections.frx":14637
      Top             =   1650
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgRedo 
      Height          =   465
      Index           =   0
      Left            =   6690
      Picture         =   "fSelections.frx":14C32
      Top             =   1620
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgUndoPics 
      Height          =   480
      Index           =   2
      Left            =   7875
      Picture         =   "fSelections.frx":1521F
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgUndoPics 
      Height          =   480
      Index           =   1
      Left            =   7245
      Picture         =   "fSelections.frx":15816
      Top             =   1065
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgUndoPics 
      Height          =   480
      Index           =   0
      Left            =   6660
      Picture         =   "fSelections.frx":15E07
      Top             =   1065
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image btnFlood 
      Height          =   480
      Left            =   6300
      Picture         =   "fSelections.frx":16402
      ToolTipText     =   "Click for temporary fill"
      Top             =   90
      Width           =   540
   End
   Begin VB.Image btnMode 
      Height          =   480
      Index           =   1
      Left            =   3375
      Picture         =   "fSelections.frx":16987
      ToolTipText     =   "Subtractive Mask Mode"
      Top             =   75
      Width           =   540
   End
   Begin VB.Image btnMode 
      Height          =   480
      Index           =   0
      Left            =   2850
      Picture         =   "fSelections.frx":16E9C
      ToolTipText     =   "Additive Mask Mode"
      Top             =   75
      Width           =   540
   End
   Begin VB.Image imgAnts 
      Height          =   120
      Index           =   1
      Left            =   7125
      Picture         =   "fSelections.frx":17315
      Top             =   6840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgAnts 
      Height          =   120
      Index           =   0
      Left            =   6960
      Picture         =   "fSelections.frx":1737F
      Top             =   6825
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   7860
      Width           =   4050
   End
   Begin VB.Image btnPicture 
      Height          =   465
      Left            =   7035
      Picture         =   "fSelections.frx":173E9
      ToolTipText     =   "Click to load a picture"
      Top             =   90
      Width           =   555
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   6
      Left            =   2070
      Picture         =   "fSelections.frx":179E0
      ToolTipText     =   "Magic Wand"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   5
      Left            =   1545
      Picture         =   "fSelections.frx":17F47
      ToolTipText     =   "Brush Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   4
      Left            =   1050
      Picture         =   "fSelections.frx":184E4
      ToolTipText     =   "Polygon Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   2
      Left            =   525
      Picture         =   "fSelections.frx":18A1B
      ToolTipText     =   "Elliptical Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   1
      Left            =   585
      Picture         =   "fSelections.frx":18F44
      ToolTipText     =   "Square Mask"
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   0
      Left            =   15
      Picture         =   "fSelections.frx":19467
      Tag             =   "Presed"
      ToolTipText     =   "Rectangle Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnRedo 
      Enabled         =   0   'False
      Height          =   465
      Left            =   8370
      Picture         =   "fSelections.frx":198D7
      ToolTipText     =   "Redo"
      Top             =   90
      Width           =   555
   End
   Begin VB.Image btnUndo 
      Enabled         =   0   'False
      Height          =   480
      Left            =   7785
      Picture         =   "fSelections.frx":19ED3
      ToolTipText     =   "Undo"
      Top             =   90
      Width           =   540
   End
End
Attribute VB_Name = "fSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================================================================
' ===================================================================================
' Selections test code ver 1.1

' by M Ferris - Intact Interactive Software.
' url       : http://www.intactinteractive.com
' email     : mferris@zfree.co.nz
' copyright : M Ferris 2001
'
' ===================================================================================
' ===================================================================================
'
' This sample code has been given to the development community by Intact Interactive
' software. The code is not finished commercial level stuff so you will find bugs and
' quirks in it I'm sure. Neither is the code a complete application, it is merely a
' testbed for the functionalily being tested!
' So you use this code on an as is basis with no warranty given or implied by Intact
' Interacive. If you are cool with this then go for it, if not, then delete this.
'
' Also remember that this is not a one way street, I have to survive as a developer
' as well, so obviously I am not giving away all my knowledge. And I would appreciate
' it if you gave me feedback, vote for me at PSC or wherever you found this posted
' and, of course, download the eval versions of my commercial stuff, if you like what
' you see you might even consider buying some of it ...
'
' A more complete example of this sample source code may be found at
' http://www.intactinteractive.com so I encourage you to visit our site and check it
' out as well as our other source code samples ...
'
' Some notes about this version.
'
' Features :
'
'           - multiple undo/redo levels.
'           - progressive mask building, i.e. you can add/subtract in multiple steps
'           - add/subtract mode of mask building
'           - rectangle,ellipse and polygon selection tools functional
'           - paint brush mask selection tool
'
' How to use :
'              - Click on a selection tool to use it.
'              - For rectangle and ellipse tool - click on canvas and drag to
'              define your selection.
'              - For polygon tool - click to start a line, move to endpoint and
'              click again to end the line and begin next line, to close the shape
'              click near to the first lines starting point.
'              - Click the plus button for additive mask, and the minus for
'              subtractive mask mode.
'              - When undo is possible, the undo button will appear, click it to
'              undo to the last level.
'              - When redo is possible, the redo button will appear, click it to
'              redo to the last level.
'              - To fill the selection with red paint - click the paint bucket button.
'
' To Do :
'           - implement the wand selection tool
'           - modify code to work as an activex class (dll) and activex control (ocx)
'           - combine code with gradients sample to show implementation of area
'           - add combine mode for xor - i.e. removes the unions
'           - add combine mode for and - i.e. keeps only the union
'           - add mask transparency functionality
'           - add mask feather functionality
'
' ===================================================================================
' this form illustrates the use of regions to create selection areas for a paint
' program, it shows how to use CombineRgn() to progressively add to a mask region,
' how to use GetRegionData() to implement an undo/redo feature, and generally shows
' how the use of regions can make the implementation of complex selections possible
' in VB !
'
' the selection tools included in this sample are :
'
'                                   rectangular selection
'                                   elliptical selection
'                                   magic wand selection (yet to be implemented)
'                                   ploygonal selection
'                                   paint brush selection (yet to be implemented)
'
' also an undo function is available to allow the mask to be undone and redone up
' to all levels of the mask creation ...
' ===================================================================================

Option Explicit

' this is a container for an array of region data
Private Type rgns
    data() As Byte
    length As Long
End Type

Private bMoving As Boolean
Private bMoveSelection As Boolean
' our undo and redo arrays and counters
Private undoRgn() As rgns
Private numUndos As Integer
Private RedoRGN() As rgns
Private NumRedos As Integer
' the master region which we use for our selection
Private MasterRgn As Long
Private currgn As Long
' mouse tracking vars
Private xOfs As Single
Private yOfs As Single
Private oldx As Single
Private oldy As Single
Private xOrig As Single
Private yOrig As Single
Private xDiff As Integer
Private yDiff As Integer
' flag to indicate selection has changed
Private bSelectionChanged As Boolean
' index var for the marching ants brushes
Private outlineType As Integer
' the marching ants brushes
Private antsBrush(1) As Integer
' a temporary region - used with paintbrush selection
Private rgnTmp As Long
' the selection tool currently active
Private CurrentSelectionTool As Integer
' the current mode for mask combination - i.e additive or subtractive
Private CombineMode As Long
' array of original button x positions
Private btnXPositions() As Integer

Private Sub AddLine(width As Integer, x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    Dim ln As LineDbl
    Dim ln2 As LineDbl
    Dim pt As PointDbl
    Dim pt2 As PointDbl
    Dim d As Double
    Dim pts(0 To 10) As POINTAPI
    Dim xDiff1 As Integer
    Dim yDiff1 As Integer
    Dim xDiff2 As Integer
    Dim yDiff2 As Integer
    Dim rgn As Long
    Dim rotn As Double
    
    ' get a line perpendicular to the line ...
    ln.ptStart.X = x1
    ln.ptStart.Y = y1
    ln.ptEnd.X = x2
    ln.ptEnd.Y = y2
    ln2 = PerpLineCenter(ln)
    ' get a point either side of the line at half the brush width ...
    pt = PointOnLine(ln2.ptStart, ln2.ptEnd, width \ 2)
    pt2 = PointOnLine(ln2.ptStart, ln2.ptEnd, -(width \ 2))
    ' calculate the difference offsets for the four points of a quadrilateral polygon
    xDiff1 = ln2.ptStart.X - pt.X
    xDiff2 = ln2.ptStart.X - pt2.X
    yDiff1 = ln2.ptStart.Y - pt.Y
    yDiff2 = ln2.ptStart.Y - pt2.Y
    ' now set up the four "corner" points for our brush stroke ...
    pts(0).X = x1 + xDiff1
    pts(0).Y = y1 + yDiff1
    pts(1).X = x2 + xDiff1
    pts(1).Y = y2 + yDiff1
    pts(5).X = x2 + xDiff2
    pts(5).Y = y2 + yDiff2
    pts(6).X = x1 + xDiff2
    pts(6).Y = y1 + yDiff2
    ' now we attempt to make the ends rounded ...
    ln.ptStart.X = x2
    ln.ptStart.Y = y2
    ln.ptEnd.X = pts(1).X
    ln.ptEnd.Y = pts(1).Y
    ' get the angle of the perpendicular line ...
    d = LineAngleDegrees(ln)
    If pts(1).Y < pts(5).Y Then rotn = 45 Else rotn = -45
    ' and rotate by 45/-45 degrees to create a "rounded" end ...
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(2).X = ln.ptEnd.X
    pts(2).Y = ln.ptEnd.Y
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(3).X = ln.ptEnd.X
    pts(3).Y = ln.ptEnd.Y
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(4).X = ln.ptEnd.X
    pts(4).Y = ln.ptEnd.Y
    ln.ptStart.X = x1
    ln.ptStart.Y = y1
    ln.ptEnd.X = pts(6).X
    ln.ptEnd.Y = pts(6).Y
    ' get the angle of the perpendicular line ...
    d = LineAngleDegrees(ln)
    If pts(6).Y > pts(0).Y Then rotn = 45 Else rotn = -45
    ' and rotate by 45/-45 degrees to create a "rounded" end ...
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(7).X = ln.ptEnd.X
    pts(7).Y = ln.ptEnd.Y
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(8).X = ln.ptEnd.X
    pts(8).Y = ln.ptEnd.Y
    RotatePoint ln.ptStart, ln.ptEnd, rotn
    pts(9).X = ln.ptEnd.X
    pts(9).Y = ln.ptEnd.Y
    pts(10).X = pts(0).X
    pts(10).Y = pts(0).Y
    ' make a region from this ...
    rgn = CreatePolygonRgn(pts(0), 11, WINDING)
    ' now check to ensure a master region exists, if not make one for the combine ...
    If MasterRgn = 0 Then
        MasterRgn = CreateRectRgn(0, 0, 1, 1)
        CombineRgn MasterRgn, MasterRgn, rgn, RGN_COPY
        DeleteObject rgn
    Else
        CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
        DeleteObject rgn
    End If
    ' signal that the selection has changed
    bSelectionChanged = True
End Sub

Private Sub AddToUndo(rgn As Long)
    numUndos = numUndos + 1
    ReDim Preserve undoRgn(1 To numUndos)
    undoRgn(numUndos).length = GetRegionData(rgn, undoRgn(numUndos).length, ByVal 0&)
    ReDim Preserve undoRgn(numUndos).data(undoRgn(numUndos).length)
    GetRegionData rgn, undoRgn(numUndos).length, undoRgn(numUndos).data(0)
    btnUndo.Enabled = True
    btnUndo = imgUndoPics(0)
End Sub

Private Sub AddToRedo(rgn As Long)
    NumRedos = NumRedos + 1
    ReDim Preserve RedoRGN(1 To NumRedos)
    RedoRGN(NumRedos).length = GetRegionData(rgn, RedoRGN(NumRedos).length, ByVal 0&)
    ReDim Preserve RedoRGN(NumRedos).data(RedoRGN(NumRedos).length)
    GetRegionData rgn, RedoRGN(NumRedos).length, RedoRGN(NumRedos).data(0)
    btnRedo.Enabled = True
    btnRedo = imgRedo(0)
End Sub

Private Sub btnFlood_Click()
    Dim hbr As Long
    Dim lbr As LOGBRUSH
    
    ' this only floods the current selection and doesn't work completely yet -
    ' i.e. when you add/subtract - undo/redo or move then mask the flood disappears
    ' at present (this is by design for now, as it only serves to show how the
    ' selection is usable) ...
    lbr.lbColor = RGB(255, 0, 0)
    lbr.lbStyle = BS_SOLID
    hbr = CreateBrushIndirect(lbr)
    FillRgn canvas.hdc, MasterRgn, hbr
    DeleteObject hbr
End Sub

Private Sub btnFlood_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnFlood = imgFlood(2)
End Sub

Private Sub btnFlood_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And 1 Then Exit Sub
    btnFlood = imgFlood(1)
End Sub

Private Sub btnMode_Click(Index As Integer)
    ' change the combine mode to either additive or subtractive
    If Index = 0 Then
        CombineMode = RGN_OR
    Else
        CombineMode = RGN_DIFF
    End If
End Sub

Private Sub btnMode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        btnMode(Index) = imgAdditive(2)
        btnMode(1) = imgNegative(0)
    Else
        btnMode(Index) = imgNegative(2)
        btnMode(0) = imgAdditive(0)
    End If
End Sub

Private Sub btnMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        If CombineMode <> RGN_OR Then btnMode(Index) = imgAdditive(1)
        If CombineMode <> RGN_DIFF Then btnMode(1) = imgNegative(0)
    Else
        If CombineMode <> RGN_DIFF Then btnMode(Index) = imgNegative(1)
        If CombineMode <> RGN_OR Then btnMode(0) = imgAdditive(0)
    End If
End Sub

Private Sub btnMove_Click()
    ' set move mode ...
    If bMoveSelection Then
        bMoveSelection = False
        btnMove = imgMove(0)
        canvas.MousePointer = 0
    Else
        bMoveSelection = True
        btnMove = imgMove(2)
        canvas.MousePointer = 15
    End If
End Sub

Private Sub btnMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bMoveSelection Then btnMove = imgMove(1)
End Sub

Private Sub btnPicture_Click()
    file.ShowOpen
On Error Resume Next
    If file.FileName <> "" Then
        canvas.Picture = LoadPicture(file.FileName)
    End If
End Sub

Private Sub btnPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnPicture = imgOpen(2)
End Sub

Private Sub btnPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And 1 Then Exit Sub
    btnPicture = imgOpen(1)
End Sub

Private Sub btnRedo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btnRedo.Enabled Then
        btnRedo = imgRedo(1)
    End If
    If btnUndo.Enabled Then btnUndo = imgUndoPics(0)
End Sub

Private Sub btnSelectionTools_Click(Index As Integer)
    If Index = 5 Then
        ' set the brush width ...
        fBrushSettings.Show 1
    End If
    
    If Index > 5 Then
        MsgBox "Visit http://www.intactinteractive.com for a more complete example of this code!", , "Not Implemented yet!"
    End If
End Sub

Private Sub btnRedo_Click()
    Dim idx As Integer
    
    idx = UBound(RedoRGN)
    ' add the existing region to the undo array ...
    AddToUndo MasterRgn
    ' recreate the mask from the redo region data ...
    DeleteObject MasterRgn
    MasterRgn = ExtCreateRegion(ByVal 0&, RedoRGN(idx).length, RedoRGN(idx).data(0))
    bSelectionChanged = True
    ' now trim off the region from the redo array ...
    Erase RedoRGN(idx).data
    If idx - 1 <= 0 Then
        ReDim RedoRGN(1 To 1)
        Erase RedoRGN(1).data
        NumRedos = 0
        btnRedo.Enabled = False
        btnRedo = imgRedo(2)
    Else
        idx = idx - 1
        ReDim Preserve RedoRGN(1 To idx)
        NumRedos = idx
    End If
End Sub

Private Sub btnSelectionTools_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    CurrentSelectionTool = Index
On Error Resume Next
    For i = 0 To btnSelectionTools.UBound
        If i = Index Then
            btnSelectionTools(i) = imgTools(i + 14)
        Else
            btnSelectionTools(i) = imgTools(i)
        End If
    Next i
End Sub

Private Sub btnSelectionTools_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
On Error Resume Next
    For i = 0 To btnSelectionTools.UBound
        If i = Index Then
            If i <> CurrentSelectionTool Then
                btnSelectionTools(i) = imgTools(i + 7)
            End If
        Else
            If i <> CurrentSelectionTool Then
                btnSelectionTools(i) = imgTools(i)
            End If
        End If
    Next i
End Sub

Private Sub btnUndo_Click()
    Dim idx As Integer
    
    idx = UBound(undoRgn)
    
    ' copy the last undo region into the redo array ...
    AddToRedo MasterRgn
    ' recreate the mask from the undo region data ...
    DeleteObject MasterRgn
    MasterRgn = ExtCreateRegion(ByVal 0&, undoRgn(idx).length, undoRgn(idx).data(0))
    bSelectionChanged = True
    ' now trim off the region from the undo array ...
    Erase undoRgn(idx).data
    If idx - 1 <= 0 Then
        ReDim undoRgn(1 To 1)
        Erase undoRgn(1).data
        numUndos = 0
        btnUndo.Enabled = False
        btnUndo = imgUndoPics(2)
    Else
        idx = idx - 1
        ReDim Preserve undoRgn(1 To idx)
        numUndos = idx
    End If
End Sub

Private Sub btnUndo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btnUndo.Enabled Then
        btnUndo = imgUndoPics(1)
    End If
    If btnRedo.Enabled Then btnRedo = imgRedo(0)
End Sub

Private Sub canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    oldx = X
    oldy = Y
    xOrig = X
    yOrig = Y
    If bMoveSelection Then
        If MasterRgn = 0 Then Exit Sub
        If PtInRegion(MasterRgn, X, Y) Then
            Dim rc As rect
            
            ' get the region bounds
            GetRgnBox MasterRgn, rc
            ' show the bounds rectangle
            shpBounds.Move rc.top, rc.left, rc.right - rc.left, rc.bottom - rc.top
            canvas.Refresh
            shpBounds.Visible = True
            xOfs = X - rc.left
            yOfs = Y - rc.top
            ' set moving to true
            bMoving = True
        End If
        Exit Sub
    End If
    ' set the beginning of mask selection depending on the type of mask tool ...
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3 ' rect,square,circle,ellipse ...
            shpSelection.left = X
            shpSelection.top = Y
            shpSelection.width = 1
            shpSelection.Height = 1
            shpSelection.Visible = True
            tmrAnts.Enabled = True
            shpSelection.Shape = CurrentSelectionTool
            
        Case 4 ' polygon
            ' add another line to the region and continue building the polygon ...
            If lnPoly(0).Visible Then
                Load lnPoly(lnPoly.Count)
                lnPoly(lnPoly.UBound).x1 = lnPoly(lnPoly.UBound - 1).x2
                lnPoly(lnPoly.UBound).y1 = lnPoly(lnPoly.UBound - 1).y2
                lnPoly(lnPoly.UBound).x2 = X
                lnPoly(lnPoly.UBound).y2 = Y
                lnPoly(lnPoly.UBound).Visible = True
            Else
                lnPoly(0).x1 = X
                lnPoly(0).y1 = Y
                lnPoly(0).x2 = X
                lnPoly(0).y2 = Y
                lnPoly(0).Visible = True
                tmrAnts.Enabled = True
                CurrentSelectionTool = 4
            End If
        Case 5 ' brush mask
            tmrAnts.Enabled = True
        Case 6 ' magic wand
    End Select
End Sub

Private Sub canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' handle the movement of the selection tool ...
    Dim pt As POINTAPI
    Dim rgn As Long
    Dim res As Long
    Dim pts() As POINTAPI
    Dim types() As Byte
    Dim sz As Long
    
    If bMoving Then
        shpBounds.Move X - xOfs, Y - yOfs
        Exit Sub
    End If
    If bMoveSelection Then Exit Sub
    xDiff = X - oldx
    yDiff = Y - oldy
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3 ' rect,square,circle,ellipse ...
            If Button = 0 Then Exit Sub
            If X < xOrig Then
                shpSelection.left = X
                shpSelection.width = xOrig - X
            Else
                shpSelection.width = shpSelection.width + xDiff
            End If
            If Y < yOrig Then
                shpSelection.top = Y
                shpSelection.Height = yOrig - Y
            Else
                shpSelection.Height = shpSelection.Height + yDiff
            End If
        Case 4 ' polygon
            lnPoly(lnPoly.UBound).x2 = X
            lnPoly(lnPoly.UBound).y2 = Y
            canvas.Refresh
        Case 5 ' brush mask
            ' we are going to do a bitmap scan for black pixels and build a mask
            ' progressively ala irregular window creation routines ...
            If Button = 0 Then Exit Sub
            If oldx <> X Or oldy <> Y Then
                AddLine BrushWidth, oldx, oldy, X, Y
            End If
        Case 6 ' magic wand
    End Select
    oldx = X
    oldy = Y
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rgnRects() As rect
    Dim rgnSize As Integer
    Dim rc As rect
    Dim sz As Integer
    Dim pts() As POINTAPI
    Dim numPoints As Integer
    Dim i As Integer
    Dim res As Long
    Dim rgn As Long
    
    If bMoving Then
        ' hide the bounding box
        shpBounds.Visible = False
        ' offset the region to the new position
        GetRgnBox MasterRgn, rc
        OffsetRgn MasterRgn, shpBounds.left - rc.left, shpBounds.top - rc.top
        bMoving = False
        bSelectionChanged = True
        AddToUndo MasterRgn
        Exit Sub
    End If
    If bMoveSelection Then MsgBox "Nothing to move!": Exit Sub
    shpSelection.Visible = False
    
    Select Case CurrentSelectionTool
        Case 0, 1 ' rectangle, square
            tmrAnts.Enabled = False
            rgn = CreateRectRgn(shpSelection.left, _
                                shpSelection.top, _
                                shpSelection.left + shpSelection.width, _
                                shpSelection.top + shpSelection.Height)
            If MasterRgn = 0 Then
                MasterRgn = CreateRectRgn(shpSelection.left, _
                                            shpSelection.top, _
                                            shpSelection.left + shpSelection.width, _
                                            shpSelection.top + shpSelection.Height)
                bSelectionChanged = True
            Else
                AddToUndo MasterRgn
                CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                bSelectionChanged = True
            End If
        Case 2, 3 ' circle, ellipse
            tmrAnts.Enabled = False
            rgn = CreateEllipticRgn(shpSelection.left, _
                                    shpSelection.top, _
                                    shpSelection.left + shpSelection.width, _
                                    shpSelection.top + shpSelection.Height)
                 
           If MasterRgn = 0 Then
                MasterRgn = CreateEllipticRgn(shpSelection.left, _
                                            shpSelection.top, _
                                            shpSelection.left + shpSelection.width, _
                                            shpSelection.top + shpSelection.Height)
                
                bSelectionChanged = True
            Else
                ' save the old master region into a data array
                AddToUndo MasterRgn
                CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                
                bSelectionChanged = True
            End If
                                    
       Case 4 ' polygon
            ' first see if the polygon is closing ...
            If lnPoly.UBound <> 0 And X > lnPoly(0).x1 - 5 And X < lnPoly(0).x1 + 5 And Y > lnPoly(0).y1 - 5 And Y < lnPoly(0).y1 + 5 Then
                ' and if it is do the following ...
                tmrAnts.Enabled = False
                ' build a polygon region based on the lines ...
                numPoints = lnPoly.Count + 1
                ReDim pts(numPoints)
                For i = 0 To lnPoly.UBound
                    pts(i).X = lnPoly(i).x1
                    pts(i).Y = lnPoly(i).y1
                Next i
                pts(numPoints - 1).X = lnPoly(0).x1
                pts(numPoints - 1).Y = lnPoly(0).y1
                rgn = CreatePolygonRgn(pts(0), numPoints, WINDING)
                ' clean up the existing lines ...
                For i = 1 To lnPoly.UBound
                    Unload lnPoly(i)
                Next i
                lnPoly(0).Visible = False
                ' combine it to the current master region ...
                If MasterRgn = 0 Then
                    MasterRgn = CreatePolygonRgn(pts(0), numPoints, WINDING)
                    bSelectionChanged = True
                Else
                    AddToUndo MasterRgn
                    CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                    bSelectionChanged = True
                End If
            End If
        Case 5 ' brush mask
            AddToUndo MasterRgn
            tmrAnts.Enabled = False
        Case 6 ' magic wand
    End Select
    tmrSelection.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ReDim btnXPositions(0 To btnSelectionTools.UBound)
    For i = 0 To btnSelectionTools.UBound
On Error Resume Next
        btnXPositions(i) = btnSelectionTools(i).left
    Next i
    CurrentSelectionTool = 0
    antsBrush(0) = CreatePatternBrush(imgAnts(0).Picture.Handle)
    antsBrush(1) = CreatePatternBrush(imgAnts(1).Picture.Handle)
    canvas.Picture = canvas.Image
    canvas.Refresh
    CombineMode = RGN_OR
    numUndos = 0
    NumRedos = 0
    picInvis.width = canvas.width
    picInvis.Height = canvas.Height
    picInvis.left = canvas.left
    picInvis.top = canvas.top
    picInvis.Refresh
    picInvis.Picture = picInvis.Image
    bMoveSelection = False
    bMoving = False
    BrushWidth = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
On Error Resume Next
    For i = 0 To btnSelectionTools.UBound
        If i <> CurrentSelectionTool Then
            btnSelectionTools(i) = imgTools(i)
        End If
    Next i
    If btnUndo.Enabled Then btnUndo = imgUndoPics(0)
    If btnRedo.Enabled Then btnRedo = imgRedo(0)
    If CombineMode = RGN_OR Then
        btnMode(1) = imgNegative(0)
    Else
        btnMode(0) = imgAdditive(0)
    End If
    btnFlood = imgFlood(0)
    btnPicture = imgOpen(0)
    If Not bMoveSelection Then btnMove = imgMove(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject antsBrush(0)
    DeleteObject antsBrush(1)
End Sub

Private Sub tmrAnts_Timer()
    Dim res As Long
    
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3
            If shpSelection.BorderStyle = 4 Then
                shpSelection.BorderStyle = 3
            Else
                shpSelection.BorderStyle = 4
            End If
        Case 4
            Dim i As Integer
            
            For i = 0 To lnPoly.UBound
                lnPoly(i).BorderStyle = IIf(lnPoly(i).BorderStyle = 4, 5, 4)
            Next i
        Case 5
            Static idx As Integer

            If bSelectionChanged Then
                bSelectionChanged = False
                canvas.Cls
                canvas.Refresh
                outlineType = 0
                idx = 0
                ' redraw the previous selection outline
                res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
                canvas.Refresh
            End If
            res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
            outlineType = IIf(outlineType = 0, 1, 0)
            res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
            canvas.Refresh
    End Select
End Sub

Private Sub tmrSelection_Timer()
    Dim res As Integer
    Static idx As Integer
    
    If bSelectionChanged Then
        bSelectionChanged = False
        canvas.Cls
        canvas.Refresh
        outlineType = 0
        idx = 0
        ' redraw the previous selection outline
        res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
        canvas.Refresh
    End If
    res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
    outlineType = IIf(outlineType = 0, 1, 0)
    res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
    canvas.Refresh
End Sub
