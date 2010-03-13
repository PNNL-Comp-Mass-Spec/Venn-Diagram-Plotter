Option Strict On

Imports System.Drawing

' -------------------------------------------------------------------------------
' Written by Kyle Littlefield for the Department of Energy (PNNL, Richland, WA)
' Software maintained by Matthew Monroe (PNNL, Richland, WA)
' Program started in August 2004
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'

Public Class TwoCircleVennDiagram
    Inherits VennDiagramBaseClass

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'TwoCircleVennDiagram
        '
        Me.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "TwoCircleVennDiagram"

    End Sub

#End Region

#Region "Properties "
    Public Property OverlapColor() As Color
        Get
            Return Me.OverlapABColor
        End Get
        Set(ByVal Value As Color)
            Me.OverlapABColor = Value
        End Set
    End Property

    Public Property OverlapSize() As Double
        Get
            Return Me.OverlapABSize
        End Get
        Set(ByVal Value As Double)
            Me.OverlapABSize = Value
        End Set
    End Property

    Public Property OverlapLabel() As String()
        Get
            Return Me.OverlapABLabel
        End Get
        Set(ByVal Value() As String)
            Me.OverlapABLabel = Value
        End Set
    End Property
#End Region

    Protected Function AreaOfArc(ByVal dblRadius As Double, ByVal dblSweepAngle As Double) As Double
        Return (1 / 2 * dblRadius ^ 2) * (dblSweepAngle - Math.Sin(dblSweepAngle))
    End Function

    'The second part in computation.  Transforms the world coordinates that have been coordinated
    'into screen coordinates that are then used to draw GDI+ shapes.
    Protected Overrides Sub ComputeScreenCoordinatesFromWorldCoordinates()
        If (Me.Height <= 0 OrElse Me.Width <= 0) Then
            'don't attempt to draw when size has been set to 0 or less in one direction
            'this occurs when the form the control is minimized, so throwing an exception
            'is a bit unreasonable
            Exit Sub
        End If

        If Not (Me.m_ComputeWorldCoordinatesValid) Then
            Me.ComputeWorldCoordinates()
        End If

        If Not Me.SizesReasonable Then
            Exit Sub
        End If

        'The width of the compute world coordinates, from the outside of circleA to the outside of circleB
        Dim computeWorldWidth As Double = Me.m_CircleB_Loc.X + Me.m_CircleB_Radius_Compute + Me.m_CircleA_Radius_Compute
        Dim computeWorldHeight As Double = Math.Max(2 * Me.m_CircleB_Radius_Compute, 2 * Me.m_CircleA_Radius_Compute)

        'Factor that world coordinates need to be scaled to make height fit in control
        Dim heightScaleFactor As Double = (Me.Height) * Me.FillFactor / computeWorldHeight
        'Factor that world coordinates need to be scaled to make width fit in control
        Dim widthScaleFactor As Double = Me.Width * Me.FillFactor / computeWorldWidth

        'Factor that world coordinates need to be scaled to make both width and height fit
        Dim scaleFactor As Double = Math.Min(heightScaleFactor, widthScaleFactor)

        Dim circleAWidth As Double = (2 * Me.m_CircleA_Radius_Compute) * scaleFactor
        Dim circleBWidth As Double = (2 * Me.m_CircleB_Radius_Compute) * scaleFactor
        Dim screenWidth As Double = computeWorldWidth * scaleFactor

        'The amount that the left edge of circleA has to be indented from the edge of the control.
        Dim xInset As Double = (Me.Width - screenWidth) / 2

        'Have to compute left and top coordinates for drawing circles
        Dim circleALeft As Double = xInset
        Dim circleBLeft As Double = xInset + (Me.m_CircleA_Radius_Compute + Me.m_CircleB_Loc.X - Me.m_CircleB_Radius_Compute) * scaleFactor
        Dim circleATop As Double = (Me.Height) / 2 - (Me.m_CircleA_Radius_Compute * scaleFactor)
        Dim circleBTop As Double = (Me.Height) / 2 - (Me.m_CircleB_Radius_Compute * scaleFactor)


        'Rectangles around circleA and circleB
        Dim circleABoundingBox As RectangleF = New RectangleF(CSng(circleALeft), CSng(circleATop), CSng(circleAWidth), CSng(circleAWidth))
        Dim circleBBoundingBox As RectangleF = New RectangleF(CSng(circleBLeft), CSng(circleBTop), CSng(circleBWidth), CSng(circleBWidth))

        'Console.WriteLine("Computing Screen Coordinates")

        'screen coordinates don't exist as class level variables, they are inherently part
        'of the DrawingPath property of VennDiagramAreaInfo
        Me.m_circleA.DrawingPath.Reset()
        Me.m_circleA.DrawingPath.AddEllipse(circleABoundingBox)

        Me.m_circleB.DrawingPath.Reset()
        Me.m_circleB.DrawingPath.AddEllipse(circleBBoundingBox)

        If Me.PaintSolidColorCircles Then
            'create overlap region by taking an arc from each circle
            DrawOverlapRegion(Me.m_overlapAB.DrawingPath, circleABoundingBox, circleBBoundingBox, Me.m_OverlapAB_alpha, Me.m_OverlapAB_beta, 0, 0)
        End If

        'Console.WriteLine("calculated screen coordinates")
        Me.m_ScreenCoordinatesValid = True

    End Sub

    'The first step in drawing a venn diagram, compute world coordinates with the
    'following properties.
    '   Position of circleA is (0, 0)
    '   Position of circleB is (x, 0)
    '   Areas of all shapes are equal to the sizes set by the user.
    Public Overrides Sub ComputeWorldCoordinates()
        'calculate positions/shapes of circles and overlap
        'implements the algorithm described in http://www.cs.uvic.ca/~ruskey/Publications/VennArea/VennArea.pdf

        'current potential range for x coordinate of circleB center
        Dim dblMinX As Double
        Dim dblMaxX As Double

        'current x coordinate of circleB center
        Dim currentCenter As Double

        'Number of tries to get a close enough center.  If it exceeds 1000, then give up 
        'and throw an exception
        Dim tries As Integer = 0

        Me.m_ComputeWorldCoordinatesValid = False

        'check for conditions that can not be painted
        If (Me.m_circleA.Size <= 0 And Me.m_circleB.Size <= 0) Then
            Me.SizesReasonable = False
            Exit Sub
        End If

        If (Me.m_overlapAB.Size > Math.Min(Me.m_circleA.Size, Me.m_circleB.Size)) Then
            Throw New ApplicationException("Overlap greater than one (or both) circles")
        End If

        'When size for one circle is equal to 0 this algorithm fails to work.  Solution
        'is to just make the size really small.  This causes it to be unnoticeable 
        'on the printed control.
        If (Me.CircleASize <= 0 And Me.CircleBSize > 0) Then
            Me.CircleASize = 0.0000000001
        End If

        If (Me.CircleBSize <= 0 And Me.CircleASize > 0) Then
            Me.CircleBSize = 0.0000000001
        End If

        'Sizes are all reasonable, so move ahead with the calculation
        Me.SizesReasonable = True

        Me.m_CircleA_Radius_Compute = Math.Sqrt(Me.m_circleA.Size / Math.PI)
        Me.m_CircleB_Radius_Compute = Math.Sqrt(Me.m_circleB.Size / Math.PI)

        'Minimum value for x coordinate, edges touching (with A completely contained in B or vice versa)
        dblMinX = Math.Abs(Me.m_CircleA_Radius_Compute - Me.m_CircleB_Radius_Compute)

        'Maximum value for x coordinate, no overlap
        dblMaxX = Me.m_CircleA_Radius_Compute + Me.m_CircleB_Radius_Compute

        Do
            Dim overlapDifference As Double
            Dim alpha As Double
            ''Dim alphanum As Double
            ''Dim alphadenom As Double
            Dim beta As Double
            ''Dim betanum As Double
            ''Dim betadenom As Double

            'Calculated area of overlap
            Dim overlap As Double

            currentCenter = (dblMaxX + dblMinX) / 2

            ' Compute the angles (in radians) that define the arcs that bound the overlap area
            alpha = 2 * AngleAFromThreeSides(Me.m_CircleB_Radius_Compute, currentCenter, Me.m_CircleA_Radius_Compute)
            beta = 2 * AngleAFromThreeSides(Me.m_CircleA_Radius_Compute, currentCenter, Me.m_CircleB_Radius_Compute)

            'compute areas of arcs and add together to get overlap area
            overlap = AreaOfArc(Me.m_CircleA_Radius_Compute, alpha) + AreaOfArc(Me.m_CircleB_Radius_Compute, beta)

            overlapDifference = overlap - Me.OverlapABSize

            'If overlapDifference is less than 10^-12 of the larger of the two circles, then
            'call computation done.
            If overlapDifference > -Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 AndAlso overlapDifference < Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 Then
                Me.m_CircleB_Loc.X = currentCenter
                Me.m_CircleB_Loc.Y = 0
                Me.m_OverlapAB_alpha = ConvertRadiansToDegrees(alpha)
                Me.m_OverlapAB_beta = ConvertRadiansToDegrees(beta)
                Exit Do
            End If

            'readjust maximum and minimum if not close enough
            If overlapDifference > 0 Then
                dblMinX = currentCenter
            ElseIf overlapDifference < 0 Then
                dblMaxX = currentCenter
            End If

            tries = tries + 1
            If tries > VennDiagramBaseClass.MAX_COMPUTE_TRIES Then
                'This should only happen in the case that the user 
                'specifies an overlap that is bigger than one or both of the base sets.
                'can also happen when size of component is very small (like 1 * 1 pixels)
                Me.m_ComputeWorldCoordinatesValid = False
                Throw New ApplicationException("Unable to compute locations.")
            End If
        Loop

        'Console.WriteLine("calculated world coordinates")
        Me.m_ComputeWorldCoordinatesValid = True

        'test
        'Dim angle As Single = 120

        'Me.m_circleA.DrawingPath = New Drawing2D.GraphicsPath
        'Me.m_circleA.DrawingPath.AddEllipse(0, 0, 100, 100)

        'Me.m_circleB.DrawingPath = New Drawing2D.GraphicsPath
        'Me.m_circleB.DrawingPath.AddEllipse(50, 0, 100, 100)

        'Me.m_overlapAB.DrawingPath = New Drawing2D.GraphicsPath
        'Me.m_overlapAB.DrawingPath.AddArc(New Rectangle(0, 0, 100, 100), -angle / 2, angle)
        'Me.m_overlapAB.DrawingPath.AddArc(New Rectangle(50, 0, 100, 100), -angle / 2 + 180, angle)
        'end test

        'Me.m_readyToDraw = True
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        'Me.m_overlapAB.Color = Color.WhiteSmoke
        'handles basic control properties, background
        'If Not (Me.m_ComputeWorldCoordinatesValid) Then
        '    Me.ComputeComputeCoordinates()
        'End If

        'If Not (Me.m_ScreenCoordinatesValid) Then
        '    Me.ComputeScreenCoordinatesFromWorldCoordinates()
        'End If

        Dim g As Graphics = e.Graphics

        'Begin a container so smoothing mode change does not impact others drawing on same graphics
        Dim container As Drawing2D.GraphicsContainer = g.BeginContainer()

        g.SmoothingMode = Me.SmoothingMode

        'Recompute screen coordinates if necessary.  Will cause compute world coordinates to be
        'recomputed if necessary
        If Not Me.m_ScreenCoordinatesValid Then
            Me.ComputeScreenCoordinatesFromWorldCoordinates()
        End If

        'second time, so if there sizes are not reasonable, then
        'return without painting anything.  For example, this occurs when the overlap area 
        'exceeds the area of one or both circles
        If Not Me.SizesReasonable Then
            Exit Sub
        End If

        'draw circle and overlap, starting with circleA and circleB, then with overlap so that it draws over them
        Me.PaintVennDiagramInfo(Me.m_circleA, g)
        Me.PaintVennDiagramInfo(Me.m_circleB, g)
        Me.PaintVennDiagramInfo(Me.m_overlapAB, g)
        'deal with labels - TODO
        'g.DrawString(Me.m_circleA.Label, Me.Font, New SolidBrush(Me.ForeColor), 0, 50)
        MyBase.OnPaint(e)
    End Sub

    ''Private Sub painted(ByVal sender As Object, ByVal args As PaintEventArgs)
    ''    'Console.WriteLine("Painting")
    ''End Sub

End Class

