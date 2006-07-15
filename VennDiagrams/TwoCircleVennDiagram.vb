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
    Inherits System.Windows.Forms.UserControl
    Implements ControlPrinter.IPrintableControl

    Protected Enum TwoCircleVennDiagramPart
        CIRCLEA
        CIRCLEB
        OVERLAP
        NONE
    End Enum

    Protected m_overlap As VennDiagramAreaInfo = New VennDiagramAreaInfo
    Protected m_circleA As VennDiagramAreaInfo = New VennDiagramAreaInfo
    Protected m_circleB As VennDiagramAreaInfo = New VennDiagramAreaInfo
    Protected m_outlinePen As Pen = Pens.Black
    Protected m_ScreenCoordinatesValid As Boolean = False
    Protected m_ComputeWorldCoordinatesValid As Boolean = False
    Protected m_smoothingMode As Drawing2D.SmoothingMode = Drawing2D.SmoothingMode.Default
    Protected m_sizesReasonable As Boolean = True

    Protected Const MAX_COMPUTE_TRIES As Integer = 1000

    'Screen Coordinates
    'Protected m_A_Center_Screen As Point
    'Protected m_B_Center_Screen As Point

    'Computation World Coordinates
    Protected Const m_A_Center_Compute As Double = 0.0
    Protected m_A_Radius_Compute As Double
    Protected m_B_Center_Compute As Double
    Protected m_B_Radius_Compute As Double
    Protected m_alpha As Double
    Protected m_beta As Double

    'the amount the inner part of the venn diagram attempts to fill - by dimension, not area
    Protected m_fillSize As Double = 0.9
    'location, size, etc. is held in the DrawingPath property of the VennDiagramAreaInfo

    Sub painted(ByVal sender As Object, ByVal args As PaintEventArgs)
        'Console.WriteLine("Painting")
    End Sub
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'm_tip.AutomaticDelay = 1000
        'm_tip.AutoPopDelay = 1000
        'm_tip.InitialDelay = 2500
        'm_tip.ShowAlways = False
        Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        AddHandler Me.Paint, AddressOf Me.painted
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
        Me.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "TwoCircleVennDiagram"

    End Sub

#End Region

    'Forces screen coordinates to be recomputed from compute world coordinates
    'for example, when the control is resized
    Protected Sub InvalidateScreenCoordinates()
        Me.m_ScreenCoordinatesValid = False
        Me.Invalidate()
    End Sub

    'Forces world coordinates to be recomputed
    'for example, when one of the region sizes is changed
    Protected Sub InvalidateComputeWorldCoordinates()
        Me.m_ScreenCoordinatesValid = False
        Me.m_ComputeWorldCoordinatesValid = False
        Me.Invalidate()
    End Sub

#Region "Properties "
    Public Property CircleASize() As Double
        Get
            Return Me.m_circleA.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_circleA.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    Public Property CircleBSize() As Double
        Get
            Return Me.m_circleB.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_circleB.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    Public Property OverlapSize() As Double
        Get
            Return Me.m_overlap.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_overlap.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    'Labels are not currently used.  They were intended to provide
    'the ability to label the diagram with intelligently placed labels
    'on the actual diagram
    Public Property CircleALabel() As String()
        Get
            Return Me.m_circleA.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_circleA.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property CircleBLabel() As String()
        Get
            Return Me.m_circleB.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_circleB.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property OverlapLabel() As String()
        Get
            Return Me.m_overlap.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_overlap.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property CircleAColor() As System.Drawing.Color
        Get
            Return Me.m_circleA.Color
        End Get
        Set(ByVal Value As System.Drawing.Color)
            Me.m_circleA.Color = Value
            'Me.UpdateOverlapColor()
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    Public Property CircleBColor() As Color
        Get
            Return Me.m_circleB.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_circleB.Color = Value
            'Me.UpdateOverlapColor()
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    Public Property OverlapColor() As Color
        Get
            Return Me.m_overlap.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_overlap.Color = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Protected Sub UpdateOverlapColor()
    '    Me.m_overlap.Color = Me.MergeTwoColors(Me.CircleAColor, Me.CircleBColor)
    'End Sub

    'The pen that is used for outlining around the circles and overlap region
    Public Property OutlinePen() As Pen
        Get
            Return Me.m_outlinePen
        End Get
        Set(ByVal Value As Pen)
            Me.m_outlinePen = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Controls how close to the bounds of the control the colored part of the diagram will go.
    'Diagram fills up this fraction of either the width or height of the control
    Public Property FillSize() As Double
        Get
            Return Me.m_fillSize
        End Get
        Set(ByVal Value As Double)
            If (Value < 0 Or Value > 1.0) Then
                Throw New ArgumentException("Must be in range of 0 to 1")
            End If
            Me.m_fillSize = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'When drawn, the control will change the smoothing mode of the graphics to this
    Public Property SmoothingMode() As Drawing2D.SmoothingMode
        Get
            Return Me.m_smoothingMode
        End Get
        Set(ByVal Value As Drawing2D.SmoothingMode)
            Me.m_smoothingMode = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Tells whether the sizes of the circles and overlap are consistent
    Protected Property SizesReasonable() As Boolean
        Get
            Return Me.m_sizesReasonable
        End Get
        Set(ByVal Value As Boolean)
            Me.m_sizesReasonable = Value
        End Set
    End Property
#End Region

    'Raised when smoothingmode, fillsize, or colors are changed
    Public Event DrawingChange(ByVal sender As TwoCircleVennDiagram)

    'Never raised
    Public Event LabelChange(ByVal sender As TwoCircleVennDiagram)

    'Raised when size of circles or overlap is changed.
    Public Event SizeChange(ByVal sender As TwoCircleVennDiagram)

    'Protected Function MergeTwoColors(ByVal colorA As Color, ByVal colorB As Color) As Color
    '    Dim alpha As Byte
    '    Dim red As Byte
    '    Dim green As Byte
    '    Dim blue As Byte

    '    alpha = CByte((CInt(colorA.A) + CInt(colorB.A)) / 2)
    '    red = CByte((CInt(colorA.R) + CInt(colorB.R)) / 2)
    '    green = CByte((CInt(colorA.G) + CInt(colorB.G)) / 2)
    '    blue = CByte((CInt(colorA.B) + CInt(colorB.B)) / 2)

    '    Return Color.FromArgb(alpha, red, green, blue)
    'End Function

    Protected Function GetPosition(ByVal x As Integer, ByVal y As Integer) As TwoCircleVennDiagramPart
        If (Me.m_overlap.DrawingPath.IsVisible(x, y)) Then
            Return TwoCircleVennDiagramPart.OVERLAP
        ElseIf (Me.m_circleA.DrawingPath.IsVisible(x, y)) Then
            Return TwoCircleVennDiagramPart.CIRCLEA
        ElseIf (Me.m_circleB.DrawingPath.IsVisible(x, y)) Then
            Return TwoCircleVennDiagramPart.CIRCLEB
        Else
            Return TwoCircleVennDiagramPart.NONE
        End If
    End Function

    Dim m_tip As ToolTip = New ToolTip

    Protected Overrides Sub OnMouseMove(ByVal args As MouseEventArgs)
        'Console.WriteLine(System.Enum.GetName(GetType(PointRegion), Me.GetPosition(args.X, args.Y)))
        'm_tip.SetToolTip(Me, System.Enum.GetName(GetType(TwoCircleVennDiagramPart), Me.GetPosition(args.X, args.Y)))
    End Sub

    'The second part in computation.  Transfroms the world coordinates that have been coordinated
    'into screen coordinates that are then used to draw GDI+ shapes.
    Protected Sub ComputeScreenCoordinatesFromWorldCoordinates()
        If (Me.Height <= 0 OrElse Me.Width <= 0) Then
            'don't attempt to draw when size has been set to 0 or less in one direction
            'this occurs when the form the control is in is minimized, so throwing an exception
            'is a bit unreasonable
            Return
        End If

        If Not (Me.m_ComputeWorldCoordinatesValid) Then
            Me.ComputeWorldCoordinates()
        End If

        If Not Me.SizesReasonable Then
            Return
        End If

        'The width of the compute world coordinates, from the outside of circleA to the outside of circleB
        Dim computeWorldWidth As Double = Me.m_B_Center_Compute + Me.m_B_Radius_Compute + Me.m_A_Radius_Compute
        Dim computeWorldHeight As Double = Math.Max(2 * Me.m_B_Radius_Compute, 2 * Me.m_A_Radius_Compute)

        'Factor that world coordinates need to be scaled to make height fit in control
        Dim heightScaleFactor As Double = (Me.Height) * Me.FillSize / computeWorldHeight
        'Factor that world coordinates need to be scaled to make width fit in control
        Dim widthScaleFactor As Double = Me.Width * Me.FillSize / computeWorldWidth

        'Factor that world coordinates need to be scaled to make both width and height fit
        Dim scaleFactor As Double = Math.Min(heightScaleFactor, widthScaleFactor)

        Dim circleAWidth As Double = (2 * Me.m_A_Radius_Compute) * scaleFactor
        Dim circleBWidth As Double = (2 * Me.m_B_Radius_Compute) * scaleFactor
        Dim screenWidth As Double = computeWorldWidth * scaleFactor

        'The amount that the left edge of circleA has to be indented from the edge of the control.
        Dim xInset As Double = (Me.Width - screenWidth) / 2

        'Have to compute left and top coordinates for drawing circles
        Dim circleALeft As Double = xInset
        Dim circleBLeft As Double = xInset + (Me.m_A_Radius_Compute + Me.m_B_Center_Compute - Me.m_B_Radius_Compute) * scaleFactor
        Dim circleATop As Double = (Me.Height) / 2 - (Me.m_A_Radius_Compute * scaleFactor)
        Dim circleBTop As Double = (Me.Height) / 2 - (Me.m_B_Radius_Compute * scaleFactor)


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

        'create overlap region by taking an arc from each circle
        Me.m_overlap.DrawingPath.Reset()
        Me.m_overlap.DrawingPath.AddArc(circleABoundingBox, -ConvertAngleToDegrees(Me.m_alpha / 2), ConvertAngleToDegrees(Me.m_alpha))
        Me.m_overlap.DrawingPath.AddArc(circleBBoundingBox, 180 - ConvertAngleToDegrees(Me.m_beta / 2), ConvertAngleToDegrees(Me.m_beta))

        'Console.WriteLine("calculated screen coordinates")
        Me.m_ScreenCoordinatesValid = True
    End Sub

    Protected Function ConvertAngleToDegrees(ByVal radians As Double) As Single
        Return CSng(radians * 180 / Math.PI)
    End Function

    'The first step in drawing a venn diagram, compute world coordinates with the
    'following properties.
    '   Position of circleA is (0, 0)
    '   Position of circleB is (x, 0)
    '   Areas of all shapes are equal to the sizes set by the user.
    Public Sub ComputeWorldCoordinates()
        'calculate positions/shapes of circles and overlap
        'implements the algorithm described in http://www.cs.uvic.ca/~ruskey/Publications/VennArea/VennArea.pdf

        'current potential range for x coordinate of circleB center
        Dim high As Double
        Dim low As Double
        'current x coordinate of circleB center
        Dim currentCenter As Double

        'Number of tries to get a close enough center.  If it exceeds 1000, then give up 
        'and through a exception
        Dim tries As Integer = 0

        Me.m_ComputeWorldCoordinatesValid = False

        'check for conditions that can not be painted
        If (Me.m_circleA.Size <= 0 And Me.m_circleB.Size <= 0) Then
            Me.SizesReasonable = False
            Return
        End If

        If (Me.m_overlap.Size > Math.Min(Me.m_circleA.Size, Me.m_circleB.Size)) Then
            Throw New ApplicationException("Overlap greater than one (or both) circles")
        End If

        'When size for one circle is equal to 0 this algorithm fails to work.  Solution
        'is to just make the size really small.  This causes it to be unnoticeable 
        'on the printed control.
        If (Me.CircleASize <= 0 And Me.CircleBSize > 0) Then
            Me.CircleASize = 0.0000000001
        End If

        If (Me.CircleBSize <= 0 And Me.CircleASize > 0) Then
            Me.CircleBSize = 0.00000000001
        End If

        'Sizes are all reasonable, so move ahead with the calculation
        Me.SizesReasonable = True

        Me.m_A_Radius_Compute = Math.Sqrt(Me.m_circleA.Size / Math.PI)
        Me.m_B_Radius_Compute = Math.Sqrt(Me.m_circleB.Size / Math.PI)

        'Maximum value for x coordinate, no overlap
        high = Me.m_A_Radius_Compute + Me.m_B_Radius_Compute
        'Minimum value for x coordinate, edges touching (with A completely contained in B or vice versa)
        low = Math.Abs(Me.m_A_Radius_Compute - Me.m_B_Radius_Compute)

        Do
            Dim overlapDifference As Double
            Dim alpha As Double
            Dim alphanum As Double
            Dim alphadenom As Double
            Dim beta As Double
            Dim betanum As Double
            Dim betadenom As Double
            Dim d As Double
            Dim r1 As Double
            Dim r2 As Double

            'Calculated area of overlap
            Dim overlap As Double

            r1 = Me.m_A_Radius_Compute
            r2 = Me.m_B_Radius_Compute
            currentCenter = (high + low) / 2
            d = currentCenter

            alphanum = (d * d) + (r1 * r1) - (r2 * r2)
            alphadenom = 2 * r1 * d
            alpha = 2 * Math.Acos(alphanum / alphadenom)

            betanum = (d * d) + (r2 * r2) - (r1 * r1)
            betadenom = 2 * r2 * d
            beta = 2 * Math.Acos(betanum / betadenom)

            'compute areas of arcs and add together to get overlap area
            overlap = (1 / 2 * r1 * r1) * (alpha - Math.Sin(alpha))
            overlap = overlap + (1 / 2 * r2 * r2) * (beta - Math.Sin(beta))

            overlapDifference = overlap - Me.OverlapSize

            'If overlapDifference is less than 10^-12 of the larger of the tow circles, then
            'call computation done.
            If overlapDifference > -Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 AndAlso overlapDifference < Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 Then
                Me.m_B_Center_Compute = currentCenter
                Me.m_alpha = alpha
                Me.m_beta = beta 'could recompute this outside of loop
                Exit Do
            End If

            'readjust maximum and minimum if not close enough
            If overlapDifference > 0 Then
                low = currentCenter
            ElseIf overlapDifference < 0 Then
                high = currentCenter
            End If

            tries = tries + 1
            If tries > Me.MAX_COMPUTE_TRIES Then
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

        'Me.m_overlap.DrawingPath = New Drawing2D.GraphicsPath
        'Me.m_overlap.DrawingPath.AddArc(New Rectangle(0, 0, 100, 100), -angle / 2, angle)
        'Me.m_overlap.DrawingPath.AddArc(New Rectangle(50, 0, 100, 100), -angle / 2 + 180, angle)
        'end test

        'Me.m_readyToDraw = True
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        'Me.m_overlap.Color = Color.WhiteSmoke
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
            Return
        End If



        'draw circle and overlap, starting with circleA and circleB, then with overlap so that it draws over them
        Me.PaintVennDiagramInfo(Me.m_circleA, g)
        Me.PaintVennDiagramInfo(Me.m_circleB, g)
        Me.PaintVennDiagramInfo(Me.m_overlap, g)
        'deal with labels - TODO
        'g.DrawString(Me.m_circleA.Label, Me.Font, New SolidBrush(Me.ForeColor), 0, 50)
        MyBase.OnPaint(e)
    End Sub

    Protected Sub PaintVennDiagramInfo(ByVal info As VennDiagramAreaInfo, ByVal graphics As Graphics)
        Dim b As Brush = New SolidBrush(info.Color)

        graphics.FillPath(b, info.DrawingPath)
        graphics.DrawPath(Me.OutlinePen, info.DrawingPath)
    End Sub

    Protected Overrides Sub OnResize(ByVal args As EventArgs)
        MyBase.OnResize(args)
        'Console.WriteLine("Resize")
        If Me.m_ComputeWorldCoordinatesValid Then
            Me.m_ScreenCoordinatesValid = False
        End If
        Me.Invalidate()
    End Sub

    Public Sub DrawOnGraphics(ByVal g As Graphics, ByVal drawBackground As Boolean) Implements ControlPrinter.IPrintableControl.DrawOnGraphics
        If (drawBackground) Then
            g.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(0, 0, Me.Width, Me.Height))
        End If
        Me.OnPaint(New PaintEventArgs(g, New Rectangle(0, 0, 0, 0)))
    End Sub


    'Public ReadOnly Property ReadyToDraw() As Boolean
    '    Get
    '        Return False
    '    End Get
    'End Property
End Class

Public Class VennDiagramAreaInfo
    Private m_size As Double = 1.0
    Private m_label As String() = New String() {}
    Private m_path As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
    Private m_color As System.Drawing.Color = Color.Blue

    'Would be more powerful if it was a Brush instead of a merely a color, allowing, for example
    'hatching or gradients, but then the TwoCircleVennDiagramPresenter would not be able to 
    'synchronize label colors with what venn diagram was drawing
    Friend Property Color() As Color
        Get
            Return Me.m_color
        End Get
        Set(ByVal Value As Color)
            Me.m_color = Value
        End Set
    End Property

    'Labels are not currently used
    Friend Property Label() As String()
        Get
            Return Me.m_label
        End Get
        Set(ByVal Value() As String)
            Me.m_label = Value
        End Set
    End Property

    Friend Property Size() As Double
        Get
            Return Me.m_size
        End Get
        Set(ByVal Value As Double)
            Me.m_size = Value
        End Set
    End Property

    Friend Property DrawingPath() As System.Drawing.Drawing2D.GraphicsPath
        Get
            Return Me.m_path
        End Get
        Set(ByVal Value As System.Drawing.Drawing2D.GraphicsPath)
            Me.m_path = Value
        End Set
    End Property
End Class

