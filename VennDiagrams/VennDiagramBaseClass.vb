Option Strict On

Imports System.Drawing

' -------------------------------------------------------------------------------
' Written by Matthew Monroe the Department of Energy (PNNL, Richland, WA)
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'

Public MustInherit Class VennDiagramBaseClass
    Inherits System.Windows.Forms.UserControl
    Implements ControlPrinter.IPrintableControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InitializeVariables()
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
        'VennDiagramBaseClass
        '
        Me.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "VennDiagramBaseClass"

    End Sub

#End Region

#Region "Events"

    'Raised when smoothingmode, FillFactor, or colors are changed
    Public Event DrawingChange(ByVal sender As VennDiagramBaseClass)

    'Never raised
    Public Event LabelChange(ByVal sender As VennDiagramBaseClass)

    'Raised when size of circles or overlap is changed.
    Public Event SizeChange(ByVal sender As VennDiagramBaseClass)

#End Region

#Region "Structures"
    Public Structure udtPointXYType
        Public X As Double
        Public Y As Double
    End Structure

    Public Structure udtBoundingRegionType
        Public UpperLeft As udtPointXYType
        Public LowerRight As udtPointXYType
    End Structure
#End Region

#Region "Constants"
    Protected Const MAX_COMPUTE_TRIES As Integer = 1000

#End Region

#Region "Member Variables"
    ' Note: Location, size, etc. for each circle is held in the DrawingPath property of the VennDiagramAreaInfo

    Protected m_circleA As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorCircleA)
    Protected m_circleB As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorCircleB)
    Protected m_overlapAB As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorOverlapAB)

    Protected m_outlinePen As Pen = Pens.Black

    Protected m_ScreenCoordinatesValid As Boolean = False
    Protected m_ComputeWorldCoordinatesValid As Boolean = False
    Protected m_smoothingMode As Drawing2D.SmoothingMode = Drawing2D.SmoothingMode.Default
    Protected m_PaintSolidColorCircles As Boolean = True

    Protected m_sizesReasonable As Boolean = True


    'Computation World Coordinates
    ''    Protected Const m_CircleA_Center_Compute As Double = 0.0
    Protected m_CircleA_Loc As udtPointXYType
    Protected m_CircleA_Radius_Compute As Double

    ''Protected m_CircleB_Center_Compute As Double            ' Distance from the center of Circle B to the center of circle A
    Protected m_CircleB_Loc As udtPointXYType
    Protected m_CircleB_Radius_Compute As Double

    Protected m_OverlapAB_alpha As Double                     ' Angle of the arc to draw that defines the overlap region (in degrees)
    Protected m_OverlapAB_beta As Double                      ' Angle of the arc to draw that defines the overlap region (in degrees)

    'the amount the inner part of the venn diagram attempts to fill - by dimension, not area
    Protected m_FillFactor As Double = 0.95        ' Larger values increase the size of the circles by decrease the white space on the edges (0 to 2)

#End Region

#Region "Properties "
    Public Property CircleASize() As Double
        Get
            Return m_circleA.Size
        End Get
        Set(ByVal Value As Double)
            m_circleA.Size = Value
            InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    Public Property CircleBSize() As Double
        Get
            Return m_circleB.Size
        End Get
        Set(ByVal Value As Double)
            m_circleB.Size = Value
            InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    Public ReadOnly Property CoordCircleA() As udtPointXYType
        Get
            Return m_CircleA_Loc
        End Get
    End Property
    Public ReadOnly Property CoordCircleARadius() As Double
        Get
            Return m_CircleA_Radius_Compute
        End Get
    End Property

    Public ReadOnly Property CoordCircleB() As udtPointXYType
        Get
            Return m_CircleB_Loc
        End Get
    End Property
    Public ReadOnly Property CoordCircleBRadius() As Double
        Get
            Return m_CircleB_Radius_Compute
        End Get
    End Property

    Public ReadOnly Property DefaultColorOverlapAB() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(192, 255, 192)
        End Get
    End Property

    Public ReadOnly Property DefaultColorCircleA() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(255, 192, 192)
        End Get
    End Property

    Public ReadOnly Property DefaultColorCircleB() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(192, 255, 255)
        End Get
    End Property

    Public ReadOnly Property OverlapABAlpha() As Double
        Get
            Return m_OverlapAB_alpha
        End Get
    End Property

    Public ReadOnly Property OverlapABBeta() As Double
        Get
            Return m_OverlapAB_beta
        End Get
    End Property

    Public Property OverlapABSize() As Double
        Get
            Return m_overlapAB.Size
        End Get
        Set(ByVal Value As Double)
            m_overlapAB.Size = Value
            InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChange(Me)
        End Set
    End Property

    'Labels are not currently used.  They were intended to provide
    'the ability to label the diagram with intelligently placed labels
    'on the actual diagram
    Public Property CircleALabel() As String()
        Get
            Return m_circleA.Label
        End Get
        Set(ByVal Value() As String)
            m_circleA.Label = Value
            InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property CircleBLabel() As String()
        Get
            Return m_circleB.Label
        End Get
        Set(ByVal Value() As String)
            m_circleB.Label = Value
            InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property OverlapABLabel() As String()
        Get
            Return m_overlapAB.Label
        End Get
        Set(ByVal Value() As String)
            m_overlapAB.Label = Value
            InvalidateScreenCoordinates()
            RaiseEvent LabelChange(Me)
        End Set
    End Property

    Public Property CircleAColor() As System.Drawing.Color
        Get
            Return m_circleA.Color
        End Get
        Set(ByVal Value As System.Drawing.Color)
            m_circleA.Color = Value
            'Me.UpdateOverlapColor()
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    Public Property CircleBColor() As Color
        Get
            Return m_circleB.Color
        End Get
        Set(ByVal Value As Color)
            m_circleB.Color = Value
            'Me.UpdateOverlapColor()
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    Public Property OverlapABColor() As Color
        Get
            Return m_overlapAB.Color
        End Get
        Set(ByVal Value As Color)
            m_overlapAB.Color = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Controls how close to the bounds of the control the colored part of the diagram will go.
    'Diagram fills up this fraction of either the width or height of the control
    Public Property FillFactor() As Double
        Get
            Return m_FillFactor
        End Get
        Set(ByVal Value As Double)
            If (Value < 0 Or Value > 2.0) Then
                Throw New ArgumentException("Must be in range of 0 to 2")
            End If
            m_FillFactor = Value
            InvalidateScreenCoordinates()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Protected Sub UpdateOverlapColor()
    '    Me.m_overlapAB.Color = Me.MergeTwoColors(Me.CircleAColor, Me.CircleBColor)
    'End Sub

    'The pen that is used for outlining around the circles and overlap region
    Public Property OutlinePen() As Pen
        Get
            Return m_outlinePen
        End Get
        Set(ByVal Value As Pen)
            m_outlinePen = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    Public Property PaintSolidColorCircles() As Boolean
        Get
            Return m_PaintSolidColorCircles
        End Get
        Set(ByVal Value As Boolean)
            If m_PaintSolidColorCircles <> Value Then
                m_PaintSolidColorCircles = Value
                Me.Invalidate()
                RaiseEvent DrawingChange(Me)
            End If
        End Set
    End Property

    'When drawn, the control will change the smoothing mode of the graphics to this
    Public Property SmoothingMode() As Drawing2D.SmoothingMode
        Get
            Return m_smoothingMode
        End Get
        Set(ByVal Value As Drawing2D.SmoothingMode)
            m_smoothingMode = Value
            Me.Invalidate()
            RaiseEvent DrawingChange(Me)
        End Set
    End Property

    'Public ReadOnly Property ReadyToDraw() As Boolean
    '    Get
    '        Return False
    '    End Get
    'End Property

    'Tells whether the sizes of the circles and overlap are consistent
    Protected Property SizesReasonable() As Boolean
        Get
            Return m_sizesReasonable
        End Get
        Set(ByVal Value As Boolean)
            m_sizesReasonable = Value
        End Set
    End Property

#End Region

    Protected Shared Function AngleAFromThreeSides(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
        Dim dblAngleRadians As Double

        ' Given triangle with sides a, b, and c, compute the angle A, which is opposite from side a
        ' Angle is returned in radians
        '
        '       C
        '       /\
        '     b/   \ a
        '     /      \
        '    /         \
        '  A-------------B
        '        c
        '

        dblAngleRadians = AcosSafe((b ^ 2 + c ^ 2 - a ^ 2) / (2 * b * c))

        Return dblAngleRadians

    End Function

    Protected Shared Function AcosSafe(ByVal Xcos As Double) As Double
        'pass in the cosine and receive the angle
        'Xcos of 1, or greater, will cause an error in the Acos
        'function.

        If Xcos > 1 Or Xcos < -1 Then
            ' Invalid input; return an angle of 0
            Return 0
        Else
            ' Could manually compute ACos using:
            ' ACos = Atn(-Xcos / Sqr(-Xcos * Xcos + 1)) + 2 * Atn(1)
            '  or
            ' ACos = Pi / 2 - 2 * Atn(x / (1 + Sqr(Abs(1 - x * x))))
            Return Math.Acos(Xcos)
        End If
    End Function

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

    ''Protected Function GetPosition(ByVal x As Integer, ByVal y As Integer) As TwoCircleVennDiagramPart
    ''    If (Me.m_overlapAB.DrawingPath.IsVisible(x, y)) Then
    ''        Return TwoCircleVennDiagramPart.OVERLAP
    ''    ElseIf (Me.m_circleA.DrawingPath.IsVisible(x, y)) Then
    ''        Return TwoCircleVennDiagramPart.CIRCLEA
    ''    ElseIf (Me.m_circleB.DrawingPath.IsVisible(x, y)) Then
    ''        Return TwoCircleVennDiagramPart.CIRCLEB
    ''    Else
    ''        Return TwoCircleVennDiagramPart.NONE
    ''    End If
    ''End Function

    ''Dim m_tip As ToolTip = New ToolTip
    ''Protected Overrides Sub OnMouseMove(ByVal args As MouseEventArgs)
    ''    'Console.WriteLine(System.Enum.GetName(GetType(PointRegion), Me.GetPosition(args.X, args.Y)))
    ''    'm_tip.SetToolTip(Me, System.Enum.GetName(GetType(TwoCircleVennDiagramPart), Me.GetPosition(args.X, args.Y)))
    ''End Sub


    'The first step in drawing a venn diagram, compute world coordinates with the
    'following properties.
    Public MustOverride Sub ComputeWorldCoordinates()

    'The second part in computation.  Transforms the world coordinates that have been coordinated
    'into screen coordinates that are then used to draw GDI+ shapes.
    Protected MustOverride Sub ComputeScreenCoordinatesFromWorldCoordinates()

    Public Shared Function ConvertDegreesToRadians(ByVal degrees As Double) As Double
        Return degrees * Math.PI / 180
    End Function

    Public Shared Function ConvertRadiansToDegrees(ByVal radians As Double) As Double
        Return radians * 180 / Math.PI
    End Function

    Public Sub DrawOnGraphics(ByVal g As Graphics, ByVal drawBackground As Boolean) Implements ControlPrinter.IPrintableControl.DrawOnGraphics
        If (drawBackground) Then
            g.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(0, 0, Me.Width, Me.Height))
        End If
        Me.OnPaint(New PaintEventArgs(g, New Rectangle(0, 0, 0, 0)))
    End Sub

    Protected Sub DrawOverlapRegion(ByRef objDrawingPath As System.Drawing.Drawing2D.GraphicsPath, ByVal BoundingBoxA As RectangleF, ByVal BoundingBoxB As RectangleF, ByVal dblAlpha As Double, ByVal dblBeta As Double, ByVal dblAlphaStartAddon As Double, ByVal dblBetaStartAddon As Double)
        objDrawingPath.Reset()
        objDrawingPath.AddArc(BoundingBoxA, CSng(-dblAlpha / 2 + dblAlphaStartAddon), CSng(dblAlpha))
        objDrawingPath.AddArc(BoundingBoxB, CSng(180 - dblBeta / 2 + dblBetaStartAddon), CSng(dblBeta))
    End Sub

    Protected Overridable Sub InitializeVariables()
        'm_tip.AutomaticDelay = 1000
        'm_tip.AutoPopDelay = 1000
        'm_tip.InitialDelay = 2500
        'm_tip.ShowAlways = False
        Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    End Sub

    'Forces screen coordinates to be recomputed from compute world coordinates
    'for example, when the control is resized
    Protected Sub InvalidateScreenCoordinates()
        m_ScreenCoordinatesValid = False
        Me.Invalidate()
    End Sub

    'Forces world coordinates to be recomputed
    'for example, when one of the region sizes is changed
    Protected Sub InvalidateComputeWorldCoordinates()
        m_ScreenCoordinatesValid = False
        m_ComputeWorldCoordinatesValid = False
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnResize(ByVal args As EventArgs)
        MyBase.OnResize(args)
        'Console.WriteLine("Resize")
        If m_ComputeWorldCoordinatesValid Then
            m_ScreenCoordinatesValid = False
        End If
        Me.Invalidate()
    End Sub

    Protected Sub PaintVennDiagramInfo(ByVal info As VennDiagramAreaInfo, ByVal graphics As Graphics)
        Dim b As Brush = New SolidBrush(info.Color)

        If Me.PaintSolidColorCircles Then
            graphics.FillPath(b, info.DrawingPath)
        End If

        graphics.DrawPath(Me.OutlinePen, info.DrawingPath)

    End Sub

    'Protected Sub UpdateOverlapColor()
    '    Me.m_overlapAB.Color = Me.MergeTwoColors(Me.CircleAColor, Me.CircleBColor)
    'End Sub

End Class

Public Class VennDiagramAreaInfo
    Private m_size As Double = 1.0
    Private m_label As String() = New String() {}
    Private m_path As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
    Private m_color As System.Drawing.Color = System.Drawing.Color.Blue

    'Would be more powerful if it was a Brush instead of a merely a color, allowing, for example
    'hatching or gradients, but then the TwoCircleVennDiagramPresenter would not be able to 
    'synchronize label colors with what venn diagram was drawing
    Friend Property Color() As Color
        Get
            Return m_color
        End Get
        Set(ByVal Value As Color)
            m_color = Value
        End Set
    End Property

    'Labels are not currently used
    Friend Property Label() As String()
        Get
            Return m_label
        End Get
        Set(ByVal Value() As String)
            m_label = Value
        End Set
    End Property

    Friend Property Size() As Double
        Get
            Return m_size
        End Get
        Set(ByVal Value As Double)
            m_size = Value
        End Set
    End Property

    Friend Property DrawingPath() As System.Drawing.Drawing2D.GraphicsPath
        Get
            Return m_path
        End Get
        Set(ByVal Value As System.Drawing.Drawing2D.GraphicsPath)
            m_path = Value
        End Set
    End Property

    Public Sub New()
        Me.New(System.Drawing.Color.Blue)
    End Sub

    Public Sub New(ByVal FillColor As System.Drawing.Color)
        m_color = FillColor
    End Sub

End Class
