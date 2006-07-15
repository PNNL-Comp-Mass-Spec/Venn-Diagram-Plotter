Imports System.Drawing

Public Class TwoCircleVennDiagram_Old
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
    Protected m_Title As String = "Venn Diagram"
    Protected m_TitleFont As Font = New Font("Arial", 14)
    Protected m_ReservedTitleSize As Integer = 25
    Protected m_ReservedLabelSize As Integer = 20

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
        Console.WriteLine("Painting")
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

#Region "Properties "
    Public Property CircleASize() As Double
        Get
            Return Me.m_circleA.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_circleA.Size = Value
            Me.m_ScreenCoordinatesValid = False
            Me.m_ComputeWorldCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property CircleBSize() As Double
        Get
            Return Me.m_circleB.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_circleB.Size = Value
            Me.m_ScreenCoordinatesValid = False
            Me.m_ComputeWorldCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property OverlapSize() As Double
        Get
            Return Me.m_overlap.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_overlap.Size = Value
            Me.m_ScreenCoordinatesValid = False
            Me.m_ComputeWorldCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property CircleALabel() As String()
        Get
            Return Me.m_circleA.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_circleA.Label = Value
            Me.Invalidate()
        End Set
    End Property

    Public Property CircleBLabel() As String()
        Get
            Return Me.m_circleB.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_circleB.Label = Value
            Me.Invalidate()
        End Set
    End Property

    Public Property OverlapLabel() As String()
        Get
            Return Me.m_overlap.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_overlap.Label = Value
            Me.Invalidate()
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
        End Set
    End Property

    Public Property OverlapColor() As Color
        Get
            Return Me.m_overlap.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_overlap.Color = Value
            Me.Invalidate()
        End Set
    End Property

    'Protected Sub UpdateOverlapColor()
    '    Me.m_overlap.Color = Me.MergeTwoColors(Me.CircleAColor, Me.CircleBColor)
    'End Sub

    Public Property OutlinePen() As Pen
        Get
            Return Me.m_outlinePen
        End Get
        Set(ByVal Value As Pen)
            Me.m_outlinePen = Value
            Me.Invalidate()
        End Set
    End Property

    Public Property FillSize() As Double
        Get
            Return Me.m_fillSize
        End Get
        Set(ByVal Value As Double)
            If (Value < 0 Or Value > 1.0) Then
                Throw New ArgumentException("Must be in range of 0 to 1")
            End If
            Me.m_fillSize = Value
            Me.m_ScreenCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property SmoothingMode() As Drawing2D.SmoothingMode
        Get
            Return Me.m_smoothingMode
        End Get
        Set(ByVal Value As Drawing2D.SmoothingMode)
            Value = Me.m_smoothingMode
        End Set
    End Property

    Public Property Title() As String
        Get
            Return Me.m_Title
        End Get
        Set(ByVal Value As String)
            Me.m_Title = Value
            Me.m_ScreenCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property TitleFont() As Font
        Get
            Return Me.m_TitleFont
        End Get
        Set(ByVal Value As Font)
            Me.m_TitleFont = Value
            Me.m_ScreenCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property ReservedTitleArea() As Integer
        Get
            Return Me.m_ReservedTitleSize
        End Get
        Set(ByVal Value As Integer)
            Me.m_ReservedTitleSize = Value
            Me.m_ScreenCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    Public Property ReservedLabelArea() As Integer
        Get
            Return Me.m_ReservedLabelSize
        End Get
        Set(ByVal Value As Integer)
            Me.m_ReservedLabelSize = Value
            Me.m_ScreenCoordinatesValid = False
            Me.Invalidate()
        End Set
    End Property

    ReadOnly Property ReservedTopArea() As Integer
        Get
            Return Me.ReservedTitleArea + Me.ReservedLabelArea
        End Get
    End Property
#End Region

    Protected Function MergeTwoColors(ByVal colorA As Color, ByVal colorB As Color) As Color
        Dim alpha As Byte
        Dim red As Byte
        Dim green As Byte
        Dim blue As Byte

        alpha = CByte((CInt(colorA.A) + CInt(colorB.A)) / 2)
        red = CByte((CInt(colorA.R) + CInt(colorB.R)) / 2)
        green = CByte((CInt(colorA.G) + CInt(colorB.G)) / 2)
        blue = CByte((CInt(colorA.B) + CInt(colorB.B)) / 2)

        Return Color.FromArgb(alpha, red, green, blue)
    End Function

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

    Protected Sub ComputeScreenCoordinatesFromWorldCoordinates()
        If Not (Me.m_ComputeWorldCoordinatesValid) Then
            Me.ComputeComputeCoordinates()
        End If

        Dim computeWorldWidth As Double = Me.m_B_Center_Compute + Me.m_B_Radius_Compute + Me.m_A_Radius_Compute
        Dim computeWorldHeight As Double = Math.Max(2 * Me.m_B_Radius_Compute, 2 * Me.m_A_Radius_Compute)
        Dim heightScaleFactor As Double = (Me.Height - Me.ReservedTopArea) * Me.FillSize / computeWorldHeight
        Dim widthScaleFactor As Double = Me.Width * Me.FillSize / computeWorldWidth
        Dim scaleFactor As Double = Math.Min(heightScaleFactor, widthScaleFactor)

        Dim circleAWidth As Double = (2 * Me.m_A_Radius_Compute) * scaleFactor
        Dim circleBWidth As Double = (2 * Me.m_B_Radius_Compute) * scaleFactor
        Dim screenWidth As Double = computeWorldWidth * scaleFactor

        Dim xInset As Double = (Me.Width - screenWidth) / 2

        Dim circleALeft As Double = xInset
        Dim circleBLeft As Double = xInset + (Me.m_A_Radius_Compute + Me.m_B_Center_Compute - Me.m_B_Radius_Compute) * scaleFactor
        Dim circleATop As Double = (Me.Height - Me.ReservedTopArea) / 2 + Me.ReservedTopArea - (Me.m_A_Radius_Compute * scaleFactor)
        Dim circleBTop As Double = (Me.Height - Me.ReservedTopArea) / 2 + Me.ReservedTopArea - (Me.m_B_Radius_Compute * scaleFactor)


        Dim circleABoundingBox As RectangleF = New RectangleF(CSng(circleALeft), CSng(circleATop), CSng(circleAWidth), CSng(circleAWidth))
        Dim circleBBoundingBox As RectangleF = New RectangleF(CSng(circleBLeft), CSng(circleBTop), CSng(circleBWidth), CSng(circleBWidth))

        'Console.WriteLine("Computing Screen Coordinates")

        'TODO, test is just copy world coordinates to screen coordinates
        'screen coordinates don't exist as class level variables, they are inherently part
        'of the DrawingPath property of VennDiagramAreaInfo
        Me.m_circleA.DrawingPath.Reset()
        Me.m_circleA.DrawingPath.AddEllipse(circleABoundingBox)

        Me.m_circleB.DrawingPath.Reset()
        Me.m_circleB.DrawingPath.AddEllipse(circleBBoundingBox)

        Me.m_overlap.DrawingPath.Reset()
        Me.m_overlap.DrawingPath.AddArc(circleABoundingBox, -ConvertAngleToDegrees(Me.m_alpha / 2), ConvertAngleToDegrees(Me.m_alpha))
        Me.m_overlap.DrawingPath.AddArc(circleBBoundingBox, 180 - ConvertAngleToDegrees(Me.m_beta / 2), ConvertAngleToDegrees(Me.m_beta))

        'Console.WriteLine("calculated screen coordinates")
        Me.m_ScreenCoordinatesValid = True
    End Sub

    Protected Function ConvertAngleToDegrees(ByVal radians As Double) As Single
        Return CSng(radians * 180 / Math.PI)
    End Function

    'must be called after setting properties, before it can be drawn
    Public Sub ComputeComputeCoordinates()
        'calculate positions/shapes of circles and overlap
        'implements the algorithm described in http://www.cs.uvic.ca/~ruskey/Publications/VennArea/VennArea.pdf
        Dim high As Double
        Dim low As Double
        Dim currentCenter As Double
        Dim tries As Integer = 0

        Me.m_A_Radius_Compute = Math.Sqrt(Me.m_circleA.Size / Math.PI)
        Me.m_B_Radius_Compute = Math.Sqrt(Me.m_circleB.Size / Math.PI)

        'start trials at no overlap
        high = Me.m_A_Radius_Compute + Me.m_B_Radius_Compute
        low = Math.Abs(Me.m_A_Radius_Compute - Me.m_B_Radius_Compute)

        Do
            'area for current center - desired area of overlap
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

            overlap = (1 / 2 * r1 * r1) * (alpha - Math.Sin(alpha))
            overlap = overlap + (1 / 2 * r2 * r2) * (beta - Math.Sin(beta))

            overlapDifference = overlap - Me.m_overlap.Size

            If overlapDifference > -Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 AndAlso overlapDifference < Math.Max(Me.m_circleA.Size, Me.m_circleB.Size) / 100000000000 Then
                Me.m_B_Center_Compute = currentCenter
                Me.m_alpha = alpha
                Me.m_beta = beta 'could recompute this outside of loop
                Exit Do
            End If
            If overlapDifference > 0 Then
                low = currentCenter
            ElseIf overlapDifference < 0 Then
                high = currentCenter
            End If
            tries = tries + 1
            If tries > Me.MAX_COMPUTE_TRIES Then    'This should only happen in the case that the user 
                'specifies an overlap that is bigger than one or both of the base sets.
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

        'draw circle and overlap, starting with circleA and circleB, then with overlap so that it draws over them
        Dim g As Graphics = e.Graphics

        'Me.DrawOnGraphics(g, False)
        If Not Me.m_ScreenCoordinatesValid Then
            Me.ComputeScreenCoordinatesFromWorldCoordinates()
        End If

        g.SmoothingMode = Me.SmoothingMode

        Me.PaintVennDiagramInfo(Me.m_circleA, g)
        Me.PaintVennDiagramInfo(Me.m_circleB, g)
        Me.PaintVennDiagramInfo(Me.m_overlap, g)
        'deal with labels - TODO
        'g.DrawString(Me.m_circleA.Label, Me.Font, New SolidBrush(Me.ForeColor), 0, 50)
        Me.DrawLabels(g)
    End Sub

    Protected Sub DrawLabels(ByVal g As Graphics)
        Dim titleSize As SizeF = g.MeasureString(Me.m_Title, Me.m_TitleFont)

        g.DrawString(Me.Title, Me.TitleFont, New SolidBrush(Me.ForeColor), CSng((Me.Width / 2) - (titleSize.Width / 2)), 1)
        g.DrawLine(New System.Drawing.Pen(Me.ForeColor, 2), 0, Me.ReservedTitleArea - 1, Me.Width, Me.ReservedTitleArea - 1)
        Me.DrawCircleALabel(g)
        Me.DrawCircleBLabel(g)
    End Sub

    Protected Sub DrawCircleALabel(ByVal g As Graphics)
        Dim i As Integer

        For i = 0 To Me.m_circleA.Label.Length - 1
            Dim s As String = Me.m_circleA.Label(i)

            g.DrawString(s, Me.Font, New SolidBrush(Me.ForeColor), CSng(Me.Width * (1 - FillSize)) / 2, Me.ReservedTitleArea + Font.Height * i)
        Next
    End Sub


    Protected Sub DrawCircleBLabel(ByVal g As Graphics)
        Dim i As Integer

        For i = 0 To Me.m_circleB.Label.Length - 1
            Dim s As String = Me.m_circleB.Label(i)
            Dim labelSize As SizeF = g.MeasureString(s, Me.Font)

            g.DrawString(s, Me.Font, New SolidBrush(Me.ForeColor), Me.Width - (CSng(Me.Width * (1 - FillSize)) / 2) - labelSize.Width, Me.ReservedTitleArea + Font.Height * i)
        Next
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
            Me.ComputeScreenCoordinatesFromWorldCoordinates()
            Me.Invalidate()
        End If
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

