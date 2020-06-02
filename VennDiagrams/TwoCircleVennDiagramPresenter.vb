Option Strict On

' -------------------------------------------------------------------------------
' Written by Kyle Littlefield for the Department of Energy (PNNL, Richland, WA)
' Software maintained by Matthew Monroe (PNNL, Richland, WA)
' Program started in August 2004
'
' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0
'

Public Class TwoCircleVennDiagramPresenter
    Inherits System.Windows.Forms.UserControl
    Implements ControlPrinter.IPrintableControlContainer

    Protected m_LegendSeparation As Integer = 8
    Protected m_LegendTopSeperation As Integer = 24

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
    Friend WithEvents lblLegendCircleA As System.Windows.Forms.Label
    Friend WithEvents lblLegendOverlap As System.Windows.Forms.Label
    Friend WithEvents lblLegendCircleB As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblFooter As System.Windows.Forms.Label
    Friend WithEvents gbxLegend As System.Windows.Forms.GroupBox
    Friend WithEvents lblDivider As System.Windows.Forms.Label
    Friend WithEvents vdgVennDiagram As VennDiagrams.TwoCircleVennDiagram
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label
        Me.lblLegendCircleA = New System.Windows.Forms.Label
        Me.lblLegendOverlap = New System.Windows.Forms.Label
        Me.lblLegendCircleB = New System.Windows.Forms.Label
        Me.gbxLegend = New System.Windows.Forms.GroupBox
        Me.lblDivider = New System.Windows.Forms.Label
        Me.lblFooter = New System.Windows.Forms.Label
        Me.vdgVennDiagram = New VennDiagrams.TwoCircleVennDiagram
        Me.gbxLegend.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(0, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(600, 48)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "Title goes here.  Title goes here.  Title goes here.  Title goes here.  Title goe" & _
        "s here.  Title goes here.  Title goes here.  "
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblLegendCircleA
        '
        Me.lblLegendCircleA.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLegendCircleA.Location = New System.Drawing.Point(8, -32)
        Me.lblLegendCircleA.Name = "lblLegendCircleA"
        Me.lblLegendCircleA.Size = New System.Drawing.Size(128, 112)
        Me.lblLegendCircleA.TabIndex = 2
        Me.lblLegendCircleA.Text = "Circle A"
        Me.lblLegendCircleA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLegendOverlap
        '
        Me.lblLegendOverlap.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLegendOverlap.Location = New System.Drawing.Point(8, 86)
        Me.lblLegendOverlap.Name = "lblLegendOverlap"
        Me.lblLegendOverlap.Size = New System.Drawing.Size(128, 160)
        Me.lblLegendOverlap.TabIndex = 3
        Me.lblLegendOverlap.Text = "Overlap"
        Me.lblLegendOverlap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLegendCircleB
        '
        Me.lblLegendCircleB.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLegendCircleB.Location = New System.Drawing.Point(8, 254)
        Me.lblLegendCircleB.Name = "lblLegendCircleB"
        Me.lblLegendCircleB.Size = New System.Drawing.Size(128, 162)
        Me.lblLegendCircleB.TabIndex = 4
        Me.lblLegendCircleB.Text = "Circle B"
        Me.lblLegendCircleB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'gbxLegend
        '
        Me.gbxLegend.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbxLegend.Controls.Add(Me.lblLegendCircleB)
        Me.gbxLegend.Controls.Add(Me.lblLegendOverlap)
        Me.gbxLegend.Controls.Add(Me.lblLegendCircleA)
        Me.gbxLegend.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxLegend.Location = New System.Drawing.Point(456, 56)
        Me.gbxLegend.Name = "gbxLegend"
        Me.gbxLegend.Size = New System.Drawing.Size(144, 368)
        Me.gbxLegend.TabIndex = 5
        Me.gbxLegend.TabStop = False
        Me.gbxLegend.Text = "Legend"
        '
        'lblDivider
        '
        Me.lblDivider.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDivider.BackColor = System.Drawing.Color.Black
        Me.lblDivider.Location = New System.Drawing.Point(0, 48)
        Me.lblDivider.Name = "lblDivider"
        Me.lblDivider.Size = New System.Drawing.Size(600, 1)
        Me.lblDivider.TabIndex = 7
        Me.lblDivider.Text = "Label1"
        '
        'lblFooter
        '
        Me.lblFooter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFooter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFooter.Location = New System.Drawing.Point(0, 428)
        Me.lblFooter.Name = "lblFooter"
        Me.lblFooter.Size = New System.Drawing.Size(600, 56)
        Me.lblFooter.TabIndex = 8
        Me.lblFooter.Text = "Footer ttttttttttttttttttttttttttttttttttttttttttt Footer ttttttttttttttttttttttt" & _
        "tttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer t" & _
        "tttttttttttttttttttttttttttttttttttttttttt Footer tttttttttttttttttttttttttttttt" & _
        "ttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttt" & _
        "ttttttttttttttttttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttt" & _
        "tttttt  Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttttttttt" & _
        "ttttttttttttttttttttttttttttt  Footer tttttttttttttttttttttttttttttttttttttttttt" & _
        "tFooter ttttttttttttttttttttttttttttttttttttttttttt Footer ttttttttttttttttttttt" & _
        "tttttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer" & _
        " ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttttttttttttttttttttttt" & _
        "ttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttt" & _
        "ttttttttttttttttttttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttt" & _
        "tttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer ttttttttttttt" & _
        "tttttttttttttttttttttttttttttt Footer tttttttttttttttttttttttttttttttttttttttttt" & _
        "t Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttttttttttttttt" & _
        "ttttttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt v Foo" & _
        "ter ttttttttttttttttttttttttttttttttttttttttttt Footer ttttttttttttttttttttttttt" & _
        "tttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer ttt" & _
        "tttttttttttttttttttttttttttttttttttttttt v Footer tttttttttttttttttttttttttttttt" & _
        "ttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttt" & _
        "ttttttttttttttttttttttttttttttttttt v Footer ttttttttttttttttttttttttttttttttttt" & _
        "tttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt Footer ttttttttttttt" & _
        "tttttttttttttttttttttttttttttt Footer tttttttttttttttttttttttttttttttttttttttttt" & _
        "t v Footer ttttttttttttttttttttttttttttttttttttttttttt Footer tttttttttttttttttt" & _
        "ttttttttttttttttttttttttt Footer ttttttttttttttttttttttttttttttttttttttttttt v v" & _
        ""
        '
        'vdgVennDiagram
        '
        Me.vdgVennDiagram.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vdgVennDiagram.CircleAColor = Me.vdgVennDiagram.DefaultColorCircleA
        Me.vdgVennDiagram.CircleALabel = New String() {Nothing}
        Me.vdgVennDiagram.CircleASize = 1
        Me.vdgVennDiagram.CircleBColor = Me.vdgVennDiagram.DefaultColorCircleB
        Me.vdgVennDiagram.CircleBLabel = New String() {Nothing}
        Me.vdgVennDiagram.CircleBSize = 1
        Me.vdgVennDiagram.FillFactor = 0.95
        Me.vdgVennDiagram.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vdgVennDiagram.Location = New System.Drawing.Point(0, 56)
        Me.vdgVennDiagram.Name = "vdgVennDiagram"
        Me.vdgVennDiagram.OverlapABColor = Me.vdgVennDiagram.DefaultColorOverlapAB
        Me.vdgVennDiagram.OverlapABLabel = New String() {Nothing}
        Me.vdgVennDiagram.OverlapABSize = 0.35
        Me.vdgVennDiagram.OverlapColor = Me.vdgVennDiagram.DefaultColorOverlapAB
        Me.vdgVennDiagram.OverlapLabel = New String() {Nothing}
        Me.vdgVennDiagram.OverlapSize = 0.35
        Me.vdgVennDiagram.Size = New System.Drawing.Size(448, 368)
        Me.vdgVennDiagram.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default
        Me.vdgVennDiagram.TabIndex = 9
        '
        'TwoCircleVennDiagramPresenter
        '
        Me.Controls.Add(Me.vdgVennDiagram)
        Me.Controls.Add(Me.lblFooter)
        Me.Controls.Add(Me.lblDivider)
        Me.Controls.Add(Me.gbxLegend)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "TwoCircleVennDiagramPresenter"
        Me.Size = New System.Drawing.Size(600, 480)
        Me.gbxLegend.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Public Sub DrawOnGraphics(ByVal g As Graphics, ByVal drawbackground As Boolean)
    '    Dim control As Control
    '    Dim args As PaintEventArgs = New PaintEventArgs(g, New Rectangle(Integer.MinValue, Integer.MinValue, Integer.MaxValue, Integer.MaxValue))
    '    Me.BackColor = Color.PaleGoldenrod

    '    Me.Invalidate()
    '    For Each control In Me.Controls
    '        control.Invalidate()
    '    Next
    '    'If (drawbackground) Then
    '    '    g.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(0, 0, Me.Width, Me.Height))
    '    'End If
    '    MyBase.OnPaint(args)
    '    Me.InvokePaint(Me, args)
    'End Sub

#Region " Labeling Properties "
    Public Property LegendCircleAText() As String
        Get
            Return Me.lblLegendCircleA.Text
        End Get
        Set(ByVal Value As String)
            Me.lblLegendCircleA.Text = Value
        End Set
    End Property

    Public Property LegendOverlapText() As String
        Get
            Return Me.lblLegendOverlap.Text
        End Get
        Set(ByVal Value As String)
            Me.lblLegendOverlap.Text = Value
        End Set
    End Property

    Public Property LegendCircleBText() As String
        Get
            Return Me.lblLegendCircleB.Text
        End Get
        Set(ByVal Value As String)
            Me.lblLegendCircleB.Text = Value
        End Set
    End Property

    Public ReadOnly Property Title() As Label
        Get
            Return lblTitle
        End Get
    End Property

    Public ReadOnly Property Footer() As Label
        Get
            Return Me.lblFooter
        End Get
    End Property

#End Region

    Public ReadOnly Property VennDiagram() As TwoCircleVennDiagram
        Get
            Return vdgVennDiagram
        End Get
    End Property

    'controls the amount of spacing between legend labels
    Public Property LegendSeparation() As Integer
        Get
            Return Me.m_LegendSeparation
        End Get
        Set(ByVal Value As Integer)
            Me.m_LegendSeparation = Value
        End Set
    End Property

    'update colors in legend when colors in diagram change
    Private Sub VennDiagram_ColorChange(ByVal sender As VennDiagrams.VennDiagramBaseClass) Handles vdgVennDiagram.DrawingChange
        Me.lblLegendCircleA.BackColor = vdgVennDiagram.CircleAColor
        Me.lblLegendCircleB.BackColor = vdgVennDiagram.CircleBColor
        Me.lblLegendOverlap.BackColor = vdgVennDiagram.OverlapABColor
    End Sub

    Private Sub gbxLegend_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbxLegend.Resize
        'redraw labels so they take up 1/3 of space each
        'anchoring already causes them to be aligned correctly horizontally
        Dim totalHeight As Integer = Me.gbxLegend.Height - 3 * Me.LegendSeparation - Me.m_LegendTopSeperation
        Dim individualHeight As Integer = CInt(totalHeight / 3)

        Me.lblLegendCircleA.Top = Me.m_LegendTopSeperation
        Me.lblLegendCircleA.Height = individualHeight

        Me.lblLegendOverlap.Top = Me.m_LegendTopSeperation + 1 * (Me.LegendSeparation + individualHeight)
        Me.lblLegendOverlap.Height = individualHeight

        Me.lblLegendCircleB.Top = Me.m_LegendTopSeperation + 2 * (Me.LegendSeparation + individualHeight)
        Me.lblLegendCircleB.Height = individualHeight
    End Sub
End Class
