Option Strict On

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
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
Public Class ThreeCircleVennDiagramPlain
    Inherits System.Windows.Forms.UserControl
    Implements ControlPrinter.IPrintableControlContainer


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
    Friend WithEvents vdgVennDiagram As VennDiagrams.ThreeCircleVennDiagram
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.vdgVennDiagram = New VennDiagrams.ThreeCircleVennDiagram
        Me.SuspendLayout()
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
        Me.vdgVennDiagram.CircleCColor = Me.vdgVennDiagram.DefaultColorCircleC
        Me.vdgVennDiagram.CircleCLabel = New String() {Nothing}
        Me.vdgVennDiagram.CircleCSize = 1
        Me.vdgVennDiagram.FillFactor = 0.95
        Me.vdgVennDiagram.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vdgVennDiagram.Location = New System.Drawing.Point(0, 0)
        Me.vdgVennDiagram.Name = "vdgVennDiagram"
        Me.vdgVennDiagram.OverlapABColor = Me.vdgVennDiagram.DefaultColorOverlapAB
        Me.vdgVennDiagram.OverlapBCColor = Me.vdgVennDiagram.DefaultColorOverlapBC
        Me.vdgVennDiagram.OverlapACColor = Me.vdgVennDiagram.DefaultColorOverlapAC
        Me.vdgVennDiagram.OverlapABCColor = Me.vdgVennDiagram.DefaultColorOverlapABC
        Me.vdgVennDiagram.OverlapABLabel = New String() {Nothing}
        Me.vdgVennDiagram.OverlapABSize = 0.35
        Me.vdgVennDiagram.OverlapBCSize = 0.15
        Me.vdgVennDiagram.OverlapACSize = 0.2
        Me.vdgVennDiagram.Size = New System.Drawing.Size(280, 232)
        Me.vdgVennDiagram.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default
        Me.vdgVennDiagram.TabIndex = 0
        '
        'ThreeCircleVennDiagramPlain
        '
        Me.Controls.Add(Me.vdgVennDiagram)
        Me.Name = "ThreeCircleVennDiagramPlain"
        Me.Size = New System.Drawing.Size(280, 232)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public ReadOnly Property VennDiagram() As ThreeCircleVennDiagram
        Get
            Return vdgVennDiagram
        End Get
    End Property

End Class
