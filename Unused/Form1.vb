Imports System.Drawing.Imaging
Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "
    Dim c As Char = ChrW(8209)
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'Me.TwoCircleVennDiagram1.CircleBColor = Color.LightPink
        'Me.TwoCircleVennDiagram1.CircleAColor = Color.LightBlue
        'Me.TwoCircleVennDiagram1.PrepareForDrawing()

        Me.TwoCircleVennDiagram1.Title.Text = "Title"
        'Me.TwoCircleVennDiagram1.Title.Font = New Font("Arial Unicode MS", 24, 0, GraphicsUnit.Point)
        'Me.TwoCircleVennDiagram1.Title.TextAlign = ContentAlignment.MiddleCenter
        Me.TwoCircleVennDiagram1.Title.Text = "Some" & c & "multiline text if I can type long enough to cause it to happen.  And that was" & c & "notenough maybe add some more text?"
    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents txtSizeA As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeOverlap As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeB As System.Windows.Forms.TextBox
    Friend WithEvents cmdUpdateVenn As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TwoCircleVennDiagram1 As VennDiagrams.TwoCircleVennDiagramPresenter
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtSizeA = New System.Windows.Forms.TextBox
        Me.txtSizeOverlap = New System.Windows.Forms.TextBox
        Me.txtSizeB = New System.Windows.Forms.TextBox
        Me.cmdUpdateVenn = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.TwoCircleVennDiagram1 = New VennDiagrams.TwoCircleVennDiagramPresenter
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSizeA
        '
        Me.txtSizeA.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSizeA.Location = New System.Drawing.Point(832, 8)
        Me.txtSizeA.Name = "txtSizeA"
        Me.txtSizeA.Size = New System.Drawing.Size(80, 20)
        Me.txtSizeA.TabIndex = 5
        Me.txtSizeA.Text = "1000"
        '
        'txtSizeOverlap
        '
        Me.txtSizeOverlap.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSizeOverlap.Location = New System.Drawing.Point(832, 32)
        Me.txtSizeOverlap.Name = "txtSizeOverlap"
        Me.txtSizeOverlap.Size = New System.Drawing.Size(80, 20)
        Me.txtSizeOverlap.TabIndex = 6
        Me.txtSizeOverlap.Text = "250"
        '
        'txtSizeB
        '
        Me.txtSizeB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSizeB.Location = New System.Drawing.Point(832, 56)
        Me.txtSizeB.Name = "txtSizeB"
        Me.txtSizeB.Size = New System.Drawing.Size(80, 20)
        Me.txtSizeB.TabIndex = 7
        Me.txtSizeB.Text = "1000"
        '
        'cmdUpdateVenn
        '
        Me.cmdUpdateVenn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdateVenn.Location = New System.Drawing.Point(832, 80)
        Me.cmdUpdateVenn.Name = "cmdUpdateVenn"
        Me.cmdUpdateVenn.Size = New System.Drawing.Size(80, 23)
        Me.cmdUpdateVenn.TabIndex = 8
        Me.cmdUpdateVenn.Text = "Enter Values"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.TwoCircleVennDiagram1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(816, 632)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(344, 592)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Button3"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.ForeColor = System.Drawing.Color.Red
        Me.Button2.Location = New System.Drawing.Point(256, 592)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        '
        'TwoCircleVennDiagram1
        '
        Me.TwoCircleVennDiagram1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TwoCircleVennDiagram1.LegendCircleAText = "Circle A Followed by a lot of text that will test the layout capabilities of the " & _
        "printer.  And we need some more"
        Me.TwoCircleVennDiagram1.LegendCircleBText = "Circle B"
        Me.TwoCircleVennDiagram1.LegendOverlapText = "Overlap"
        Me.TwoCircleVennDiagram1.LegendSeparation = 8
        Me.TwoCircleVennDiagram1.Location = New System.Drawing.Point(8, 16)
        Me.TwoCircleVennDiagram1.Name = "TwoCircleVennDiagram1"
        Me.TwoCircleVennDiagram1.Size = New System.Drawing.Size(800, 568)
        Me.TwoCircleVennDiagram1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(832, 112)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 24)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Export"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(920, 650)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdUpdateVenn)
        Me.Controls.Add(Me.txtSizeB)
        Me.Controls.Add(Me.txtSizeOverlap)
        Me.Controls.Add(Me.txtSizeA)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.ExperimentationDiagram1.Width = Me.ExperimentationDiagram1.Width + 10
        Me.TwoCircleVennDiagram1.Width = Me.TwoCircleVennDiagram1.Width + 10
        'Console.WriteLine(Me.ExperimentationDiagram1.Width)
        Me.Refresh()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.ExperimentationDiagram1.LineOn = Not Me.ExperimentationDiagram1.LineOn
        Me.Refresh()
    End Sub

    Private Sub cmdUpdateVenn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateVenn.Click
        Me.TwoCircleVennDiagram1.VennDiagram.CircleASize = Double.Parse(Me.txtSizeA.Text)
        Me.TwoCircleVennDiagram1.VennDiagram.CircleBSize = Double.Parse(Me.txtSizeB.Text)
        Me.TwoCircleVennDiagram1.VennDiagram.OverlapSize = Double.Parse(Me.txtSizeOverlap.Text)
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim bitmap As Bitmap = New Bitmap(Me.GroupBox1.Width, Me.GroupBox1.Height)
        Dim g As Graphics = Graphics.FromImage(bitmap)
        Dim newImage As ImageFormat
        'g.SetClip(New Rectangle(0, 0, 30, 1000), Drawing2D.CombineMode.Exclude)

        'Me.TwoCircleVennDiagram1.DrawOnGraphics(Graphics.FromImage(bitmap), True)
        ControlPrinter.ControlPrinter.DrawControl(Me.GroupBox1, g, True)
        bitmap.Save("Output.png")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Button2.Enabled = Not Me.Button2.Enabled
    End Sub
End Class
