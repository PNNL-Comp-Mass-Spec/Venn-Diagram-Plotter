Option Strict On

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
' Notice: This computer software was prepared by Battelle Memorial Institute, 
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
' Department of Energy (DOE).  All rights in the computer software are reserved 
' by DOE on behalf of the United States Government and the Contractor as 
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
' SOFTWARE.  This notice including this sentence must appear on any copies of 
' this computer software.

Public Class DisplayForm
    Inherits System.Windows.Forms.Form
    Public prevAvalue As String
    Public prevBvalue As String
    Public prevOverlapValue As String

    Private Const XML_SECTION_OPTIONS As String = "CircleParams"
    Private m_IniFileName As String = "VennDiagramPlotter.xml"

    Private Const PROGRAM_DATE As String = "October 5, 2005"

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents vdgCircles As VennDiagrams.TestPresenter
    Friend WithEvents dlgColor As System.Windows.Forms.ColorDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSizeOverlap As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeB As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeA As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCopyToClipboard As System.Windows.Forms.Button
    Friend WithEvents btnSaveToDisk As System.Windows.Forms.Button
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnCircleBColor As System.Windows.Forms.Button
    Friend WithEvents btnCircleAColor As System.Windows.Forms.Button
    Friend WithEvents btnOverlapColor As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnBackgroundColor As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents ToolTipControl As System.Windows.Forms.ToolTip
    Friend WithEvents btnExpand As System.Windows.Forms.Button
    Friend WithEvents btnShrink As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLoadDefaults As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveDefaults As System.Windows.Forms.MenuItem
    Friend WithEvents mnuReset As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSep2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditRefreshPlot As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.vdgCircles = New VennDiagrams.TestPresenter
        Me.dlgColor = New System.Windows.Forms.ColorDialog
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnBackgroundColor = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnOverlapColor = New System.Windows.Forms.Button
        Me.btnCircleBColor = New System.Windows.Forms.Button
        Me.btnCircleAColor = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSizeOverlap = New System.Windows.Forms.TextBox
        Me.txtSizeB = New System.Windows.Forms.TextBox
        Me.txtSizeA = New System.Windows.Forms.TextBox
        Me.btnExpand = New System.Windows.Forms.Button
        Me.btnShrink = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnCopyToClipboard = New System.Windows.Forms.Button
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.btnSaveToDisk = New System.Windows.Forms.Button
        Me.dlgSave = New System.Windows.Forms.SaveFileDialog
        Me.ToolTipControl = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuSaveFile = New System.Windows.Forms.MenuItem
        Me.mnuFileSep1 = New System.Windows.Forms.MenuItem
        Me.mnuLoadDefaults = New System.Windows.Forms.MenuItem
        Me.mnuSaveDefaults = New System.Windows.Forms.MenuItem
        Me.mnuReset = New System.Windows.Forms.MenuItem
        Me.mnuFileSep2 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuEdit = New System.Windows.Forms.MenuItem
        Me.mnuEditCopy = New System.Windows.Forms.MenuItem
        Me.mnuEditSep1 = New System.Windows.Forms.MenuItem
        Me.mnuEditRefreshPlot = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'vdgCircles
        '
        Me.vdgCircles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vdgCircles.Location = New System.Drawing.Point(8, 152)
        Me.vdgCircles.Name = "vdgCircles"
        Me.vdgCircles.Size = New System.Drawing.Size(448, 248)
        Me.vdgCircles.TabIndex = 2
        Me.vdgCircles.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnReset)
        Me.GroupBox1.Controls.Add(Me.btnBackgroundColor)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btnOverlapColor)
        Me.GroupBox1.Controls.Add(Me.btnCircleBColor)
        Me.GroupBox1.Controls.Add(Me.btnCircleAColor)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtSizeOverlap)
        Me.GroupBox1.Controls.Add(Me.txtSizeB)
        Me.GroupBox1.Controls.Add(Me.txtSizeA)
        Me.GroupBox1.Controls.Add(Me.btnExpand)
        Me.GroupBox1.Controls.Add(Me.btnShrink)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 136)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Circle Parameters"
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(24, 112)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(48, 20)
        Me.btnReset.TabIndex = 10
        Me.btnReset.Text = "Reset"
        '
        'btnBackgroundColor
        '
        Me.btnBackgroundColor.Location = New System.Drawing.Point(184, 112)
        Me.btnBackgroundColor.Name = "btnBackgroundColor"
        Me.btnBackgroundColor.Size = New System.Drawing.Size(40, 20)
        Me.btnBackgroundColor.TabIndex = 9
        Me.btnBackgroundColor.Text = "Color"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(88, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Background:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOverlapColor
        '
        Me.btnOverlapColor.Location = New System.Drawing.Point(184, 82)
        Me.btnOverlapColor.Name = "btnOverlapColor"
        Me.btnOverlapColor.Size = New System.Drawing.Size(40, 20)
        Me.btnOverlapColor.TabIndex = 8
        Me.btnOverlapColor.Text = "Color"
        '
        'btnCircleBColor
        '
        Me.btnCircleBColor.Location = New System.Drawing.Point(184, 50)
        Me.btnCircleBColor.Name = "btnCircleBColor"
        Me.btnCircleBColor.Size = New System.Drawing.Size(40, 20)
        Me.btnCircleBColor.TabIndex = 7
        Me.btnCircleBColor.Text = "Color"
        '
        'btnCircleAColor
        '
        Me.btnCircleAColor.Location = New System.Drawing.Point(184, 18)
        Me.btnCircleAColor.Name = "btnCircleAColor"
        Me.btnCircleAColor.Size = New System.Drawing.Size(40, 20)
        Me.btnCircleAColor.TabIndex = 6
        Me.btnCircleAColor.Text = "Color"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 16)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Circle &B:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "&Overlap:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Circle &A:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtSizeOverlap
        '
        Me.txtSizeOverlap.Location = New System.Drawing.Point(80, 82)
        Me.txtSizeOverlap.Name = "txtSizeOverlap"
        Me.txtSizeOverlap.TabIndex = 5
        Me.txtSizeOverlap.Text = "18"
        '
        'txtSizeB
        '
        Me.txtSizeB.Location = New System.Drawing.Point(80, 50)
        Me.txtSizeB.Name = "txtSizeB"
        Me.txtSizeB.TabIndex = 4
        Me.txtSizeB.Text = "25"
        '
        'txtSizeA
        '
        Me.txtSizeA.Location = New System.Drawing.Point(80, 18)
        Me.txtSizeA.Name = "txtSizeA"
        Me.txtSizeA.TabIndex = 3
        Me.txtSizeA.Text = "45"
        '
        'btnExpand
        '
        Me.btnExpand.Location = New System.Drawing.Point(8, 64)
        Me.btnExpand.Name = "btnExpand"
        Me.btnExpand.Size = New System.Drawing.Size(16, 16)
        Me.btnExpand.TabIndex = 13
        Me.btnExpand.Text = "+"
        Me.btnExpand.Visible = False
        '
        'btnShrink
        '
        Me.btnShrink.Location = New System.Drawing.Point(8, 40)
        Me.btnShrink.Name = "btnShrink"
        Me.btnShrink.Size = New System.Drawing.Size(16, 16)
        Me.btnShrink.TabIndex = 12
        Me.btnShrink.Text = "-"
        Me.btnShrink.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnCopyToClipboard)
        Me.GroupBox2.Controls.Add(Me.btnRefresh)
        Me.GroupBox2.Controls.Add(Me.btnSaveToDisk)
        Me.GroupBox2.Location = New System.Drawing.Point(256, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(112, 136)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Tasks"
        '
        'btnCopyToClipboard
        '
        Me.btnCopyToClipboard.Location = New System.Drawing.Point(8, 24)
        Me.btnCopyToClipboard.Name = "btnCopyToClipboard"
        Me.btnCopyToClipboard.Size = New System.Drawing.Size(88, 24)
        Me.btnCopyToClipboard.TabIndex = 0
        Me.btnCopyToClipboard.Text = "&Clipboard"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(8, 88)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(88, 24)
        Me.btnRefresh.TabIndex = 2
        Me.btnRefresh.Text = "&Refresh"
        '
        'btnSaveToDisk
        '
        Me.btnSaveToDisk.Location = New System.Drawing.Point(8, 56)
        Me.btnSaveToDisk.Name = "btnSaveToDisk"
        Me.btnSaveToDisk.Size = New System.Drawing.Size(88, 24)
        Me.btnSaveToDisk.TabIndex = 1
        Me.btnSaveToDisk.Text = "&Save file"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.MenuItem1})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSaveFile, Me.mnuFileSep1, Me.mnuLoadDefaults, Me.mnuSaveDefaults, Me.mnuReset, Me.mnuFileSep2, Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuSaveFile
        '
        Me.mnuSaveFile.Index = 0
        Me.mnuSaveFile.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSaveFile.Text = "&Save Plot to File"
        '
        'mnuFileSep1
        '
        Me.mnuFileSep1.Index = 1
        Me.mnuFileSep1.Text = "-"
        '
        'mnuLoadDefaults
        '
        Me.mnuLoadDefaults.Index = 2
        Me.mnuLoadDefaults.Text = "&Load Settings"
        '
        'mnuSaveDefaults
        '
        Me.mnuSaveDefaults.Index = 3
        Me.mnuSaveDefaults.Shortcut = System.Windows.Forms.Shortcut.CtrlD
        Me.mnuSaveDefaults.Text = "&Save Settings"
        '
        'mnuReset
        '
        Me.mnuReset.Index = 4
        Me.mnuReset.Text = "&Reset Settings"
        '
        'mnuFileSep2
        '
        Me.mnuFileSep2.Index = 5
        Me.mnuFileSep2.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 6
        Me.mnuExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditCopy, Me.mnuEditSep1, Me.mnuEditRefreshPlot})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Index = 0
        Me.mnuEditCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuEditCopy.Text = "&Copy Plot to Clipboard"
        '
        'mnuEditSep1
        '
        Me.mnuEditSep1.Index = 1
        Me.mnuEditSep1.Text = "-"
        '
        'mnuEditRefreshPlot
        '
        Me.mnuEditRefreshPlot.Index = 2
        Me.mnuEditRefreshPlot.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuEditRefreshPlot.Text = "&Refresh Plot"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "&Help"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "&About"
        '
        'DisplayForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 406)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.vdgCircles)
        Me.Menu = Me.MainMenu1
        Me.Name = "DisplayForm"
        Me.Text = "Venn Diagram Plotter"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnBackgroundColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackgroundColor.Click
        dlgColor.Color = Me.BackColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgCircles.VennDiagram.BackColor = dlgColor.Color
            RefreshVennDiagrams()
        End If
    End Sub

    Private Sub btnCopyToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyToClipboard.Click
        CopyVennToClipboard()
    End Sub

    Private Sub btnCircleAColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCircleAColor.Click
        'Dim tmpstr As String
        'dlgColor.Color = vdgCircles.VennDiagram.CircleAColor
        'tmpstr = vdgCircles.VennDiagram.CircleAColor.ToString
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgCircles.VennDiagram.CircleAColor = dlgColor.Color
            RefreshVennDiagrams()
        End If
    End Sub

    Private Sub btnCircleBColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCircleBColor.Click
        dlgColor.Color = vdgCircles.VennDiagram.CircleBColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgCircles.VennDiagram.CircleBColor = dlgColor.Color
            RefreshVennDiagrams()
        End If
    End Sub

    Private Sub btnExpand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpand.Click
        AutoResizeExpand()
    End Sub

    Private Sub btnOverlapColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOverlapColor.Click
        dlgColor.Color = vdgCircles.VennDiagram.OverlapColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgCircles.VennDiagram.OverlapColor = dlgColor.Color
            RefreshVennDiagrams()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        RefreshVennDiagrams()
    End Sub

    Private Sub btnSaveToDisk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveToDisk.Click
        SaveGraphicToDisk()
    End Sub

    Private Sub btnShrink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShrink.Click
        AutoResizeShrink()
    End Sub

    Private Sub DisplayForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With vdgCircles.VennDiagram
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .BackColor = dlgColor.Color.White
        End With

        SetToolTips()
        RefreshVennDiagrams()
        LoadDefaults()

        If Me.txtSizeA.Text <> "" Then
            prevAvalue = Me.txtSizeA.Text
            prevBvalue = Me.txtSizeB.Text
            prevOverlapValue = Me.txtSizeOverlap.Text
        End If

    End Sub

    Private Sub DisplayForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        vdgCircles.VennDiagram.Anchor = vdgCircles.Anchor
    End Sub

    Private Sub AutoResizeExpand()
        vdgCircles.VennDiagram.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        vdgCircles.VennDiagram.Height = CInt(vdgCircles.VennDiagram.Height + vdgCircles.VennDiagram.Height * 0.1)
        vdgCircles.Height = CInt(vdgCircles.Height + vdgCircles.Height * 0.1)
        vdgCircles.VennDiagram.Width = CInt(vdgCircles.VennDiagram.Width + vdgCircles.VennDiagram.Width * 0.1)
        vdgCircles.Width = CInt(vdgCircles.Width + vdgCircles.Width * 0.1)
    End Sub

    Private Sub AutoResizeShrink()
        vdgCircles.VennDiagram.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        vdgCircles.VennDiagram.Height = CInt(vdgCircles.VennDiagram.Height - vdgCircles.VennDiagram.Height * 0.1)
        vdgCircles.Height = CInt(vdgCircles.Height - vdgCircles.Height * 0.1)
        vdgCircles.VennDiagram.Width = CInt(vdgCircles.VennDiagram.Width - vdgCircles.VennDiagram.Width * 0.1)
        vdgCircles.Width = CInt(vdgCircles.Width - vdgCircles.Width * 0.1)
    End Sub

    Private Sub CopyVennToClipboard()
        Try
            Dim bitmap As Bitmap = New Bitmap(Me.vdgCircles.Width, Me.vdgCircles.Height)
            Dim g As Graphics = Graphics.FromImage(bitmap)

            ControlPrinter.ControlPrinter.DrawControl(Me.vdgCircles, g, True)

            Clipboard.SetDataObject(bitmap)
        Catch ex As Exception
            MsgBox("Error copying Venn diagram to clipboard: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Function GetIniFilePath(ByVal IniFileName As String) As String
        Dim fi As New System.IO.FileInfo(Application.ExecutablePath)
        Return System.IO.Path.Combine(fi.DirectoryName, IniFileName)
    End Function

    Public Function ColorToString(ByVal c As Color) As String

        Dim s As String = c.ToString()
        Dim tmpStr As String

        s = s.Split(New Char() {"["c, "]"c})(1)

        Dim strings() As String = s.Split(New Char() {"="c, ","c})
        If strings(0) <> "" Then
            tmpStr = strings(0)
        End If

        If strings.GetLength(0) > 7 Then
            s = strings(1) + "," + strings(3) + "," + strings(5) + "," + strings(7)
        End If

        Return s

    End Function

    Private Sub LoadDefaults()
        Dim outputFilepath As String = String.Empty

        Dim objXmlFile As New PRISM.Files.XmlSettingsFileAccessor
        Dim valueNotPresent As Boolean

        Try
            outputFilepath = GetIniFilePath(m_IniFileName)
            If Not System.IO.File.Exists(outputFilepath) Then
                SaveDefaults()
            End If

            With objXmlFile
                ' Pass False to .LoadSettings() here to turn off case sensitive matching
                .LoadSettings(outputFilepath, False)


                Try
                    valueNotPresent = False
                    Try
                        Me.txtSizeA.Text = .GetParam(XML_SECTION_OPTIONS, "CircleADia", 45, valueNotPresent).ToString
                    Catch ex As Exception
                        valueNotPresent = True
                    End Try

                    If valueNotPresent Then
                        Me.txtSizeA.Text = "45"
                        Me.txtSizeB.Text = "25"
                        Me.txtSizeOverlap.Text = "18"
                    Else
                        Me.txtSizeB.Text = .GetParam(XML_SECTION_OPTIONS, "CircleBDia", txtSizeA.Text).ToString

                        Try
                            Me.txtSizeOverlap.Text = (Math.Min(Double.Parse(Me.txtSizeA.Text), Double.Parse(Me.txtSizeB.Text)) / 2).ToString
                        Catch ex2 As Exception
                            Me.txtSizeOverlap.Text = "0"
                        End Try

                        Me.txtSizeOverlap.Text = .GetParam(XML_SECTION_OPTIONS, "Overlap", txtSizeOverlap.Text)
                    End If

                    Try
                        Me.vdgCircles.VennDiagram.CircleAColor = StringToColor(.GetParam(XML_SECTION_OPTIONS, "CircleAColor", ColorToString(vdgCircles.VennDiagram.CircleAColor)))
                        Me.vdgCircles.VennDiagram.CircleBColor = StringToColor(.GetParam(XML_SECTION_OPTIONS, "CircleBColor", ColorToString(vdgCircles.VennDiagram.CircleBColor)))
                        Me.vdgCircles.VennDiagram.OverlapColor = StringToColor(.GetParam(XML_SECTION_OPTIONS, "OverlapColor", ColorToString(vdgCircles.VennDiagram.OverlapColor)))
                        Me.vdgCircles.VennDiagram.BackColor = StringToColor(.GetParam(XML_SECTION_OPTIONS, "BackgroundColor", ColorToString(vdgCircles.VennDiagram.BackColor)))
                    Catch ex As Exception
                        ' Don't worry about errors here
                    End Try

                    Me.Width = .GetParam(XML_SECTION_OPTIONS, "WindowWidth", Me.Width)
                    Me.Height = .GetParam(XML_SECTION_OPTIONS, "WindowHeight", Me.Height)

                Catch ex As Exception
                    System.Windows.Forms.MessageBox.Show("Invalid parameter in settings file: " & outputFilepath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try
            End With

            RefreshVennDiagrams()

        Catch ex As Exception
            MsgBox("Error loading Defaults from " & outputFilepath & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub RefreshVennDiagrams()
        Try
            If ValidDimensionsPresent() Then
                Me.vdgCircles.VennDiagram.CircleASize = Double.Parse(Me.txtSizeA.Text)
                Me.vdgCircles.VennDiagram.CircleBSize = Double.Parse(Me.txtSizeB.Text)
                Me.vdgCircles.VennDiagram.OverlapSize = Double.Parse(Me.txtSizeOverlap.Text)
                prevAvalue = Me.txtSizeA.Text
                prevBvalue = Me.txtSizeB.Text
                prevOverlapValue = Me.txtSizeOverlap.Text
            Else
                RestorePreviousDimensions(True)
            End If
        Catch ex As Exception
            MsgBox("Error refreshing Venn diagrams: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub ResetValues()
        Me.Width = 475
        Me.Height = 440

        Me.txtSizeA.Text = "45"
        Me.txtSizeB.Text = "25"
        Me.txtSizeOverlap.Text = "18"

        vdgCircles.VennDiagram.CircleAColor = Color.FromArgb(255, 255, 192, 192)
        vdgCircles.VennDiagram.CircleBColor = Color.FromArgb(255, 192, 255, 255)
        vdgCircles.VennDiagram.BackColor = Color.FromArgb(255, 192, 255, 192)
        vdgCircles.VennDiagram.BackColor = Color.White

        RefreshVennDiagrams()
    End Sub

    Private Sub RestorePreviousDimensions(ByVal blnInformUser As Boolean)
        If blnInformUser Then
            MsgBox("The overlap value cannot be larger than circle A value or Circle B value.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Invalid numbers")
        End If

        Me.txtSizeA.Text = prevAvalue
        Me.txtSizeB.Text = prevBvalue
        Me.txtSizeOverlap.Text = prevOverlapValue

    End Sub

    Private Sub SaveDefaults()
        Dim outputFilepath As String = String.Empty
        Dim objOutFile As System.IO.StreamWriter

        Dim objXmlFile As New PRISM.Files.XmlSettingsFileAccessor

        Try
            outputFilepath = GetIniFilePath(m_IniFileName)

            If Not System.IO.File.Exists(outputFilepath) Then
                ' Need to create a new, blank XML file

                Try
                    objOutFile = System.IO.File.CreateText(outputFilepath)
                    objOutFile.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
                    objOutFile.WriteLine(" <sections>")
                    objOutFile.WriteLine(" <section name=""" & XML_SECTION_OPTIONS & """>")
                    objOutFile.WriteLine("  <item key=""CircleADia"" value=""50"" />")
                    objOutFile.WriteLine(" </section>")
                    objOutFile.WriteLine("</sections>")
                    objOutFile.Close()

                    System.Threading.Thread.Sleep(100)

                Catch ex As Exception
                    MsgBox("Error creating new Default XML file at " & outputFilepath & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                    Return
                End Try
            End If

            If ValidDimensionsPresent() Then
                With objXmlFile
                    ' Pass True to .LoadSettings() to turn on case sensitive matching
                    .LoadSettings(outputFilepath, True)

                    Try
                        .SetParam(XML_SECTION_OPTIONS, "CircleADia", txtSizeA.Text)
                        .SetParam(XML_SECTION_OPTIONS, "CircleBDia", txtSizeB.Text)
                        .SetParam(XML_SECTION_OPTIONS, "Overlap", txtSizeOverlap.Text)
                        .SetParam(XML_SECTION_OPTIONS, "CircleAColor", ColorToString(vdgCircles.VennDiagram.CircleAColor))
                        .SetParam(XML_SECTION_OPTIONS, "CircleBColor", ColorToString(vdgCircles.VennDiagram.CircleBColor))
                        .SetParam(XML_SECTION_OPTIONS, "OverlapColor", ColorToString(vdgCircles.VennDiagram.OverlapColor))
                        .SetParam(XML_SECTION_OPTIONS, "BackgroundColor", ColorToString(vdgCircles.VennDiagram.BackColor))
                        .SetParam(XML_SECTION_OPTIONS, "WindowWidth", Me.Width.ToString)
                        .SetParam(XML_SECTION_OPTIONS, "WindowHeight", Me.Height.ToString)

                        .SaveSettings()

                    Catch ex As Exception
                        System.Windows.Forms.MessageBox.Show("Error storing parameter in settings file: " & outputFilepath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End Try

                End With
            Else
                RestorePreviousDimensions(False)
            End If

        Catch ex As Exception
            MsgBox("Error saving Defaults to " & outputFilepath & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub SaveGraphicToDisk()
        Dim outputFilepath As String = String.Empty
        Dim ext As String

        Try
            Dim bitmap As Bitmap = New Bitmap(Me.vdgCircles.Width, Me.vdgCircles.Height)
            Dim g As Graphics = Graphics.FromImage(bitmap)

            dlgSave.Filter = "bmp files (*.bmp)|*.bmp|png files (*.png)|*.png"

            If dlgSave.ShowDialog = DialogResult.OK Then
                outputFilepath = dlgSave.FileName

                ControlPrinter.ControlPrinter.DrawControl(Me.vdgCircles, g, True)
                ext = Mid(outputFilepath, Len(outputFilepath) - 3).ToLower
                If ext = ".bmp" Then
                    bitmap.Save(outputFilepath, System.Drawing.Imaging.ImageFormat.Bmp)
                End If
                If ext = ".png" Then
                    bitmap.Save(outputFilepath, System.Drawing.Imaging.ImageFormat.Png)
                End If
            End If
        Catch ex As Exception
            MsgBox("Error saving Venn diagram to file " & outputFilepath & ":" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub SetToolTips()

        With ToolTipControl
            .SetToolTip(btnRefresh, "Refresh plot (Ctrl+R)")
            .SetToolTip(btnSaveToDisk, "Save plot to disk (Ctrl+S)")
            .SetToolTip(btnCopyToClipboard, "Copy plot to clipboard (Ctrl+C)")
            .SetToolTip(btnReset, "Reset settings to defaults")
        End With

    End Sub

    Private Sub ShowAboutBox()
        Dim strMessage As String

        strMessage = String.Empty
        strMessage &= "Program written by Kyle Littlefield for the Department of Energy (PNNL, Richland, WA) in 2004" & ControlChars.NewLine
        strMessage &= "Software maintained by Matthew Monroe (PNNL, Richland, WA)" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "This is version " & System.Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com" & ControlChars.NewLine
        strMessage &= "Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
        strMessage &= "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "Notice: This computer software was prepared by Battelle Memorial Institute, "
        strMessage &= "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the "
        strMessage &= "Department of Energy (DOE).  All rights in the computer software are reserved "
        strMessage &= "by DOE on behalf of the United States Government and the Contractor as "
        strMessage &= "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY "
        strMessage &= "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS "
        strMessage &= "SOFTWARE.  This notice including this sentence must appear on any copies of "
        strMessage &= "this computer software." & ControlChars.NewLine

        Windows.Forms.MessageBox.Show(strMessage, "About", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Public Function StringToColor(ByVal s As String) As Color

        Return CType(System.ComponentModel.TypeDescriptor.GetConverter(GetType(Color)).ConvertFromString(s), Color)

    End Function

    Private Function ValidDimensionsPresent() As Boolean
        Try
            Return Not (Double.Parse(Me.txtSizeOverlap.Text) > Double.Parse(Me.txtSizeA.Text) Or Double.Parse(Me.txtSizeOverlap.Text) > Double.Parse(Me.txtSizeB.Text))
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        ResetValues()
    End Sub

    Private Sub txtSizeA_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeA.LostFocus
        If ValidDimensionsPresent() Then RefreshVennDiagrams()
    End Sub

    Private Sub txtSizeB_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeB.LostFocus
        If ValidDimensionsPresent() Then RefreshVennDiagrams()
    End Sub

    Private Sub txtSizeOverlap_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeOverlap.LostFocus
        If ValidDimensionsPresent() Then RefreshVennDiagrams()
    End Sub

    Private Sub mnuSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveFile.Click
        SaveGraphicToDisk()
    End Sub

    Private Sub mnuLoadDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLoadDefaults.Click
        LoadDefaults()
    End Sub

    Private Sub mnuSaveDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveDefaults.Click
        SaveDefaults()
    End Sub

    Private Sub mnuReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReset.Click
        Reset()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub mnuEditCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditCopy.Click
        CopyVennToClipboard()
    End Sub

    Private Sub mnuEditRefreshPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditRefreshPlot.Click
        RefreshVennDiagrams()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        ShowAboutBox()
    End Sub

End Class
