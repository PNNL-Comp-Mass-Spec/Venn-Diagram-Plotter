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

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InitializeControls()

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
    Friend WithEvents dlgColor As System.Windows.Forms.ColorDialog
    Friend WithEvents txtSizeB As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeA As System.Windows.Forms.TextBox
    Friend WithEvents cmdCopyToClipboard As System.Windows.Forms.Button
    Friend WithEvents cmdSaveToDisk As System.Windows.Forms.Button
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents ToolTipControl As System.Windows.Forms.ToolTip
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
    Friend WithEvents txtDistinctA As System.Windows.Forms.TextBox
    Friend WithEvents txtDistinctB As System.Windows.Forms.TextBox
    Friend WithEvents optTotal As System.Windows.Forms.RadioButton
    Friend WithEvents optCountDistinct As System.Windows.Forms.RadioButton
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditResetValues As System.Windows.Forms.MenuItem
    Friend WithEvents lblCircleB As System.Windows.Forms.Label
    Friend WithEvents lblOverlap As System.Windows.Forms.Label
    Friend WithEvents lblCircleA As System.Windows.Forms.Label
    Friend WithEvents txtSizeOverlapAB As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeC As System.Windows.Forms.TextBox
    Friend WithEvents chkCircleC As System.Windows.Forms.CheckBox
    Friend WithEvents txtSizeOverlapBC As System.Windows.Forms.TextBox
    Friend WithEvents txtSizeOverlapAC As System.Windows.Forms.TextBox
    Friend WithEvents fraParameters As System.Windows.Forms.GroupBox
    Friend WithEvents fraThreeCircleRegionCounts As System.Windows.Forms.GroupBox
    Friend WithEvents txtRegionAx As System.Windows.Forms.TextBox
    Friend WithEvents lblRegionAx As System.Windows.Forms.Label
    Friend WithEvents txtRegionCx As System.Windows.Forms.TextBox
    Friend WithEvents txtRegionBx As System.Windows.Forms.TextBox
    Friend WithEvents txtRegionABC As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtRegionABx As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRegionBCx As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtRegionACx As System.Windows.Forms.TextBox
    Friend WithEvents cboRegionCountMode As System.Windows.Forms.ComboBox
    Friend WithEvents txtRegionCountValue As System.Windows.Forms.TextBox
    Friend WithEvents cmdRegionCountComputeOptimal As System.Windows.Forms.Button
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents lblRegionUniqueItemCountAllCircles As System.Windows.Forms.Label
    Friend WithEvents txtRegionUniqueItemCountAllCircles As System.Windows.Forms.TextBox
    Friend WithEvents cmdUpdateRegionCounts As System.Windows.Forms.Button
    Friend WithEvents fraMessageDisplayOptions As System.Windows.Forms.GroupBox
    Friend WithEvents tbarDuplicateMessageIgnoreWindowSeconds As System.Windows.Forms.TrackBar
    Friend WithEvents tbarMessageDisplayTimeSeconds As System.Windows.Forms.TrackBar
    Friend WithEvents lblMessageDuplicateIgnoreWindowCaption As System.Windows.Forms.Label
    Friend WithEvents lblMessageDisplayTimeSecondsCaption As System.Windows.Forms.Label
    Friend WithEvents lblMessageDisplayTimeSeconds As System.Windows.Forms.Label
    Friend WithEvents lblMessageDuplicateIgnoreWindow As System.Windows.Forms.Label
    Friend WithEvents chkHideMessagesOnSuccessfulUpdate As System.Windows.Forms.CheckBox
    Friend WithEvents lblOverlapAB As System.Windows.Forms.Label
    Friend WithEvents lblOverlapBC As System.Windows.Forms.Label
    Friend WithEvents lblOverlapAC As System.Windows.Forms.Label
    Friend WithEvents fraImageAdjustmentControls As System.Windows.Forms.GroupBox
    Friend WithEvents lblImgRotation As System.Windows.Forms.Label
    Friend WithEvents tbarImgRotation As System.Windows.Forms.TrackBar
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents vdgThreeCircles As VennDiagrams.ThreeCircleVennDiagramPlain
    Friend WithEvents vdgTwoCircles As VennDiagrams.TwoCircleVennDiagramPlain
    Friend WithEvents tbarImgXOffset As System.Windows.Forms.TrackBar
    Friend WithEvents tbarImgZoom As System.Windows.Forms.TrackBar
    Friend WithEvents lblImgZoom As System.Windows.Forms.Label
    Friend WithEvents lblImgOffsetX As System.Windows.Forms.Label
    Friend WithEvents tbarImgYOffset As System.Windows.Forms.TrackBar
    Friend WithEvents lblImgOffsetY As System.Windows.Forms.Label
    Friend WithEvents cmdImgAdjustmentReset As System.Windows.Forms.Button
    Friend WithEvents pnlColorButtons As System.Windows.Forms.Panel
    Friend WithEvents cmdOverlapABCColor As System.Windows.Forms.Button
    Friend WithEvents cmdOverlapACColor As System.Windows.Forms.Button
    Friend WithEvents cmdOverlapBCColor As System.Windows.Forms.Button
    Friend WithEvents cmdOverlapABColor As System.Windows.Forms.Button
    Friend WithEvents cmdCircleCColor As System.Windows.Forms.Button
    Friend WithEvents cmdBackgroundColor As System.Windows.Forms.Button
    Friend WithEvents cmdCircleBColor As System.Windows.Forms.Button
    Friend WithEvents cmdCircleAColor As System.Windows.Forms.Button
    Friend WithEvents chkFillCirclesWithColor As System.Windows.Forms.CheckBox
    Friend WithEvents fraTrasks As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.dlgColor = New System.Windows.Forms.ColorDialog
        Me.fraParameters = New System.Windows.Forms.GroupBox
        Me.pnlColorButtons = New System.Windows.Forms.Panel
        Me.cmdOverlapABCColor = New System.Windows.Forms.Button
        Me.cmdOverlapACColor = New System.Windows.Forms.Button
        Me.cmdOverlapBCColor = New System.Windows.Forms.Button
        Me.cmdOverlapABColor = New System.Windows.Forms.Button
        Me.cmdCircleCColor = New System.Windows.Forms.Button
        Me.cmdBackgroundColor = New System.Windows.Forms.Button
        Me.cmdCircleBColor = New System.Windows.Forms.Button
        Me.cmdCircleAColor = New System.Windows.Forms.Button
        Me.lblOverlapAC = New System.Windows.Forms.Label
        Me.lblOverlapBC = New System.Windows.Forms.Label
        Me.lblOverlapAB = New System.Windows.Forms.Label
        Me.chkCircleC = New System.Windows.Forms.CheckBox
        Me.txtSizeOverlapAC = New System.Windows.Forms.TextBox
        Me.txtSizeOverlapBC = New System.Windows.Forms.TextBox
        Me.txtSizeC = New System.Windows.Forms.TextBox
        Me.txtDistinctA = New System.Windows.Forms.TextBox
        Me.txtDistinctB = New System.Windows.Forms.TextBox
        Me.lblCircleB = New System.Windows.Forms.Label
        Me.lblOverlap = New System.Windows.Forms.Label
        Me.lblCircleA = New System.Windows.Forms.Label
        Me.txtSizeOverlapAB = New System.Windows.Forms.TextBox
        Me.txtSizeB = New System.Windows.Forms.TextBox
        Me.txtSizeA = New System.Windows.Forms.TextBox
        Me.optCountDistinct = New System.Windows.Forms.RadioButton
        Me.optTotal = New System.Windows.Forms.RadioButton
        Me.fraTrasks = New System.Windows.Forms.GroupBox
        Me.chkFillCirclesWithColor = New System.Windows.Forms.CheckBox
        Me.cmdCopyToClipboard = New System.Windows.Forms.Button
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.cmdSaveToDisk = New System.Windows.Forms.Button
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
        Me.mnuEditRefreshPlot = New System.Windows.Forms.MenuItem
        Me.mnuEditSep1 = New System.Windows.Forms.MenuItem
        Me.mnuEditResetValues = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
        Me.fraThreeCircleRegionCounts = New System.Windows.Forms.GroupBox
        Me.cmdUpdateRegionCounts = New System.Windows.Forms.Button
        Me.txtRegionUniqueItemCountAllCircles = New System.Windows.Forms.TextBox
        Me.lblRegionUniqueItemCountAllCircles = New System.Windows.Forms.Label
        Me.cmdRegionCountComputeOptimal = New System.Windows.Forms.Button
        Me.txtRegionCountValue = New System.Windows.Forms.TextBox
        Me.cboRegionCountMode = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtRegionACx = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtRegionBCx = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtRegionABx = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblRegionAx = New System.Windows.Forms.Label
        Me.txtRegionABC = New System.Windows.Forms.TextBox
        Me.txtRegionCx = New System.Windows.Forms.TextBox
        Me.txtRegionAx = New System.Windows.Forms.TextBox
        Me.txtRegionBx = New System.Windows.Forms.TextBox
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.fraMessageDisplayOptions = New System.Windows.Forms.GroupBox
        Me.chkHideMessagesOnSuccessfulUpdate = New System.Windows.Forms.CheckBox
        Me.lblMessageDuplicateIgnoreWindow = New System.Windows.Forms.Label
        Me.lblMessageDisplayTimeSeconds = New System.Windows.Forms.Label
        Me.lblMessageDuplicateIgnoreWindowCaption = New System.Windows.Forms.Label
        Me.lblMessageDisplayTimeSecondsCaption = New System.Windows.Forms.Label
        Me.tbarDuplicateMessageIgnoreWindowSeconds = New System.Windows.Forms.TrackBar
        Me.tbarMessageDisplayTimeSeconds = New System.Windows.Forms.TrackBar
        Me.fraImageAdjustmentControls = New System.Windows.Forms.GroupBox
        Me.cmdImgAdjustmentReset = New System.Windows.Forms.Button
        Me.lblImgOffsetY = New System.Windows.Forms.Label
        Me.tbarImgYOffset = New System.Windows.Forms.TrackBar
        Me.lblImgOffsetX = New System.Windows.Forms.Label
        Me.lblImgZoom = New System.Windows.Forms.Label
        Me.tbarImgXOffset = New System.Windows.Forms.TrackBar
        Me.tbarImgZoom = New System.Windows.Forms.TrackBar
        Me.lblImgRotation = New System.Windows.Forms.Label
        Me.tbarImgRotation = New System.Windows.Forms.TrackBar
        Me.vdgThreeCircles = New VennDiagrams.ThreeCircleVennDiagramPlain
        Me.vdgTwoCircles = New VennDiagrams.TwoCircleVennDiagramPlain
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.fraParameters.SuspendLayout()
        Me.pnlColorButtons.SuspendLayout()
        Me.fraTrasks.SuspendLayout()
        Me.fraThreeCircleRegionCounts.SuspendLayout()
        Me.fraMessageDisplayOptions.SuspendLayout()
        CType(Me.tbarDuplicateMessageIgnoreWindowSeconds, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tbarMessageDisplayTimeSeconds, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraImageAdjustmentControls.SuspendLayout()
        CType(Me.tbarImgYOffset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tbarImgXOffset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tbarImgZoom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tbarImgRotation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraParameters
        '
        Me.fraParameters.Controls.Add(Me.pnlColorButtons)
        Me.fraParameters.Controls.Add(Me.lblOverlapAC)
        Me.fraParameters.Controls.Add(Me.lblOverlapBC)
        Me.fraParameters.Controls.Add(Me.lblOverlapAB)
        Me.fraParameters.Controls.Add(Me.chkCircleC)
        Me.fraParameters.Controls.Add(Me.txtSizeOverlapAC)
        Me.fraParameters.Controls.Add(Me.txtSizeOverlapBC)
        Me.fraParameters.Controls.Add(Me.txtSizeC)
        Me.fraParameters.Controls.Add(Me.txtDistinctA)
        Me.fraParameters.Controls.Add(Me.txtDistinctB)
        Me.fraParameters.Controls.Add(Me.lblCircleB)
        Me.fraParameters.Controls.Add(Me.lblOverlap)
        Me.fraParameters.Controls.Add(Me.lblCircleA)
        Me.fraParameters.Controls.Add(Me.txtSizeOverlapAB)
        Me.fraParameters.Controls.Add(Me.txtSizeB)
        Me.fraParameters.Controls.Add(Me.txtSizeA)
        Me.fraParameters.Controls.Add(Me.optCountDistinct)
        Me.fraParameters.Controls.Add(Me.optTotal)
        Me.fraParameters.Location = New System.Drawing.Point(8, 8)
        Me.fraParameters.Name = "fraParameters"
        Me.fraParameters.Size = New System.Drawing.Size(472, 140)
        Me.fraParameters.TabIndex = 0
        Me.fraParameters.TabStop = False
        Me.fraParameters.Text = "Circle Parameters"
        '
        'pnlColorButtons
        '
        Me.pnlColorButtons.Controls.Add(Me.cmdOverlapABCColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdOverlapACColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdOverlapBCColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdOverlapABColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdCircleCColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdBackgroundColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdCircleBColor)
        Me.pnlColorButtons.Controls.Add(Me.cmdCircleAColor)
        Me.pnlColorButtons.Location = New System.Drawing.Point(296, 40)
        Me.pnlColorButtons.Name = "pnlColorButtons"
        Me.pnlColorButtons.Size = New System.Drawing.Size(168, 96)
        Me.pnlColorButtons.TabIndex = 31
        '
        'cmdOverlapABCColor
        '
        Me.cmdOverlapABCColor.Location = New System.Drawing.Point(88, 72)
        Me.cmdOverlapABCColor.Name = "cmdOverlapABCColor"
        Me.cmdOverlapABCColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdOverlapABCColor.TabIndex = 38
        Me.cmdOverlapABCColor.Text = "ABC Overlap"
        Me.cmdOverlapABCColor.Visible = False
        '
        'cmdOverlapACColor
        '
        Me.cmdOverlapACColor.Location = New System.Drawing.Point(88, 48)
        Me.cmdOverlapACColor.Name = "cmdOverlapACColor"
        Me.cmdOverlapACColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdOverlapACColor.TabIndex = 37
        Me.cmdOverlapACColor.Text = "AC Overlap"
        Me.cmdOverlapACColor.Visible = False
        '
        'cmdOverlapBCColor
        '
        Me.cmdOverlapBCColor.Location = New System.Drawing.Point(88, 24)
        Me.cmdOverlapBCColor.Name = "cmdOverlapBCColor"
        Me.cmdOverlapBCColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdOverlapBCColor.TabIndex = 36
        Me.cmdOverlapBCColor.Text = "BC Overlap"
        Me.cmdOverlapBCColor.Visible = False
        '
        'cmdOverlapABColor
        '
        Me.cmdOverlapABColor.Location = New System.Drawing.Point(88, 0)
        Me.cmdOverlapABColor.Name = "cmdOverlapABColor"
        Me.cmdOverlapABColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdOverlapABColor.TabIndex = 35
        Me.cmdOverlapABColor.Text = "AB Overlap"
        '
        'cmdCircleCColor
        '
        Me.cmdCircleCColor.Location = New System.Drawing.Point(0, 48)
        Me.cmdCircleCColor.Name = "cmdCircleCColor"
        Me.cmdCircleCColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdCircleCColor.TabIndex = 33
        Me.cmdCircleCColor.Text = "Circle C"
        '
        'cmdBackgroundColor
        '
        Me.cmdBackgroundColor.Location = New System.Drawing.Point(0, 72)
        Me.cmdBackgroundColor.Name = "cmdBackgroundColor"
        Me.cmdBackgroundColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdBackgroundColor.TabIndex = 34
        Me.cmdBackgroundColor.Text = "Background"
        '
        'cmdCircleBColor
        '
        Me.cmdCircleBColor.Location = New System.Drawing.Point(0, 24)
        Me.cmdCircleBColor.Name = "cmdCircleBColor"
        Me.cmdCircleBColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdCircleBColor.TabIndex = 32
        Me.cmdCircleBColor.Text = "Circle B"
        '
        'cmdCircleAColor
        '
        Me.cmdCircleAColor.Location = New System.Drawing.Point(0, 0)
        Me.cmdCircleAColor.Name = "cmdCircleAColor"
        Me.cmdCircleAColor.Size = New System.Drawing.Size(80, 20)
        Me.cmdCircleAColor.TabIndex = 31
        Me.cmdCircleAColor.Text = "Circle A"
        '
        'lblOverlapAC
        '
        Me.lblOverlapAC.Location = New System.Drawing.Point(240, 96)
        Me.lblOverlapAC.Name = "lblOverlapAC"
        Me.lblOverlapAC.Size = New System.Drawing.Size(40, 16)
        Me.lblOverlapAC.TabIndex = 25
        Me.lblOverlapAC.Text = "A/C"
        Me.lblOverlapAC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOverlapBC
        '
        Me.lblOverlapBC.Location = New System.Drawing.Point(176, 96)
        Me.lblOverlapBC.Name = "lblOverlapBC"
        Me.lblOverlapBC.Size = New System.Drawing.Size(40, 16)
        Me.lblOverlapBC.TabIndex = 24
        Me.lblOverlapBC.Text = "B/C"
        Me.lblOverlapBC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOverlapAB
        '
        Me.lblOverlapAB.Location = New System.Drawing.Point(112, 96)
        Me.lblOverlapAB.Name = "lblOverlapAB"
        Me.lblOverlapAB.Size = New System.Drawing.Size(40, 16)
        Me.lblOverlapAB.TabIndex = 23
        Me.lblOverlapAB.Text = "A/B"
        Me.lblOverlapAB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkCircleC
        '
        Me.chkCircleC.Location = New System.Drawing.Point(200, 12)
        Me.chkCircleC.Name = "chkCircleC"
        Me.chkCircleC.Size = New System.Drawing.Size(72, 24)
        Me.chkCircleC.TabIndex = 2
        Me.chkCircleC.Text = "Circle &C"
        '
        'txtSizeOverlapAC
        '
        Me.txtSizeOverlapAC.Location = New System.Drawing.Point(232, 112)
        Me.txtSizeOverlapAC.Name = "txtSizeOverlapAC"
        Me.txtSizeOverlapAC.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeOverlapAC.TabIndex = 14
        Me.txtSizeOverlapAC.Text = "10"
        '
        'txtSizeOverlapBC
        '
        Me.txtSizeOverlapBC.Location = New System.Drawing.Point(168, 112)
        Me.txtSizeOverlapBC.Name = "txtSizeOverlapBC"
        Me.txtSizeOverlapBC.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeOverlapBC.TabIndex = 13
        Me.txtSizeOverlapBC.Text = "10"
        '
        'txtSizeC
        '
        Me.txtSizeC.Location = New System.Drawing.Point(200, 40)
        Me.txtSizeC.Name = "txtSizeC"
        Me.txtSizeC.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeC.TabIndex = 6
        Me.txtSizeC.Text = "20"
        '
        'txtDistinctA
        '
        Me.txtDistinctA.Enabled = False
        Me.txtDistinctA.Location = New System.Drawing.Point(72, 72)
        Me.txtDistinctA.Name = "txtDistinctA"
        Me.txtDistinctA.Size = New System.Drawing.Size(56, 20)
        Me.txtDistinctA.TabIndex = 8
        Me.txtDistinctA.Text = "25"
        '
        'txtDistinctB
        '
        Me.txtDistinctB.Enabled = False
        Me.txtDistinctB.Location = New System.Drawing.Point(136, 72)
        Me.txtDistinctB.Name = "txtDistinctB"
        Me.txtDistinctB.Size = New System.Drawing.Size(56, 20)
        Me.txtDistinctB.TabIndex = 9
        Me.txtDistinctB.Text = "5"
        '
        'lblCircleB
        '
        Me.lblCircleB.Location = New System.Drawing.Point(136, 16)
        Me.lblCircleB.Name = "lblCircleB"
        Me.lblCircleB.Size = New System.Drawing.Size(48, 16)
        Me.lblCircleB.TabIndex = 1
        Me.lblCircleB.Text = "Circle &B:"
        Me.lblCircleB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOverlap
        '
        Me.lblOverlap.Location = New System.Drawing.Point(48, 112)
        Me.lblOverlap.Name = "lblOverlap"
        Me.lblOverlap.Size = New System.Drawing.Size(48, 16)
        Me.lblOverlap.TabIndex = 11
        Me.lblOverlap.Text = "&Overlap:"
        Me.lblOverlap.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCircleA
        '
        Me.lblCircleA.Location = New System.Drawing.Point(72, 16)
        Me.lblCircleA.Name = "lblCircleA"
        Me.lblCircleA.Size = New System.Drawing.Size(48, 16)
        Me.lblCircleA.TabIndex = 0
        Me.lblCircleA.Text = "Circle &A:"
        Me.lblCircleA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtSizeOverlapAB
        '
        Me.txtSizeOverlapAB.Location = New System.Drawing.Point(104, 112)
        Me.txtSizeOverlapAB.Name = "txtSizeOverlapAB"
        Me.txtSizeOverlapAB.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeOverlapAB.TabIndex = 12
        Me.txtSizeOverlapAB.Text = "18"
        '
        'txtSizeB
        '
        Me.txtSizeB.Location = New System.Drawing.Point(136, 40)
        Me.txtSizeB.Name = "txtSizeB"
        Me.txtSizeB.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeB.TabIndex = 5
        Me.txtSizeB.Text = "25"
        '
        'txtSizeA
        '
        Me.txtSizeA.Location = New System.Drawing.Point(72, 40)
        Me.txtSizeA.Name = "txtSizeA"
        Me.txtSizeA.Size = New System.Drawing.Size(56, 20)
        Me.txtSizeA.TabIndex = 4
        Me.txtSizeA.Text = "45"
        '
        'optCountDistinct
        '
        Me.optCountDistinct.Location = New System.Drawing.Point(8, 64)
        Me.optCountDistinct.Name = "optCountDistinct"
        Me.optCountDistinct.Size = New System.Drawing.Size(64, 32)
        Me.optCountDistinct.TabIndex = 7
        Me.optCountDistinct.Text = "Count Distinct"
        '
        'optTotal
        '
        Me.optTotal.Checked = True
        Me.optTotal.Location = New System.Drawing.Point(8, 40)
        Me.optTotal.Name = "optTotal"
        Me.optTotal.Size = New System.Drawing.Size(48, 16)
        Me.optTotal.TabIndex = 3
        Me.optTotal.TabStop = True
        Me.optTotal.Text = "Total"
        '
        'fraTrasks
        '
        Me.fraTrasks.Controls.Add(Me.chkFillCirclesWithColor)
        Me.fraTrasks.Controls.Add(Me.cmdCopyToClipboard)
        Me.fraTrasks.Controls.Add(Me.cmdRefresh)
        Me.fraTrasks.Controls.Add(Me.cmdSaveToDisk)
        Me.fraTrasks.Location = New System.Drawing.Point(488, 8)
        Me.fraTrasks.Name = "fraTrasks"
        Me.fraTrasks.Size = New System.Drawing.Size(106, 136)
        Me.fraTrasks.TabIndex = 1
        Me.fraTrasks.TabStop = False
        Me.fraTrasks.Text = "Tasks"
        '
        'chkFillCirclesWithColor
        '
        Me.chkFillCirclesWithColor.Checked = True
        Me.chkFillCirclesWithColor.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFillCirclesWithColor.Location = New System.Drawing.Point(8, 112)
        Me.chkFillCirclesWithColor.Name = "chkFillCirclesWithColor"
        Me.chkFillCirclesWithColor.Size = New System.Drawing.Size(88, 16)
        Me.chkFillCirclesWithColor.TabIndex = 27
        Me.chkFillCirclesWithColor.Text = "Use fill color"
        '
        'cmdCopyToClipboard
        '
        Me.cmdCopyToClipboard.Location = New System.Drawing.Point(8, 16)
        Me.cmdCopyToClipboard.Name = "cmdCopyToClipboard"
        Me.cmdCopyToClipboard.Size = New System.Drawing.Size(88, 24)
        Me.cmdCopyToClipboard.TabIndex = 0
        Me.cmdCopyToClipboard.Text = "Clipboar&d"
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(8, 80)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(88, 24)
        Me.cmdRefresh.TabIndex = 2
        Me.cmdRefresh.Text = "&Refresh"
        '
        'cmdSaveToDisk
        '
        Me.cmdSaveToDisk.Location = New System.Drawing.Point(8, 48)
        Me.cmdSaveToDisk.Name = "cmdSaveToDisk"
        Me.cmdSaveToDisk.Size = New System.Drawing.Size(88, 24)
        Me.cmdSaveToDisk.TabIndex = 1
        Me.cmdSaveToDisk.Text = "&Save file"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuHelp})
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
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditCopy, Me.mnuEditRefreshPlot, Me.mnuEditSep1, Me.mnuEditResetValues})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Index = 0
        Me.mnuEditCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuEditCopy.Text = "&Copy Plot to Clipboard"
        '
        'mnuEditRefreshPlot
        '
        Me.mnuEditRefreshPlot.Index = 1
        Me.mnuEditRefreshPlot.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuEditRefreshPlot.Text = "&Refresh Plot"
        '
        'mnuEditSep1
        '
        Me.mnuEditSep1.Index = 2
        Me.mnuEditSep1.Text = "-"
        '
        'mnuEditResetValues
        '
        Me.mnuEditResetValues.Index = 3
        Me.mnuEditResetValues.Text = "Reset Values"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 2
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Index = 0
        Me.mnuHelpAbout.Text = "&About"
        '
        'fraThreeCircleRegionCounts
        '
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.cmdUpdateRegionCounts)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionUniqueItemCountAllCircles)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.lblRegionUniqueItemCountAllCircles)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.cmdRegionCountComputeOptimal)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionCountValue)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.cboRegionCountMode)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label6)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionACx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label5)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionBCx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label4)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionABx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label3)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label2)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.Label1)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.lblRegionAx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionABC)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionCx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionAx)
        Me.fraThreeCircleRegionCounts.Controls.Add(Me.txtRegionBx)
        Me.fraThreeCircleRegionCounts.Location = New System.Drawing.Point(8, 152)
        Me.fraThreeCircleRegionCounts.Name = "fraThreeCircleRegionCounts"
        Me.fraThreeCircleRegionCounts.Size = New System.Drawing.Size(304, 232)
        Me.fraThreeCircleRegionCounts.TabIndex = 3
        Me.fraThreeCircleRegionCounts.TabStop = False
        Me.fraThreeCircleRegionCounts.Text = "Region counts for Three Circle Venn Diagram"
        '
        'cmdUpdateRegionCounts
        '
        Me.cmdUpdateRegionCounts.Location = New System.Drawing.Point(184, 136)
        Me.cmdUpdateRegionCounts.Name = "cmdUpdateRegionCounts"
        Me.cmdUpdateRegionCounts.Size = New System.Drawing.Size(112, 24)
        Me.cmdUpdateRegionCounts.TabIndex = 30
        Me.cmdUpdateRegionCounts.Text = "&Update Counts"
        '
        'txtRegionUniqueItemCountAllCircles
        '
        Me.txtRegionUniqueItemCountAllCircles.Location = New System.Drawing.Point(112, 160)
        Me.txtRegionUniqueItemCountAllCircles.Name = "txtRegionUniqueItemCountAllCircles"
        Me.txtRegionUniqueItemCountAllCircles.ReadOnly = True
        Me.txtRegionUniqueItemCountAllCircles.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionUniqueItemCountAllCircles.TabIndex = 28
        Me.txtRegionUniqueItemCountAllCircles.Text = "5"
        Me.txtRegionUniqueItemCountAllCircles.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblRegionUniqueItemCountAllCircles
        '
        Me.lblRegionUniqueItemCountAllCircles.Location = New System.Drawing.Point(24, 152)
        Me.lblRegionUniqueItemCountAllCircles.Name = "lblRegionUniqueItemCountAllCircles"
        Me.lblRegionUniqueItemCountAllCircles.Size = New System.Drawing.Size(80, 40)
        Me.lblRegionUniqueItemCountAllCircles.TabIndex = 29
        Me.lblRegionUniqueItemCountAllCircles.Text = "Unique item count across all circles"
        Me.lblRegionUniqueItemCountAllCircles.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdRegionCountComputeOptimal
        '
        Me.cmdRegionCountComputeOptimal.Location = New System.Drawing.Point(184, 168)
        Me.cmdRegionCountComputeOptimal.Name = "cmdRegionCountComputeOptimal"
        Me.cmdRegionCountComputeOptimal.Size = New System.Drawing.Size(112, 24)
        Me.cmdRegionCountComputeOptimal.TabIndex = 27
        Me.cmdRegionCountComputeOptimal.Text = "Compute &Optimal"
        '
        'txtRegionCountValue
        '
        Me.txtRegionCountValue.Location = New System.Drawing.Point(240, 200)
        Me.txtRegionCountValue.Name = "txtRegionCountValue"
        Me.txtRegionCountValue.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionCountValue.TabIndex = 26
        Me.txtRegionCountValue.Text = "18"
        '
        'cboRegionCountMode
        '
        Me.cboRegionCountMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRegionCountMode.Location = New System.Drawing.Point(16, 200)
        Me.cboRegionCountMode.Name = "cboRegionCountMode"
        Me.cboRegionCountMode.Size = New System.Drawing.Size(216, 21)
        Me.cboRegionCountMode.TabIndex = 25
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "In A and C"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRegionACx
        '
        Me.txtRegionACx.Location = New System.Drawing.Point(32, 96)
        Me.txtRegionACx.Name = "txtRegionACx"
        Me.txtRegionACx.ReadOnly = True
        Me.txtRegionACx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionACx.TabIndex = 23
        Me.txtRegionACx.Text = "25"
        Me.txtRegionACx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(184, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "In B and C"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRegionBCx
        '
        Me.txtRegionBCx.Location = New System.Drawing.Point(192, 96)
        Me.txtRegionBCx.Name = "txtRegionBCx"
        Me.txtRegionBCx.ReadOnly = True
        Me.txtRegionBCx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionBCx.TabIndex = 21
        Me.txtRegionBCx.Text = "25"
        Me.txtRegionBCx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(104, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "In A and B"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRegionABx
        '
        Me.txtRegionABx.Location = New System.Drawing.Point(112, 40)
        Me.txtRegionABx.Name = "txtRegionABx"
        Me.txtRegionABx.ReadOnly = True
        Me.txtRegionABx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionABx.TabIndex = 19
        Me.txtRegionABx.Text = "25"
        Me.txtRegionABx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(112, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "In All 3"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(112, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Only in C"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(208, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Only in B"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblRegionAx
        '
        Me.lblRegionAx.Location = New System.Drawing.Point(24, 16)
        Me.lblRegionAx.Name = "lblRegionAx"
        Me.lblRegionAx.Size = New System.Drawing.Size(56, 16)
        Me.lblRegionAx.TabIndex = 15
        Me.lblRegionAx.Text = "Only in A"
        Me.lblRegionAx.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRegionABC
        '
        Me.txtRegionABC.Location = New System.Drawing.Point(112, 80)
        Me.txtRegionABC.Name = "txtRegionABC"
        Me.txtRegionABC.ReadOnly = True
        Me.txtRegionABC.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionABC.TabIndex = 14
        Me.txtRegionABC.Text = "25"
        Me.txtRegionABC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRegionCx
        '
        Me.txtRegionCx.Location = New System.Drawing.Point(112, 128)
        Me.txtRegionCx.Name = "txtRegionCx"
        Me.txtRegionCx.ReadOnly = True
        Me.txtRegionCx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionCx.TabIndex = 13
        Me.txtRegionCx.Text = "5"
        Me.txtRegionCx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRegionAx
        '
        Me.txtRegionAx.Location = New System.Drawing.Point(24, 32)
        Me.txtRegionAx.Name = "txtRegionAx"
        Me.txtRegionAx.ReadOnly = True
        Me.txtRegionAx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionAx.TabIndex = 11
        Me.txtRegionAx.Text = "25"
        Me.txtRegionAx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRegionBx
        '
        Me.txtRegionBx.Location = New System.Drawing.Point(208, 32)
        Me.txtRegionBx.Name = "txtRegionBx"
        Me.txtRegionBx.ReadOnly = True
        Me.txtRegionBx.Size = New System.Drawing.Size(56, 20)
        Me.txtRegionBx.TabIndex = 12
        Me.txtRegionBx.Text = "5"
        Me.txtRegionBx.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtStatus
        '
        Me.txtStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStatus.Location = New System.Drawing.Point(0, 592)
        Me.txtStatus.Multiline = True
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(808, 10)
        Me.txtStatus.TabIndex = 7
        Me.txtStatus.Text = ""
        '
        'fraMessageDisplayOptions
        '
        Me.fraMessageDisplayOptions.Controls.Add(Me.chkHideMessagesOnSuccessfulUpdate)
        Me.fraMessageDisplayOptions.Controls.Add(Me.lblMessageDuplicateIgnoreWindow)
        Me.fraMessageDisplayOptions.Controls.Add(Me.lblMessageDisplayTimeSeconds)
        Me.fraMessageDisplayOptions.Controls.Add(Me.lblMessageDuplicateIgnoreWindowCaption)
        Me.fraMessageDisplayOptions.Controls.Add(Me.lblMessageDisplayTimeSecondsCaption)
        Me.fraMessageDisplayOptions.Controls.Add(Me.tbarDuplicateMessageIgnoreWindowSeconds)
        Me.fraMessageDisplayOptions.Controls.Add(Me.tbarMessageDisplayTimeSeconds)
        Me.fraMessageDisplayOptions.Location = New System.Drawing.Point(608, 8)
        Me.fraMessageDisplayOptions.Name = "fraMessageDisplayOptions"
        Me.fraMessageDisplayOptions.Size = New System.Drawing.Size(192, 128)
        Me.fraMessageDisplayOptions.TabIndex = 2
        Me.fraMessageDisplayOptions.TabStop = False
        Me.fraMessageDisplayOptions.Text = "Message Options"
        '
        'chkHideMessagesOnSuccessfulUpdate
        '
        Me.chkHideMessagesOnSuccessfulUpdate.Location = New System.Drawing.Point(8, 96)
        Me.chkHideMessagesOnSuccessfulUpdate.Name = "chkHideMessagesOnSuccessfulUpdate"
        Me.chkHideMessagesOnSuccessfulUpdate.Size = New System.Drawing.Size(176, 24)
        Me.chkHideMessagesOnSuccessfulUpdate.TabIndex = 6
        Me.chkHideMessagesOnSuccessfulUpdate.Text = "Hide old msgs on valid update"
        '
        'lblMessageDuplicateIgnoreWindow
        '
        Me.lblMessageDuplicateIgnoreWindow.Location = New System.Drawing.Point(128, 64)
        Me.lblMessageDuplicateIgnoreWindow.Name = "lblMessageDuplicateIgnoreWindow"
        Me.lblMessageDuplicateIgnoreWindow.Size = New System.Drawing.Size(48, 16)
        Me.lblMessageDuplicateIgnoreWindow.TabIndex = 5
        Me.lblMessageDuplicateIgnoreWindow.Text = "4 sec."
        '
        'lblMessageDisplayTimeSeconds
        '
        Me.lblMessageDisplayTimeSeconds.Location = New System.Drawing.Point(128, 32)
        Me.lblMessageDisplayTimeSeconds.Name = "lblMessageDisplayTimeSeconds"
        Me.lblMessageDisplayTimeSeconds.Size = New System.Drawing.Size(48, 16)
        Me.lblMessageDisplayTimeSeconds.TabIndex = 2
        Me.lblMessageDisplayTimeSeconds.Text = "8 sec."
        '
        'lblMessageDuplicateIgnoreWindowCaption
        '
        Me.lblMessageDuplicateIgnoreWindowCaption.Location = New System.Drawing.Point(8, 56)
        Me.lblMessageDuplicateIgnoreWindowCaption.Name = "lblMessageDuplicateIgnoreWindowCaption"
        Me.lblMessageDuplicateIgnoreWindowCaption.Size = New System.Drawing.Size(48, 24)
        Me.lblMessageDuplicateIgnoreWindowCaption.TabIndex = 3
        Me.lblMessageDuplicateIgnoreWindowCaption.Text = "Ignore window"
        '
        'lblMessageDisplayTimeSecondsCaption
        '
        Me.lblMessageDisplayTimeSecondsCaption.Location = New System.Drawing.Point(8, 24)
        Me.lblMessageDisplayTimeSecondsCaption.Name = "lblMessageDisplayTimeSecondsCaption"
        Me.lblMessageDisplayTimeSecondsCaption.Size = New System.Drawing.Size(48, 24)
        Me.lblMessageDisplayTimeSecondsCaption.TabIndex = 0
        Me.lblMessageDisplayTimeSecondsCaption.Text = "Display Time"
        '
        'tbarDuplicateMessageIgnoreWindowSeconds
        '
        Me.tbarDuplicateMessageIgnoreWindowSeconds.Location = New System.Drawing.Point(56, 56)
        Me.tbarDuplicateMessageIgnoreWindowSeconds.Maximum = 40
        Me.tbarDuplicateMessageIgnoreWindowSeconds.Name = "tbarDuplicateMessageIgnoreWindowSeconds"
        Me.tbarDuplicateMessageIgnoreWindowSeconds.Size = New System.Drawing.Size(72, 45)
        Me.tbarDuplicateMessageIgnoreWindowSeconds.TabIndex = 4
        Me.tbarDuplicateMessageIgnoreWindowSeconds.TickFrequency = 4
        Me.tbarDuplicateMessageIgnoreWindowSeconds.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        Me.tbarDuplicateMessageIgnoreWindowSeconds.Value = 4
        '
        'tbarMessageDisplayTimeSeconds
        '
        Me.tbarMessageDisplayTimeSeconds.Location = New System.Drawing.Point(56, 24)
        Me.tbarMessageDisplayTimeSeconds.Maximum = 40
        Me.tbarMessageDisplayTimeSeconds.Name = "tbarMessageDisplayTimeSeconds"
        Me.tbarMessageDisplayTimeSeconds.Size = New System.Drawing.Size(72, 45)
        Me.tbarMessageDisplayTimeSeconds.TabIndex = 1
        Me.tbarMessageDisplayTimeSeconds.TickFrequency = 4
        Me.tbarMessageDisplayTimeSeconds.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        Me.tbarMessageDisplayTimeSeconds.Value = 8
        '
        'fraImageAdjustmentControls
        '
        Me.fraImageAdjustmentControls.Controls.Add(Me.cmdImgAdjustmentReset)
        Me.fraImageAdjustmentControls.Controls.Add(Me.lblImgOffsetY)
        Me.fraImageAdjustmentControls.Controls.Add(Me.tbarImgYOffset)
        Me.fraImageAdjustmentControls.Controls.Add(Me.lblImgOffsetX)
        Me.fraImageAdjustmentControls.Controls.Add(Me.lblImgZoom)
        Me.fraImageAdjustmentControls.Controls.Add(Me.tbarImgXOffset)
        Me.fraImageAdjustmentControls.Controls.Add(Me.tbarImgZoom)
        Me.fraImageAdjustmentControls.Controls.Add(Me.lblImgRotation)
        Me.fraImageAdjustmentControls.Controls.Add(Me.tbarImgRotation)
        Me.fraImageAdjustmentControls.Location = New System.Drawing.Point(8, 392)
        Me.fraImageAdjustmentControls.Name = "fraImageAdjustmentControls"
        Me.fraImageAdjustmentControls.Size = New System.Drawing.Size(272, 136)
        Me.fraImageAdjustmentControls.TabIndex = 4
        Me.fraImageAdjustmentControls.TabStop = False
        Me.fraImageAdjustmentControls.Text = "Image Adjustment"
        '
        'cmdImgAdjustmentReset
        '
        Me.cmdImgAdjustmentReset.Location = New System.Drawing.Point(128, 16)
        Me.cmdImgAdjustmentReset.Name = "cmdImgAdjustmentReset"
        Me.cmdImgAdjustmentReset.Size = New System.Drawing.Size(72, 24)
        Me.cmdImgAdjustmentReset.TabIndex = 31
        Me.cmdImgAdjustmentReset.Text = "Reset"
        '
        'lblImgOffsetY
        '
        Me.lblImgOffsetY.Location = New System.Drawing.Point(128, 64)
        Me.lblImgOffsetY.Name = "lblImgOffsetY"
        Me.lblImgOffsetY.Size = New System.Drawing.Size(88, 13)
        Me.lblImgOffsetY.TabIndex = 5
        Me.lblImgOffsetY.Text = "Y Offset: 0%"
        '
        'tbarImgYOffset
        '
        Me.tbarImgYOffset.Location = New System.Drawing.Point(224, 16)
        Me.tbarImgYOffset.Maximum = 75
        Me.tbarImgYOffset.Minimum = -75
        Me.tbarImgYOffset.Name = "tbarImgYOffset"
        Me.tbarImgYOffset.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.tbarImgYOffset.Size = New System.Drawing.Size(45, 104)
        Me.tbarImgYOffset.TabIndex = 7
        Me.tbarImgYOffset.TickFrequency = 25
        Me.tbarImgYOffset.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        '
        'lblImgOffsetX
        '
        Me.lblImgOffsetX.Location = New System.Drawing.Point(128, 48)
        Me.lblImgOffsetX.Name = "lblImgOffsetX"
        Me.lblImgOffsetX.Size = New System.Drawing.Size(88, 13)
        Me.lblImgOffsetX.TabIndex = 4
        Me.lblImgOffsetX.Text = "X Offset: 0%"
        '
        'lblImgZoom
        '
        Me.lblImgZoom.Location = New System.Drawing.Point(32, 103)
        Me.lblImgZoom.Name = "lblImgZoom"
        Me.lblImgZoom.Size = New System.Drawing.Size(80, 13)
        Me.lblImgZoom.TabIndex = 3
        Me.lblImgZoom.Text = "Zoom: 100%"
        '
        'tbarImgXOffset
        '
        Me.tbarImgXOffset.Location = New System.Drawing.Point(120, 80)
        Me.tbarImgXOffset.Maximum = 75
        Me.tbarImgXOffset.Minimum = -75
        Me.tbarImgXOffset.Name = "tbarImgXOffset"
        Me.tbarImgXOffset.Size = New System.Drawing.Size(104, 45)
        Me.tbarImgXOffset.TabIndex = 6
        Me.tbarImgXOffset.TickFrequency = 25
        Me.tbarImgXOffset.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        '
        'tbarImgZoom
        '
        Me.tbarImgZoom.Location = New System.Drawing.Point(8, 72)
        Me.tbarImgZoom.Maximum = 200
        Me.tbarImgZoom.Minimum = 5
        Me.tbarImgZoom.Name = "tbarImgZoom"
        Me.tbarImgZoom.Size = New System.Drawing.Size(104, 45)
        Me.tbarImgZoom.TabIndex = 2
        Me.tbarImgZoom.TickFrequency = 25
        Me.tbarImgZoom.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        Me.tbarImgZoom.Value = 100
        '
        'lblImgRotation
        '
        Me.lblImgRotation.Location = New System.Drawing.Point(32, 47)
        Me.lblImgRotation.Name = "lblImgRotation"
        Me.lblImgRotation.Size = New System.Drawing.Size(80, 13)
        Me.lblImgRotation.TabIndex = 1
        Me.lblImgRotation.Text = "Rotation: 0"
        '
        'tbarImgRotation
        '
        Me.tbarImgRotation.LargeChange = 10
        Me.tbarImgRotation.Location = New System.Drawing.Point(8, 16)
        Me.tbarImgRotation.Maximum = 360
        Me.tbarImgRotation.Name = "tbarImgRotation"
        Me.tbarImgRotation.Size = New System.Drawing.Size(104, 45)
        Me.tbarImgRotation.TabIndex = 0
        Me.tbarImgRotation.TickFrequency = 45
        Me.tbarImgRotation.TickStyle = System.Windows.Forms.TickStyle.TopLeft
        '
        'vdgThreeCircles
        '
        Me.vdgThreeCircles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vdgThreeCircles.BackColor = System.Drawing.SystemColors.ControlDark
        Me.vdgThreeCircles.Location = New System.Drawing.Point(320, 152)
        Me.vdgThreeCircles.Name = "vdgThreeCircles"
        Me.vdgThreeCircles.Size = New System.Drawing.Size(480, 432)
        Me.vdgThreeCircles.TabIndex = 9
        '
        'vdgTwoCircles
        '
        Me.vdgTwoCircles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vdgTwoCircles.Location = New System.Drawing.Point(8, 160)
        Me.vdgTwoCircles.Name = "vdgTwoCircles"
        Me.vdgTwoCircles.Size = New System.Drawing.Size(792, 424)
        Me.vdgTwoCircles.TabIndex = 10
        '
        'DisplayForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(808, 609)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.fraImageAdjustmentControls)
        Me.Controls.Add(Me.fraThreeCircleRegionCounts)
        Me.Controls.Add(Me.vdgThreeCircles)
        Me.Controls.Add(Me.vdgTwoCircles)
        Me.Controls.Add(Me.fraMessageDisplayOptions)
        Me.Controls.Add(Me.fraTrasks)
        Me.Controls.Add(Me.fraParameters)
        Me.Menu = Me.MainMenu1
        Me.Name = "DisplayForm"
        Me.Text = "Venn Diagram Plotter"
        Me.fraParameters.ResumeLayout(False)
        Me.pnlColorButtons.ResumeLayout(False)
        Me.fraTrasks.ResumeLayout(False)
        Me.fraThreeCircleRegionCounts.ResumeLayout(False)
        Me.fraMessageDisplayOptions.ResumeLayout(False)
        CType(Me.tbarDuplicateMessageIgnoreWindowSeconds, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tbarMessageDisplayTimeSeconds, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraImageAdjustmentControls.ResumeLayout(False)
        CType(Me.tbarImgYOffset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tbarImgXOffset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tbarImgZoom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tbarImgRotation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constants"

    Private Const XML_SECTION_OPTIONS As String = "CircleParams"

    Private Const STATUS_HEIGHT_NO_MESSAGE As Integer = 10
    Private Const STATUS_HEIGHT_SINGLE_LINE As Integer = 20
    Private Const STATUS_HEIGHT_PER_LINE As Integer = 12

    Private Const DEFAULT_SECONDS_TO_DISPLAY_EACH_MESSAGE As Integer = 8
    Private Const DEFAULT_DUPLICATE_IGNORE_WINDOW_SECONDS As Integer = 4

    Private Const DEFAULT_DATA_CIRCLE_A As Integer = 45
    Private Const DEFAULT_DATA_CIRCLE_B As Integer = 25
    Private Const DEFAULT_DATA_CIRCLE_C As Integer = 20
    Private Const DEFAULT_DATA_OVERLAP_AB As Integer = 18
    Private Const DEFAULT_DATA_OVERLAP_BC As Integer = 10
    Private Const DEFAULT_DATA_OVERLAP_AC As Integer = 15

    Private Const MINIMUM_WINDOW_HEIGHT_TWO_CIRCLE As Integer = 450

    Private Const DEFAULT_WINDOW_HEIGHT_TWO_CIRCLE As Integer = 610
    Private Const DEFAULT_WINDOW_HEIGHT_THREE_CIRCLE As Integer = 610

    Private Const DEFAULT_WINDOW_WIDTH_TWO_CIRCLE As Integer = 637
    Private Const DEFAULT_WINDOW_WIDTH_THREE_CIRCLE As Integer = 820

    Private Const PROGRAM_DATE As String = "September 19, 2007"
#End Region

#Region "Structures and Enums"
    Enum eRegionCountModeConstants
        CountOverlappingAllCircles = 0
        UniqueItemCountAcrossAllCircles = 1
    End Enum

    Protected Structure udtMessageQueueType
        Public Message As String
        Public DisplayTime As DateTime
        Public MessageNumber As Integer
    End Structure

    Protected Structure udtCircleDimensionsType
        Public CircleA As Double
        Public CircleB As Double
        Public CircleC As Double
        Public OverlapAB As Double
        Public OverlapBC As Double
        Public OverlapAC As Double
    End Structure

#End Region

#Region "Module Variables"

    Protected mCircleDimensionsSaved As udtCircleDimensionsType
    Protected mIniFileName As String = "VennDiagramPlotter.xml"

    Protected mMessageQueueCount As Integer
    Protected mMessageQueue() As udtMessageQueueType
    Protected mMessageQueueGlobalCount As Integer

    Protected WithEvents mStatusTimer As Timer
#End Region

#Region "Properties"
    Protected Property MessageDisplayTime() As Integer
        Get
            Return tbarMessageDisplayTimeSeconds.Value
        End Get
        Set(ByVal Value As Integer)
            SetTrackbarValue(tbarMessageDisplayTimeSeconds, Value)
        End Set
    End Property

    Protected Property DuplicateMessageIgnoreWindow() As Integer
        Get
            Return tbarDuplicateMessageIgnoreWindowSeconds.Value
        End Get
        Set(ByVal Value As Integer)
            SetTrackbarValue(tbarDuplicateMessageIgnoreWindowSeconds, Value)
        End Set
    End Property

#End Region

    ''' Unused function
    ''Private Sub AutoResizeExpand()
    ''    vdgTwoCircles.VennDiagram.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
    ''    vdgTwoCircles.VennDiagram.Height = CInt(vdgTwoCircles.VennDiagram.Height + vdgTwoCircles.VennDiagram.Height * 0.1)
    ''    vdgTwoCircles.Height = CInt(vdgTwoCircles.Height + vdgTwoCircles.Height * 0.1)
    ''    vdgTwoCircles.VennDiagram.Width = CInt(vdgTwoCircles.VennDiagram.Width + vdgTwoCircles.VennDiagram.Width * 0.1)
    ''    vdgTwoCircles.Width = CInt(vdgTwoCircles.Width + vdgTwoCircles.Width * 0.1)
    ''End Sub

    ''' Unused function
    ''Private Sub AutoResizeShrink()
    ''    vdgTwoCircles.VennDiagram.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
    ''    vdgTwoCircles.VennDiagram.Height = CInt(vdgTwoCircles.VennDiagram.Height - vdgTwoCircles.VennDiagram.Height * 0.1)
    ''    vdgTwoCircles.Height = CInt(vdgTwoCircles.Height - vdgTwoCircles.Height * 0.1)
    ''    vdgTwoCircles.VennDiagram.Width = CInt(vdgTwoCircles.VennDiagram.Width - vdgTwoCircles.VennDiagram.Width * 0.1)
    ''    vdgTwoCircles.Width = CInt(vdgTwoCircles.Width - vdgTwoCircles.Width * 0.1)
    ''End Sub

    Private Sub AppendToStatusLog(ByVal strMessage As String)
        Dim intIndex As Integer
        Dim blnSkipMessage As Boolean = False

        If mMessageQueue Is Nothing Then
            ReDim mMessageQueue(9)
        End If

        If mMessageQueueCount >= mMessageQueue.Length Then
            ' Reserve more space in mMessageQueue
            ReDim Preserve mMessageQueue(mMessageQueue.Length * 2 - 1)
        End If

        ' See if the queue already contains strMessage, posted within the last second
        ' If it does; then don't re-add it
        For intIndex = 0 To mMessageQueueCount - 1
            If mMessageQueue(intIndex).Message = strMessage Then
                If System.DateTime.Now.Subtract(mMessageQueue(intIndex).DisplayTime).TotalSeconds <= Me.DuplicateMessageIgnoreWindow Then
                    blnSkipMessage = True
                    Exit For
                End If
            End If
        Next intIndex

        If Not blnSkipMessage Then
            With mMessageQueue(mMessageQueueCount)
                .DisplayTime = System.DateTime.Now
                .Message = String.Copy(strMessage)
                mMessageQueueGlobalCount += 1
                .MessageNumber = mMessageQueueGlobalCount
            End With
            mMessageQueueCount += 1

            DisplayStatusLog()
        End If

    End Sub

    Private Sub AutoSizeWindow()
        If chkCircleC.Checked Then
            Me.Width = DEFAULT_WINDOW_WIDTH_THREE_CIRCLE
            Me.Height = DEFAULT_WINDOW_HEIGHT_THREE_CIRCLE
        Else
            Me.Width = DEFAULT_WINDOW_WIDTH_TWO_CIRCLE
            Me.Height = DEFAULT_WINDOW_HEIGHT_TWO_CIRCLE
        End If
    End Sub

    Private Sub AutoUpdateRegionCountValue()

        Dim eRegionCountMode As eRegionCountModeConstants
        eRegionCountMode = GetRegionCountMode()

        Select Case eRegionCountMode
            Case eRegionCountModeConstants.CountOverlappingAllCircles
                txtRegionCountValue.Text = txtRegionABC.Text
            Case eRegionCountModeConstants.UniqueItemCountAcrossAllCircles
                txtRegionCountValue.Text = txtRegionUniqueItemCountAllCircles.Text
            Case Else
                ' Unknown value
                UpdateStatus("Unknown Region Count Mode in Sub AutoUpdateRegionCountValue")
        End Select

    End Sub

    Private Sub ClearMessages()
        If mMessageQueueCount > 0 Then
            mMessageQueueCount = 0
            DisplayStatusLog()
        End If
    End Sub

    Private Sub CopyVennToClipboard()
        Try
            Dim bitmap As Bitmap

            If chkCircleC.Checked Then
                bitmap = New Bitmap(Me.vdgThreeCircles.Width, Me.vdgThreeCircles.Height)
            Else
                bitmap = New Bitmap(Me.vdgTwoCircles.Width, Me.vdgTwoCircles.Height)
            End If

            Dim g As Graphics = Graphics.FromImage(bitmap)

            If chkCircleC.Checked Then
                ControlPrinter.ControlPrinter.DrawControl(Me.vdgThreeCircles, g, True)
            Else
                ControlPrinter.ControlPrinter.DrawControl(Me.vdgTwoCircles, g, True)
            End If


            Clipboard.SetDataObject(bitmap)
        Catch ex As Exception
            MsgBox("Error copying Venn diagram to clipboard: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

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

    Private Sub ComputeOptimalRegionCount(ByVal blnUpdateDisplayedDistinctCounts As Boolean)
        Dim eRegionCountMode As eRegionCountModeConstants
        Dim dblRegionCountValue As Double

        Dim udtCircleRegions As VennDiagrams.ThreeCircleVennDiagram.udtThreeCircleRegionsType

        eRegionCountMode = GetRegionCountMode()
        dblRegionCountValue = 5

        If GetCircleDimensions(udtCircleRegions, True, False) Then

            Me.Cursor = Cursors.WaitCursor

            Select Case eRegionCountMode
                Case eRegionCountModeConstants.CountOverlappingAllCircles
                    With udtCircleRegions
                        ' Initially define the default value as 50% of the smallest circle size
                        dblRegionCountValue = Math.Min(Math.Min(.CircleA, .CircleB), .CircleC) / 2.0
                        If dblRegionCountValue <= 0 Then dblRegionCountValue = 1
                    End With

                    If VennDiagrams.ThreeCircleVennDiagram.ComputeThreeCircleAreasGivenOverlapABC(udtCircleRegions, 0) Then
                        ' Successfully computed the optimal value
                        ' Obtain the optimal value from .ABC
                        dblRegionCountValue = udtCircleRegions.ABC
                    Else
                        UpdateStatus("Unable to determine an optimal value for ABC overlap")
                    End If


                Case eRegionCountModeConstants.UniqueItemCountAcrossAllCircles
                    With udtCircleRegions
                        ' Initially define the default value as 50% of the sum of the three circles
                        dblRegionCountValue = (.CircleA + .CircleB + .CircleC) / 2.0
                        If dblRegionCountValue <= 0 Then dblRegionCountValue = 1
                    End With

                    If VennDiagrams.ThreeCircleVennDiagram.ComputeThreeCircleAreasGivenTotalUniqueCount(udtCircleRegions, 0) Then
                        ' Successfully computed the optimal value
                        ' Obtain the optimal value from .TotalUniqueCount
                        dblRegionCountValue = udtCircleRegions.TotalUniqueCount
                    Else
                        UpdateStatus("Unable to determine an optimal value for total unique count")
                    End If


                Case Else
                    ' Unknown value
                    UpdateStatus("Unknown Region Count Mode in Sub ComputeOptimalRegionCount")
            End Select

            Me.Cursor = Cursors.Default
        End If

        txtRegionCountValue.Text = NumToString(dblRegionCountValue, 1)

        If blnUpdateDisplayedDistinctCounts Then
            UpdateThreeCircleDistinctCounts()
        End If
    End Sub

    Private Sub DisplayDistinctRegionCount(ByVal dblValue As Double, ByVal objTextbox As TextBox, ByVal strRegionDescription As String)
        If dblValue >= 0 Then
            objTextbox.Text = NumToString(dblValue, 1)
        Else
            objTextbox.Text = String.Empty
            If Not strRegionDescription Is Nothing AndAlso strRegionDescription.Length > 0 Then
                UpdateStatus("Computed a value of " & NumToString(dblValue, 1) & " for the points " & strRegionDescription & "; use 'Compute Optimal' to auto-determine valid values.")
            End If
        End If
    End Sub

    Private Sub DisplayForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        vdgTwoCircles.VennDiagram.Anchor = vdgTwoCircles.Anchor
    End Sub

    Private Sub DisplayStatusLog()
        Const MAX_CONTROL_HEIGHT As Integer = 96
        Const BUFFER_DISTANCE As Integer = 60

        Dim strMessageList As String
        Dim intControlHeight As Integer

        Dim intIndex As Integer

        If mMessageQueueCount > 0 Then
            If mMessageQueueCount = 1 Then
                intControlHeight = STATUS_HEIGHT_SINGLE_LINE
            Else
                intControlHeight = STATUS_HEIGHT_SINGLE_LINE + mMessageQueueCount * STATUS_HEIGHT_PER_LINE
            End If

            If intControlHeight > MAX_CONTROL_HEIGHT Then intControlHeight = MAX_CONTROL_HEIGHT

            For intIndex = 0 To mMessageQueueCount - 1
                If intIndex > 0 Then strMessageList &= ControlChars.NewLine

                With mMessageQueue(intIndex)
                    strMessageList &= "(" & (.MessageNumber).ToString & ") " & .DisplayTime.ToLongTimeString & ": " & .Message
                End With
            Next intIndex
        Else
            strMessageList = String.Empty
            intControlHeight = STATUS_HEIGHT_NO_MESSAGE
        End If

        If txtStatus.Height <> intControlHeight Then
            txtStatus.Top = Me.Height - intControlHeight - BUFFER_DISTANCE
            txtStatus.Height = intControlHeight

            If txtStatus.Height = STATUS_HEIGHT_NO_MESSAGE OrElse mMessageQueueCount = 1 Then
                txtStatus.ScrollBars = ScrollBars.None
            Else
                txtStatus.ScrollBars = ScrollBars.Vertical
            End If

        End If
        txtStatus.Text = strMessageList
    End Sub

    Private Sub EnableDisableControls()

        Dim blnEnableTotals As Boolean
        Dim blnThreeCircleMode As Boolean

        blnThreeCircleMode = chkCircleC.Checked
        If blnThreeCircleMode Then
            ' Force Totals to be enabled
            blnEnableTotals = True
            If Not optTotal.Checked Then optTotal.Checked = True

            ' Make sure form is at least DEFAULT_WINDOW_HEIGHT_THREE_CIRCLE units hight
            If Me.Height < DEFAULT_WINDOW_HEIGHT_THREE_CIRCLE Then Me.Height = DEFAULT_WINDOW_HEIGHT_THREE_CIRCLE
        Else
            blnEnableTotals = optTotal.Checked
        End If

        fraThreeCircleRegionCounts.Visible = blnThreeCircleMode
        fraImageAdjustmentControls.Visible = blnThreeCircleMode

        txtSizeA.Enabled = blnEnableTotals
        txtSizeB.Enabled = blnEnableTotals

        txtSizeC.Enabled = blnThreeCircleMode And blnEnableTotals
        txtSizeC.Visible = blnThreeCircleMode

        txtDistinctA.Enabled = Not blnEnableTotals
        txtDistinctB.Enabled = Not blnEnableTotals

        optCountDistinct.Enabled = Not blnThreeCircleMode

        txtSizeOverlapBC.Visible = blnThreeCircleMode
        txtSizeOverlapAC.Visible = blnThreeCircleMode
        lblOverlapBC.Visible = blnThreeCircleMode
        lblOverlapAC.Visible = blnThreeCircleMode

        cmdCircleCColor.Visible = blnThreeCircleMode

        cmdOverlapBCColor.Visible = blnThreeCircleMode
        cmdOverlapACColor.Visible = blnThreeCircleMode
        cmdOverlapABCColor.Visible = blnThreeCircleMode

        If Not vdgTwoCircles Is Nothing Then
            vdgTwoCircles.Visible = Not blnThreeCircleMode
            vdgThreeCircles.Visible = blnThreeCircleMode

            If blnThreeCircleMode Then
                txtStatus.Left = vdgThreeCircles.Left
                txtStatus.Width = vdgThreeCircles.Width
            Else
                txtStatus.Left = vdgTwoCircles.Left
                txtStatus.Width = vdgTwoCircles.Width
            End If
        End If

        PositionControls()

    End Sub

    Private Function GetCircleDimensions(ByRef udtCircleRegions As VennDiagrams.ThreeCircleVennDiagram.udtThreeCircleRegionsType, ByVal blnIncludeThreeCircleValues As Boolean, ByVal blnWarnIfInvalid As Boolean) As Boolean
        Dim udtCircleDimensions As udtCircleDimensionsType
        Dim blnSuccess As Boolean

        blnSuccess = GetCircleDimensions(udtCircleDimensions, blnIncludeThreeCircleValues, blnWarnIfInvalid)

        ' Populate udtCircleRegions using udtCircleDimensions
        With udtCircleDimensions
            udtCircleRegions.CircleA = .CircleA
            udtCircleRegions.CircleB = .CircleB
            udtCircleRegions.CircleC = .CircleC
            udtCircleRegions.OverlapAB = .OverlapAB
            udtCircleRegions.OverlapBC = .OverlapBC
            udtCircleRegions.OverlapAC = .OverlapAC
        End With

        Return blnSuccess
    End Function

    Private Function GetCircleDimensions(ByRef udtCircleDimensions As udtCircleDimensionsType, ByVal blnIncludeThreeCircleValues As Boolean, ByVal blnWarnIfInvalid As Boolean) As Boolean
        ' Returns True if the dimensions are valid; false if not

        Dim strMessage As String
        Dim blnValid As Boolean

        blnValid = True
        strMessage = String.Empty

        With udtCircleDimensions
            .CircleA = 0
            .CircleB = 0
            .CircleC = 0
            .OverlapAB = 0
            .OverlapBC = 0
            .OverlapAC = 0

            If IsNumber(txtSizeA) Then
                .CircleA = TextBoxToDbl(txtSizeA)
            Else
                If blnWarnIfInvalid Then strMessage &= "Value not provided for Circle A total;"
                blnValid = False
            End If

            If IsNumber(txtSizeB) Then
                .CircleB = TextBoxToDbl(txtSizeB)
            Else
                If blnWarnIfInvalid Then strMessage &= "Value not provided for Circle B total;"
                blnValid = False
            End If

            If IsNumber(txtSizeOverlapAB) Then
                .OverlapAB = TextBoxToDbl(txtSizeOverlapAB)
            Else
                If blnWarnIfInvalid Then strMessage &= "Value not provided for Overlap of Circles A and B;"
                blnValid = False
            End If

            If blnIncludeThreeCircleValues Then
                If IsNumber(txtSizeC) Then
                    .CircleC = TextBoxToDbl(txtSizeC)
                Else
                    If blnWarnIfInvalid Then strMessage &= "Value not provided for Circle C total;"
                    blnValid = False
                End If

                If IsNumber(txtSizeOverlapBC) Then
                    .OverlapBC = TextBoxToDbl(txtSizeOverlapBC)
                Else
                    If blnWarnIfInvalid Then strMessage &= "Value not provided for Overlap of Circles B and C;"
                    blnValid = False
                End If

                If IsNumber(txtSizeOverlapAC) Then
                    .OverlapAC = TextBoxToDbl(txtSizeOverlapAC)
                Else
                    If blnWarnIfInvalid Then strMessage &= "Value not provided for Overlap of Circles A and C;"
                    blnValid = False
                End If
            End If
        End With

        If blnWarnIfInvalid AndAlso strMessage.Length > 0 Then
            UpdateStatus(strMessage.TrimEnd(";"c))
        End If

        Return blnValid

    End Function

    Private Function GetIniFilePath(ByVal IniFileName As String) As String
        Dim fi As New System.IO.FileInfo(Application.ExecutablePath)
        Return System.IO.Path.Combine(fi.DirectoryName, IniFileName)
    End Function

    Private Function GetRegionCountMode() As eRegionCountModeConstants
        If cboRegionCountMode.SelectedIndex = eRegionCountModeConstants.CountOverlappingAllCircles Then
            Return eRegionCountModeConstants.CountOverlappingAllCircles
        Else
            Return eRegionCountModeConstants.UniqueItemCountAcrossAllCircles
        End If
    End Function

    Private Sub InitializeControls()

        mMessageQueueCount = 0
        ReDim mMessageQueue(9)
        mMessageQueueGlobalCount = 0

        Me.MessageDisplayTime = DEFAULT_SECONDS_TO_DISPLAY_EACH_MESSAGE
        Me.DuplicateMessageIgnoreWindow = DEFAULT_DUPLICATE_IGNORE_WINDOW_SECONDS

        mStatusTimer = New Timer
        With mStatusTimer
            .Interval = 200
            .Enabled = True
        End With

        With cboRegionCountMode
            With .Items
                .Clear()
                .Add("Define count overlapping all 3 circles ")
                .Add("Define total unique item count")
            End With
            .SelectedIndex = eRegionCountModeConstants.CountOverlappingAllCircles
        End With

        ResetImageAdjustmentValues()

        With vdgTwoCircles
            With .VennDiagram
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                .BackColor = dlgColor.Color.White
            End With
            .Visible = True
        End With

        With vdgThreeCircles
            With .VennDiagram
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                .BackColor = dlgColor.Color.White
            End With
            .Visible = False
            .Left = fraThreeCircleRegionCounts.Left + fraThreeCircleRegionCounts.Width + 12
            .Top = vdgTwoCircles.Top
            .Width = vdgTwoCircles.Width - (vdgThreeCircles.Left - vdgTwoCircles.Left)
            .Height = vdgTwoCircles.Height
        End With

        SetToolTips()

        EnableDisableControls()

        RefreshVennDiagrams(True)
        LoadDefaults()

        ' Initialize the default circle dimensions
        Dim udtCircleDimensions As udtCircleDimensionsType

        If GetCircleDimensions(udtCircleDimensions, True, False) Then
            StoreCurrentDimensions(udtCircleDimensions)
        Else
            With mCircleDimensionsSaved
                .CircleA = 45
                .CircleB = 25
                .CircleC = 15
                .OverlapAB = 18
                .OverlapBC = 5
                .OverlapAC = 10
            End With

            RestorePreviousDimensions(mCircleDimensionsSaved, False)
        End If

    End Sub

    Public Shared Function IsNumber(ByVal strValue As String) As Boolean
        Dim objFormatProvider As System.Globalization.NumberFormatInfo
        Try
            Return Double.TryParse(strValue, Globalization.NumberStyles.Any, objFormatProvider, 0)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function IsNumber(ByVal objTextbox As TextBox) As Boolean
        Try
            If objTextbox.TextLength > 0 Then
                Return IsNumber(objTextbox.Text)
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub LoadDefaults()
        Dim outputFilepath As String = String.Empty

        Dim objXmlFile As New XmlSettingsFileAccessor
        Dim valueNotPresent As Boolean

        Try
            outputFilepath = GetIniFilePath(mIniFileName)
            If Not System.IO.File.Exists(outputFilepath) Then
                SaveDefaults()
            End If

            ' Pass False to .LoadSettings() here to turn off case sensitive matching
            objXmlFile.LoadSettings(outputFilepath, False)


            Try
                valueNotPresent = False
                Try
                    txtSizeA.Text = objXmlFile.GetParam(XML_SECTION_OPTIONS, "CircleADia", 45, valueNotPresent).ToString
                Catch ex As Exception
                    valueNotPresent = True
                End Try

                If valueNotPresent Then
                    txtSizeA.Text = DEFAULT_DATA_CIRCLE_A.ToString
                    txtSizeB.Text = DEFAULT_DATA_CIRCLE_B.ToString
                    txtSizeC.Text = DEFAULT_DATA_CIRCLE_C.ToString
                    txtSizeOverlapAB.Text = DEFAULT_DATA_OVERLAP_AB.ToString
                    txtSizeOverlapBC.Text = DEFAULT_DATA_OVERLAP_BC.ToString
                    txtSizeOverlapAC.Text = DEFAULT_DATA_OVERLAP_AC.ToString
                Else
                    With objXmlFile
                        txtSizeB.Text = .GetParam(XML_SECTION_OPTIONS, "CircleBDia", txtSizeB.Text).ToString
                        txtSizeC.Text = .GetParam(XML_SECTION_OPTIONS, "CircleCDia", txtSizeC.Text).ToString

                        Try
                            txtSizeOverlapAB.Text = NumToString(Math.Min(TextBoxToDbl(txtSizeA), TextBoxToDbl(txtSizeB)) / 2, 1)
                        Catch ex2 As Exception
                            txtSizeOverlapAB.Text = "0"
                        End Try
                        txtSizeOverlapAB.Text = .GetParam(XML_SECTION_OPTIONS, "Overlap", txtSizeOverlapAB.Text)

                        Try
                            txtSizeOverlapBC.Text = NumToString(Math.Min(TextBoxToDbl(txtSizeB), TextBoxToDbl(txtSizeC)) / 2, 1)
                        Catch ex2 As Exception
                            txtSizeOverlapBC.Text = "0"
                        End Try
                        txtSizeOverlapBC.Text = .GetParam(XML_SECTION_OPTIONS, "OverlapBC", txtSizeOverlapBC.Text)

                        Try
                            txtSizeOverlapAC.Text = NumToString(Math.Min(TextBoxToDbl(txtSizeA), TextBoxToDbl(txtSizeC)) / 2, 1)
                        Catch ex2 As Exception
                            txtSizeOverlapAC.Text = "0"
                        End Try
                        txtSizeOverlapAC.Text = .GetParam(XML_SECTION_OPTIONS, "OverlapAC", txtSizeOverlapAC.Text)
                    End With
                End If

                Try
                    vdgTwoCircles.VennDiagram.CircleAColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "CircleAColor", vdgTwoCircles.VennDiagram.CircleAColor)
                    vdgTwoCircles.VennDiagram.CircleBColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "CircleBColor", vdgTwoCircles.VennDiagram.CircleBColor)

                    vdgThreeCircles.VennDiagram.CircleAColor = vdgTwoCircles.VennDiagram.CircleAColor
                    vdgThreeCircles.VennDiagram.CircleBColor = vdgTwoCircles.VennDiagram.CircleBColor
                    vdgThreeCircles.VennDiagram.CircleCColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "CircleCColor", vdgThreeCircles.VennDiagram.CircleCColor)

                    vdgTwoCircles.VennDiagram.OverlapColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "OverlapColor", vdgTwoCircles.VennDiagram.OverlapColor)
                    vdgThreeCircles.VennDiagram.OverlapABColor = vdgTwoCircles.VennDiagram.OverlapColor

                    vdgThreeCircles.VennDiagram.OverlapBCColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "OverlapBCColor", vdgThreeCircles.VennDiagram.OverlapBCColor)
                    vdgThreeCircles.VennDiagram.OverlapACColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "OverlapACColor", vdgThreeCircles.VennDiagram.OverlapACColor)
                    vdgThreeCircles.VennDiagram.OverlapABCColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "OverlapABCColor", vdgThreeCircles.VennDiagram.OverlapABCColor)

                    vdgTwoCircles.VennDiagram.BackColor = LoadDefaultColorVal(objXmlFile, XML_SECTION_OPTIONS, "BackgroundColor", vdgTwoCircles.VennDiagram.BackColor)
                    vdgThreeCircles.VennDiagram.BackColor = vdgTwoCircles.VennDiagram.BackColor

                Catch ex As Exception
                    ' Don't worry about errors here
                End Try

                With objXmlFile
                    chkFillCirclesWithColor.Checked = .GetParam(XML_SECTION_OPTIONS, "FillCirclesWithColor", chkFillCirclesWithColor.Checked)

                    optTotal.Checked = .GetParam(XML_SECTION_OPTIONS, "EnterCircleTotals", optTotal.Checked)

                    chkCircleC.Checked = .GetParam(XML_SECTION_OPTIONS, "ThreeCircleMode", chkCircleC.Checked)

                    Me.MessageDisplayTime = .GetParam(XML_SECTION_OPTIONS, "MessageDisplayTime.", Me.MessageDisplayTime)
                    Me.DuplicateMessageIgnoreWindow = .GetParam(XML_SECTION_OPTIONS, "DuplicateMessageIgnoreWindow", Me.DuplicateMessageIgnoreWindow)
                    chkHideMessagesOnSuccessfulUpdate.Checked = .GetParam(XML_SECTION_OPTIONS, "HideMessagesOnSuccessfulUpdate", chkHideMessagesOnSuccessfulUpdate.Checked)

                    Try
                        cboRegionCountMode.SelectedIndex = .GetParam(XML_SECTION_OPTIONS, "RegionCountMode", cboRegionCountMode.SelectedIndex)
                    Catch ex As Exception
                        ' Ignore errors here
                    End Try
                    txtRegionCountValue.Text = .GetParam(XML_SECTION_OPTIONS, "RegionCountValue", txtRegionCountValue.Text)

                    Me.Width = .GetParam(XML_SECTION_OPTIONS, "WindowWidth", Me.Width)
                    Me.Height = .GetParam(XML_SECTION_OPTIONS, "WindowHeight", Me.Height)
                End With

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Invalid parameter in settings file: " & outputFilepath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try

            RefreshVennDiagrams(True)

        Catch ex As Exception
            MsgBox("Error loading Defaults from " & outputFilepath & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Protected Function LoadDefaultColorVal(ByVal objXmlFile As XmlSettingsFileAccessor, ByVal strSection As String, ByVal strKeyName As String, ByVal objColorIfMissing As System.Drawing.Color) As System.Drawing.Color
        Try
            Return StringToColor(objXmlFile.GetParam(strSection, strKeyName, ColorToString(objColorIfMissing)))
        Catch ex As Exception
            Return objColorIfMissing
        End Try
    End Function

    Protected Function NumToString(ByVal dblValue As Double, ByVal intDigitsToRound As Integer) As String
        Dim strValue As String
        Dim intPeriodIndex As Integer

        If intDigitsToRound < 0 Then intDigitsToRound = 0
        strValue = Math.Round(dblValue, intDigitsToRound).ToString

        If intDigitsToRound > 0 Then
            ' If the number looks like "34.230" then trim to "34.23"
            ' If the number looks like "34.0000" then trim to "34"
            intPeriodIndex = strValue.IndexOf("."c)
            If intPeriodIndex >= 0 Then
                strValue = strValue.Substring(0, intPeriodIndex) & strValue.Substring(intPeriodIndex).TrimEnd("0"c).TrimEnd("."c)
            End If
        End If

        Return strValue

    End Function

    Private Sub PositionControls()

        Const ColorButtonsPanelLeftA As Integer = 208
        Const ColorButtonsPanelLeftB As Integer = 296

        Const ColorButtonsPanelWidthA As Integer = 80
        Const ColorButtonsPanelWidthB As Integer = 168

        Const ParametersFrameWidth As Integer = 472
        Const FrameSpacing As Integer = 8

        If chkCircleC.Checked Then
            pnlColorButtons.Left = ColorButtonsPanelLeftB
            pnlColorButtons.Width = ColorButtonsPanelWidthB

            cmdOverlapABColor.Left = 88
            cmdOverlapABColor.Top = 0

            fraParameters.Width = ParametersFrameWidth
        Else
            pnlColorButtons.Left = ColorButtonsPanelLeftA
            pnlColorButtons.Width = ColorButtonsPanelWidthA

            cmdOverlapABColor.Left = 0
            cmdOverlapABColor.Top = 48

            fraParameters.Width = ParametersFrameWidth - (ColorButtonsPanelLeftB - ColorButtonsPanelLeftA) - (ColorButtonsPanelWidthB - ColorButtonsPanelWidthA)
        End If

        fraTrasks.Left = fraParameters.Left + fraParameters.Width + FrameSpacing
        fraMessageDisplayOptions.Left = fraTrasks.Left + fraTrasks.Width + FrameSpacing

        fraImageAdjustmentControls.Top = fraThreeCircleRegionCounts.Top + fraThreeCircleRegionCounts.Height + 2

    End Sub

    Private Sub RefreshVennDiagrams(ByVal blnUpdateCountDistinct As Boolean)
        RefreshVennDiagrams(blnUpdateCountDistinct, True)
    End Sub

    Private Sub RefreshVennDiagrams(ByVal blnUpdateCountDistinct As Boolean, ByVal blnWarnIfInvalid As Boolean)
        Dim udtCircleDimensions As udtCircleDimensionsType

        Try
            ' First check that each of the controls contains a number
            If GetCircleDimensions(udtCircleDimensions, chkCircleC.Checked, blnWarnIfInvalid) Then

                ' Now make sure the dimensions are reasonable
                If ValidDimensionsPresent(udtCircleDimensions, blnWarnIfInvalid) Then
                    If chkCircleC.Checked Then
                        With vdgThreeCircles.VennDiagram
                            .CircleASize = udtCircleDimensions.CircleA
                            .CircleBSize = udtCircleDimensions.CircleB
                            .CircleCSize = udtCircleDimensions.CircleC
                            .OverlapABSize = udtCircleDimensions.OverlapAB
                            .OverlapBCSize = udtCircleDimensions.OverlapBC
                            .OverlapACSize = udtCircleDimensions.OverlapAC
                        End With
                    Else
                        With vdgTwoCircles.VennDiagram
                            .CircleASize = udtCircleDimensions.CircleA
                            .CircleBSize = udtCircleDimensions.CircleB
                            .OverlapSize = udtCircleDimensions.OverlapAB
                        End With
                    End If

                    If mMessageQueueCount > 0 OrElse (mMessageQueueCount = 0 AndAlso txtStatus.TextLength > 0) Then
                        If chkHideMessagesOnSuccessfulUpdate.Checked Then
                            ' Truncate to only display the most recent message and to hide it in 1 second
                            mMessageQueueCount = 1
                            mMessageQueue(0).DisplayTime = System.DateTime.Now.Subtract(New System.TimeSpan(0, 0, DEFAULT_SECONDS_TO_DISPLAY_EACH_MESSAGE - 1))
                        End If
                    End If

                    StoreCurrentDimensions(udtCircleDimensions)
                Else
                    RestorePreviousDimensions(mCircleDimensionsSaved, True)
                End If

                If blnUpdateCountDistinct Then
                    UpdateCountDistinct()
                End If

            End If

        Catch ex As Exception
            MsgBox("Error refreshing Venn diagrams: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub ResetImageAdjustmentValues()
        tbarImgRotation.Value = 0
        tbarImgZoom.Value = 100
        tbarImgXOffset.Value = 0
        tbarImgYOffset.Value = 0
    End Sub

    Private Sub ResetValues(ByVal blnResetSettings As Boolean)
        Me.txtSizeA.Text = DEFAULT_DATA_CIRCLE_A.ToString
        Me.txtSizeB.Text = DEFAULT_DATA_CIRCLE_B.ToString
        Me.txtSizeC.Text = DEFAULT_DATA_CIRCLE_C.ToString
        Me.txtSizeOverlapAB.Text = DEFAULT_DATA_OVERLAP_AB.ToString
        Me.txtSizeOverlapBC.Text = DEFAULT_DATA_OVERLAP_BC.ToString
        Me.txtSizeOverlapAC.Text = DEFAULT_DATA_OVERLAP_AC.ToString

        vdgTwoCircles.VennDiagram.CircleAColor = vdgTwoCircles.VennDiagram.DefaultColorCircleA
        vdgTwoCircles.VennDiagram.CircleBColor = vdgTwoCircles.VennDiagram.DefaultColorCircleB
        vdgTwoCircles.VennDiagram.OverlapABColor = vdgTwoCircles.VennDiagram.DefaultColorOverlapAB

        vdgTwoCircles.VennDiagram.BackColor = Color.White

        vdgThreeCircles.VennDiagram.CircleAColor = vdgThreeCircles.VennDiagram.DefaultColorCircleA
        vdgThreeCircles.VennDiagram.CircleBColor = vdgThreeCircles.VennDiagram.DefaultColorCircleB
        vdgThreeCircles.VennDiagram.CircleCColor = vdgThreeCircles.VennDiagram.DefaultColorCircleC

        vdgThreeCircles.VennDiagram.OverlapABColor = vdgThreeCircles.VennDiagram.DefaultColorOverlapAB
        vdgThreeCircles.VennDiagram.OverlapBCColor = vdgThreeCircles.VennDiagram.DefaultColorOverlapBC
        vdgThreeCircles.VennDiagram.OverlapACColor = vdgThreeCircles.VennDiagram.DefaultColorOverlapAC

        vdgThreeCircles.VennDiagram.BackColor = Color.White

        If blnResetSettings Then
            AutoSizeWindow()

            Me.MessageDisplayTime = DEFAULT_SECONDS_TO_DISPLAY_EACH_MESSAGE
            Me.DuplicateMessageIgnoreWindow = DEFAULT_DUPLICATE_IGNORE_WINDOW_SECONDS
            chkHideMessagesOnSuccessfulUpdate.Checked = False
            chkFillCirclesWithColor.Checked = True

            ResetImageAdjustmentValues()
        End If

        RefreshVennDiagrams(True, True)
        UpdateThreeCircleDistinctCounts()

    End Sub

    Private Sub RestorePreviousDimensions(ByVal udtCircleDimensions As udtCircleDimensionsType, ByVal blnInformUser As Boolean)
        If blnInformUser Then
            MsgBox("Invalid overlap values; restoring the previous, valid values.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Invalid numbers")
        End If

        With udtCircleDimensions
            Me.txtSizeA.Text = NumToString(.CircleA, 1)
            Me.txtSizeB.Text = NumToString(.CircleB, 1)
            Me.txtSizeC.Text = NumToString(.CircleC, 1)
            Me.txtSizeOverlapAB.Text = NumToString(.OverlapAB, 1)
            Me.txtSizeOverlapBC.Text = NumToString(.OverlapBC, 1)
            Me.txtSizeOverlapAC.Text = NumToString(.OverlapAC, 1)
        End With

    End Sub

    Private Sub SaveDefaults()
        Dim outputFilepath As String = String.Empty
        Dim objOutFile As System.IO.StreamWriter

        Dim objXmlFile As New XmlSettingsFileAccessor

        Try
            outputFilepath = GetIniFilePath(mIniFileName)

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

            With objXmlFile
                ' Pass True to .LoadSettings() to turn on case sensitive matching
                .LoadSettings(outputFilepath, True)

                Try
                    If ValidDimensionsPresent(False) Then
                        .SetParam(XML_SECTION_OPTIONS, "CircleADia", txtSizeA.Text)
                        .SetParam(XML_SECTION_OPTIONS, "CircleBDia", txtSizeB.Text)
                        .SetParam(XML_SECTION_OPTIONS, "CircleCDia", txtSizeC.Text)
                        .SetParam(XML_SECTION_OPTIONS, "Overlap", txtSizeOverlapAB.Text)
                        .SetParam(XML_SECTION_OPTIONS, "OverlapBC", txtSizeOverlapBC.Text)
                        .SetParam(XML_SECTION_OPTIONS, "OverlapAC", txtSizeOverlapAC.Text)
                    End If
                    .SetParam(XML_SECTION_OPTIONS, "CircleAColor", ColorToString(vdgTwoCircles.VennDiagram.CircleAColor))
                    .SetParam(XML_SECTION_OPTIONS, "CircleBColor", ColorToString(vdgTwoCircles.VennDiagram.CircleBColor))
                    .SetParam(XML_SECTION_OPTIONS, "CircleCColor", ColorToString(vdgThreeCircles.VennDiagram.CircleCColor))
                    .SetParam(XML_SECTION_OPTIONS, "OverlapColor", ColorToString(vdgTwoCircles.VennDiagram.OverlapColor))
                    .SetParam(XML_SECTION_OPTIONS, "OverlapBCColor", ColorToString(vdgThreeCircles.VennDiagram.OverlapBCColor))
                    .SetParam(XML_SECTION_OPTIONS, "OverlapACColor", ColorToString(vdgThreeCircles.VennDiagram.OverlapACColor))
                    .SetParam(XML_SECTION_OPTIONS, "OverlapABCColor", ColorToString(vdgThreeCircles.VennDiagram.OverlapABCColor))
                    .SetParam(XML_SECTION_OPTIONS, "FillCirclesWithColor", chkFillCirclesWithColor.Checked)

                    .SetParam(XML_SECTION_OPTIONS, "EnterCircleTotals", optTotal.Checked)

                    .SetParam(XML_SECTION_OPTIONS, "ThreeCircleMode", chkCircleC.Checked)

                    .SetParam(XML_SECTION_OPTIONS, "MessageDisplayTime.", Me.MessageDisplayTime)
                    .SetParam(XML_SECTION_OPTIONS, "DuplicateMessageIgnoreWindow", Me.DuplicateMessageIgnoreWindow)
                    .SetParam(XML_SECTION_OPTIONS, "HideMessagesOnSuccessfulUpdate", chkHideMessagesOnSuccessfulUpdate.Checked)

                    .SetParam(XML_SECTION_OPTIONS, "RegionCountMode", cboRegionCountMode.SelectedIndex)
                    .SetParam(XML_SECTION_OPTIONS, "RegionCountValue", txtRegionCountValue.Text)


                    .SetParam(XML_SECTION_OPTIONS, "BackgroundColor", ColorToString(vdgTwoCircles.VennDiagram.BackColor))
                    .SetParam(XML_SECTION_OPTIONS, "WindowWidth", Me.Width.ToString)
                    .SetParam(XML_SECTION_OPTIONS, "WindowHeight", Me.Height.ToString)

                    .SaveSettings()
                Catch ex As Exception
                    System.Windows.Forms.MessageBox.Show("Error storing parameter in settings file: " & outputFilepath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try
            End With

        Catch ex As Exception
            MsgBox("Error saving Defaults to " & outputFilepath & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try
    End Sub

    Private Sub SaveGraphicToDisk()
        If chkCircleC.Checked Then
            SaveGraphicToDisk(vdgThreeCircles.VennDiagram)
        Else
            SaveGraphicToDisk(vdgTwoCircles.VennDiagram)
        End If
    End Sub

    Private Sub SaveGraphicToDisk(ByVal objVennDiagrams As VennDiagrams.VennDiagramBaseClass)
        Dim outputFilepath As String = String.Empty
        Dim ext As String

        Try
            Dim bitmap As Bitmap = New Bitmap(objVennDiagrams.Width, objVennDiagrams.Height)
            Dim g As Graphics = Graphics.FromImage(bitmap)

            dlgSave.Filter = "bmp files (*.bmp)|*.bmp|png files (*.png)|*.png"

            If dlgSave.ShowDialog = DialogResult.OK Then
                outputFilepath = dlgSave.FileName

                ControlPrinter.ControlPrinter.DrawControl(objVennDiagrams, g, True)
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
            .SetToolTip(cmdRefresh, "Refresh plot (Ctrl+R)")
            .SetToolTip(cmdSaveToDisk, "Save plot to disk (Ctrl+S)")
            .SetToolTip(cmdCopyToClipboard, "Copy plot to clipboard (Ctrl+C)")

            .SetToolTip(txtSizeOverlapAB, "Overlap of circles A and B")
            .SetToolTip(txtSizeOverlapBC, "Overlap of circles B and C")
            .SetToolTip(txtSizeOverlapAC, "Overlap of circles A and C")

            .SetToolTip(txtSizeA, "Total size of Circle A")
            .SetToolTip(txtSizeB, "Total size of Circle B")
            .SetToolTip(txtSizeC, "Total size of Circle C")

            .SetToolTip(txtDistinctA, "Number of items only in Circle A")
            .SetToolTip(txtDistinctB, "Number of items only in Circle B")

            .SetToolTip(tbarMessageDisplayTimeSeconds, "Number of seconds to display each message in the status window")
            .SetToolTip(tbarDuplicateMessageIgnoreWindowSeconds, "Number of seconds between which duplicate messages will be displayed")

            .SetToolTip(chkHideMessagesOnSuccessfulUpdate, "Hide old messages 1 second after a Venn Diagram plot is successfully made.  Additionally, clears any displayed messages immediately when the checkbox is first checked.")

        End With

    End Sub

    Private Sub SetTrackbarValue(ByRef objTrackbar As System.Windows.Forms.TrackBar, ByVal intValue As Integer)

        If intValue <= objTrackbar.Minimum Then
            intValue = objTrackbar.Minimum
        ElseIf intValue >= objTrackbar.Maximum Then
            intValue = objTrackbar.Maximum
        End If

        objTrackbar.Value = intValue
    End Sub

    Private Sub ShowAboutBox()
        Dim strMessage As String

        strMessage = String.Empty
        strMessage &= "Program originally written by Kyle Littlefield for the Department of Energy (PNNL, Richland, WA) in 2004" & ControlChars.NewLine
        strMessage &= "Three circle overlap added by Matthew Monroe (PNNL, Richland, WA) in 2007" & ControlChars.NewLine & ControlChars.NewLine

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

    Private Sub StoreCurrentDimensions(ByVal udtCircleDimensions As udtCircleDimensionsType)
        mCircleDimensionsSaved = udtCircleDimensions
    End Sub

    Public Function StringToColor(ByVal s As String) As Color

        Return CType(System.ComponentModel.TypeDescriptor.GetConverter(GetType(Color)).ConvertFromString(s), Color)

    End Function

    Private Function TextBoxToDbl(ByVal objTextbox As TextBox) As Double
        Return TextBoxToDbl(objTextbox, 0)
    End Function

    Private Function TextBoxToDbl(ByVal objTextbox As TextBox, ByVal dblDefaultValue As Double) As Double
        Try
            Return Double.Parse(objTextbox.Text)
        Catch ex As Exception
            Return dblDefaultValue
        End Try
    End Function

    Private Sub UpdateCountDistinct()
        UpdateCountDistinctA()
        UpdateCountDistinctB()

        If chkCircleC.Checked Then
            UpdateThreeCircleDistinctCounts()
        End If
    End Sub

    Private Sub UpdateCountDistinctA()
        Dim dblSizeA As Double
        Dim dblDistinctA As Double
        Dim dblOverlap As Double

        Try
            If IsNumber(txtSizeA) AndAlso IsNumber(txtSizeOverlapAB) Then
                dblSizeA = TextBoxToDbl(Me.txtSizeA)
                dblOverlap = TextBoxToDbl(Me.txtSizeOverlapAB)

                dblDistinctA = dblSizeA - dblOverlap
                If dblDistinctA >= 0 Then
                    txtDistinctA.Text = NumToString(dblDistinctA, 1)
                Else
                    UpdateStatus("Error: Circle A's size is less than the overlap value; unable to update Count Distinct A")
                End If
            End If
        Catch ex As Exception
            ' Error parsing number; do not continue
        End Try
    End Sub

    Private Sub UpdateCountDistinctB()
        Dim dblSizeB As Double
        Dim dblDistinctB As Double
        Dim dblOverlap As Double

        Try
            If IsNumber(txtSizeB) AndAlso IsNumber(txtSizeOverlapAB) Then
                dblSizeB = TextBoxToDbl(Me.txtSizeB)
                dblOverlap = TextBoxToDbl(Me.txtSizeOverlapAB)

                dblDistinctB = dblSizeB - dblOverlap
                If dblDistinctB >= 0 Then
                    txtDistinctB.Text = NumToString(dblDistinctB, 1)
                Else
                    UpdateStatus("Error: Circle B's size is less than the overlap value; unable to update Count Distinct B")
                End If
            End If
        Catch ex As Exception
            ' Error parsing number; do not continue
        End Try

    End Sub

    Private Sub UpdateImagePositionValues()
        lblImgRotation.Text = "Rotation: " & tbarImgRotation.Value.ToString
        lblImgZoom.Text = "Zoom: " & tbarImgZoom.Value.ToString & "%"
        lblImgOffsetX.Text = "X Offset: " & tbarImgXOffset.Value.ToString & "%"
        lblImgOffsetY.Text = "Y Offset: " & (-tbarImgYOffset.Value).ToString & "%"

        vdgThreeCircles.VennDiagram.Rotation = tbarImgRotation.Value
        vdgThreeCircles.VennDiagram.ZoomPct = tbarImgZoom.Value
        vdgThreeCircles.VennDiagram.XOffset = tbarImgXOffset.Value
        vdgThreeCircles.VennDiagram.YOffset = -tbarImgYOffset.Value

        RefreshVennDiagrams(False)
    End Sub

    Private Sub UpdateThreeCircleDistinctCounts()
        Dim eRegionCountMode As eRegionCountModeConstants
        Dim dblRegionCountValue As Double

        Dim udtCircleRegions As VennDiagrams.ThreeCircleVennDiagram.udtThreeCircleRegionsType

        Dim dblPossibleUniqueItemCount As Double

        Dim blnSuccess As Boolean

        eRegionCountMode = GetRegionCountMode()

        If Not IsNumber(txtRegionCountValue) Then
            ' Auto-compute the value for txtRegionCountValue
            ComputeOptimalRegionCount(False)
        End If

        If Not IsNumber(txtRegionCountValue) Then
            UpdateStatus("Region Count Value must be a number to update the three circle overlap distinct counts")
            Exit Sub
        Else
            dblRegionCountValue = TextBoxToDbl(txtRegionCountValue)
        End If

        ' Get the currently defined circle dimensions
        If GetCircleDimensions(udtCircleRegions, True, True) Then

            ' Sum up the sizes of each of the circles to compute dblPossibleUniqueItemCount
            dblPossibleUniqueItemCount = udtCircleRegions.CircleA + udtCircleRegions.CircleB + udtCircleRegions.CircleC

            If eRegionCountMode = eRegionCountModeConstants.UniqueItemCountAcrossAllCircles Then
                If dblRegionCountValue > dblPossibleUniqueItemCount Then
                    UpdateStatus("Specified unique item count of " & NumToString(dblRegionCountValue, 0) & " is impossibly large; it cannot be larger than A+B+C = " & NumToString(dblPossibleUniqueItemCount, 0))
                End If
            End If

            Select Case eRegionCountMode
                Case eRegionCountModeConstants.CountOverlappingAllCircles
                    blnSuccess = vdgThreeCircles.VennDiagram.ComputeThreeCircleAreasGivenOverlapABC(udtCircleRegions, dblRegionCountValue)

                    If dblRegionCountValue <= 0 Then
                        UpdateStatus("Region count value of " & NumToString(dblRegionCountValue, 1) & " is <= 0; this is not allowable; using optimal value of " & NumToString(udtCircleRegions.ABC, 0) & " instead")
                    End If

                Case eRegionCountModeConstants.UniqueItemCountAcrossAllCircles
                    blnSuccess = vdgThreeCircles.VennDiagram.ComputeThreeCircleAreasGivenTotalUniqueCount(udtCircleRegions, dblRegionCountValue)
                Case Else
                    ' Unknown value
                    UpdateStatus("Unknown Region Count Mode in Sub UpdateThreeCircleDistinctCounts")
                    blnSuccess = False
            End Select

            If blnSuccess Then
                With udtCircleRegions
                    DisplayDistinctRegionCount(.Ax, txtRegionAx, "only in circle A")
                    DisplayDistinctRegionCount(.Bx, txtRegionBx, "only in circle B")
                    DisplayDistinctRegionCount(.Cx, txtRegionCx, "only in circle C")
                    DisplayDistinctRegionCount(.ABx, txtRegionABx, "in the A/B overlap region")
                    DisplayDistinctRegionCount(.BCx, txtRegionBCx, "in the B/C overlap region")
                    DisplayDistinctRegionCount(.ACx, txtRegionACx, "in the A/C overlap region")
                    DisplayDistinctRegionCount(.ABC, txtRegionABC, "in the A/B/C overlap region")
                    DisplayDistinctRegionCount(.TotalUniqueCount, txtRegionUniqueItemCountAllCircles, "")
                End With
            End If

        End If

    End Sub

    Private Sub UpdateSizeTotals()
        UpdateSizeTotalA()
        UpdateSizeTotalB()

        If chkCircleC.Checked Then
            'UpdateSizeTotalC()
        End If
    End Sub

    Private Sub UpdateSizeTotalA()
        Dim dblSizeA As Double
        Dim dblDistinctA As Double
        Dim dblOverlap As Double

        Try
            If IsNumber(txtDistinctA) AndAlso IsNumber(txtSizeOverlapAB) Then
                dblDistinctA = TextBoxToDbl(Me.txtDistinctA)
                dblOverlap = TextBoxToDbl(Me.txtSizeOverlapAB)

                dblSizeA = dblDistinctA + dblOverlap
                txtSizeA.Text = NumToString(dblSizeA, 1)
            End If
        Catch ex As Exception
            ' Error parsing number; do not continue
        End Try
    End Sub

    Private Sub UpdateSizeTotalB()
        Dim dblSizeB As Double
        Dim dblDistinctB As Double
        Dim dblOverlap As Double

        Try
            If IsNumber(txtDistinctB) AndAlso IsNumber(txtSizeOverlapAB) Then
                dblDistinctB = TextBoxToDbl(Me.txtDistinctB)
                dblOverlap = TextBoxToDbl(Me.txtSizeOverlapAB)

                dblSizeB = dblDistinctB + dblOverlap
                txtSizeB.Text = NumToString(dblSizeB, 1)
            End If
        Catch ex As Exception
            ' Error parsing number; do not continue
        End Try
    End Sub

    Private Sub UpdateStatus(ByVal strMessage As String)
        UpdateStatus(strMessage, True)
    End Sub

    Private Sub UpdateStatus(ByVal strMessage As String, ByVal blnUseMessageQueue As Boolean)

        If blnUseMessageQueue Then
            AppendToStatusLog(strMessage)
        Else
            mMessageQueueCount = 0
            txtStatus.Text = strMessage
            If txtStatus.Text.Length > 0 Then
                txtStatus.Height = STATUS_HEIGHT_SINGLE_LINE
            Else
                txtStatus.Height = STATUS_HEIGHT_NO_MESSAGE
            End If
            txtStatus.ScrollBars = ScrollBars.None

        End If
    End Sub

    Private Sub RefreshMessageListUsingQueue(ByVal sender As Object, ByVal e As System.EventArgs) Handles mStatusTimer.Tick

        Dim intIndex As Integer
        Dim intFirstIndexToKeep As Integer
        Dim dblElapsedTimeSeconds As Double

        If mMessageQueueCount > 0 Then
            ' See if any messages should be removed from DisplayStatusLog

            intFirstIndexToKeep = 0
            For intIndex = 0 To mMessageQueueCount - 1
                dblElapsedTimeSeconds = System.DateTime.Now.Subtract(mMessageQueue(intIndex).DisplayTime).TotalSeconds
                If dblElapsedTimeSeconds >= Me.MessageDisplayTime Then
                    intFirstIndexToKeep = intIndex + 1
                End If
            Next intIndex

            If intFirstIndexToKeep > 0 Then
                ' Remove some entries

                If intFirstIndexToKeep >= mMessageQueueCount Then
                    ' Remove all entries
                    mMessageQueueCount = 0
                Else
                    ' Remove some of the entries
                    For intIndex = intFirstIndexToKeep To mMessageQueueCount - 1
                        mMessageQueue(intIndex - intFirstIndexToKeep) = mMessageQueue(intIndex)
                    Next intIndex
                    mMessageQueueCount -= intFirstIndexToKeep
                End If

                DisplayStatusLog()
            End If

        End If
    End Sub

    Private Sub ValidateUpDownControlValue(ByRef objControl As System.Windows.Forms.NumericUpDown)
        If objControl.Value < objControl.Minimum Then
            objControl.Value = objControl.Minimum
        ElseIf objControl.Value > objControl.Maximum Then
            objControl.Value = objControl.Maximum
        End If
    End Sub

    Private Function ValidDimensionsPresent() As Boolean
        Return ValidDimensionsPresent(True)
    End Function

    Private Function ValidDimensionsPresent(ByVal blnWarnIfInvalid As Boolean) As Boolean
        Dim udtCircleDimensions As udtCircleDimensionsType

        If GetCircleDimensions(udtCircleDimensions, chkCircleC.Checked, False) Then
            Return ValidDimensionsPresent(udtCircleDimensions, blnWarnIfInvalid)
        Else
            Return False
        End If

    End Function

    Private Function ValidDimensionsPresent(ByVal udtCircleDimensions As udtCircleDimensionsType, ByVal blnWarnIfInvalid As Boolean) As Boolean
        Dim blnValidAB As Boolean
        Dim blnValidBC As Boolean
        Dim blnValidAC As Boolean
        Dim blnValidDims As Boolean

        Dim strMessage As String

        Try
            blnValidAB = False
            blnValidBC = False
            blnValidAC = False
            strMessage = String.Empty

            If udtCircleDimensions.CircleA < udtCircleDimensions.OverlapAB Then
                If blnWarnIfInvalid Then
                    strMessage &= "Circle A is smaller than the specified A/B Overlap Value;"
                End If
            ElseIf udtCircleDimensions.CircleB < udtCircleDimensions.OverlapAB Then
                If blnWarnIfInvalid Then
                    strMessage &= "Circle B is smaller than the specified A/B Overlap Value;"
                End If
            Else
                blnValidAB = True
            End If

            If chkCircleC.Checked AndAlso blnValidAB Then
                If udtCircleDimensions.CircleB < udtCircleDimensions.OverlapBC Then
                    If blnWarnIfInvalid Then
                        strMessage &= "Circle B is smaller than the specified B/C Overlap Value;"
                    End If
                ElseIf udtCircleDimensions.CircleC < udtCircleDimensions.OverlapBC Then
                    If blnWarnIfInvalid Then
                        strMessage &= "Circle C is smaller than the specified B/C Overlap Value;"
                    End If
                Else
                    blnValidBC = True

                    If udtCircleDimensions.CircleA < udtCircleDimensions.OverlapAC Then
                        If blnWarnIfInvalid Then
                            strMessage &= "Circle A is smaller than the specified A/C Overlap Value;"
                        End If
                    ElseIf udtCircleDimensions.CircleC < udtCircleDimensions.OverlapAC Then
                        If blnWarnIfInvalid Then
                            strMessage &= "Circle C is smaller than the specified A/C Overlap Value;"
                        End If
                    Else
                        blnValidAC = True
                    End If
                End If
                blnValidDims = blnValidAB AndAlso blnValidBC AndAlso blnValidAC
            Else
                blnValidDims = blnValidAB
            End If

        Catch ex As Exception
            Return False
        End Try

        If strMessage.Length > 0 Then
            UpdateStatus("Error: " & strMessage.TrimEnd(";"c))
        End If

        Return blnValidDims
    End Function

#Region "Button Handlers"

    Private Sub cmdBackgroundColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = Me.BackColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgTwoCircles.VennDiagram.BackColor = dlgColor.Color
            vdgThreeCircles.VennDiagram.BackColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdCopyToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyToClipboard.Click
        CopyVennToClipboard()
    End Sub

    Private Sub cmdCircleAColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgTwoCircles.VennDiagram.CircleAColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgTwoCircles.VennDiagram.CircleAColor = dlgColor.Color
            vdgThreeCircles.VennDiagram.CircleAColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdCircleBColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgTwoCircles.VennDiagram.CircleBColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgTwoCircles.VennDiagram.CircleBColor = dlgColor.Color
            vdgThreeCircles.VennDiagram.CircleBColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdCircleCColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgThreeCircles.VennDiagram.CircleCColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgThreeCircles.VennDiagram.CircleCColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdRegionCountComputeOptimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegionCountComputeOptimal.Click
        ComputeOptimalRegionCount(True)
    End Sub

    Private Sub cmdOverlapABColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgTwoCircles.VennDiagram.OverlapColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgTwoCircles.VennDiagram.OverlapColor = dlgColor.Color
            vdgThreeCircles.VennDiagram.OverlapABColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdOverlapBCColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgThreeCircles.VennDiagram.OverlapBCColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgThreeCircles.VennDiagram.OverlapBCColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdOverlapACColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgThreeCircles.VennDiagram.OverlapACColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgThreeCircles.VennDiagram.OverlapACColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdOverlapABCColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dlgColor.Color = vdgThreeCircles.VennDiagram.OverlapABCColor
        If dlgColor.ShowDialog() = DialogResult.OK Then
            vdgThreeCircles.VennDiagram.OverlapABCColor = dlgColor.Color
            RefreshVennDiagrams(False)
        End If
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        RefreshVennDiagrams(True)
    End Sub

    Private Sub cmdImgAdjustmentReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImgAdjustmentReset.Click
        ResetImageAdjustmentValues()
        RefreshVennDiagrams(False)
    End Sub

    Private Sub cmdSaveToDisk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveToDisk.Click
        SaveGraphicToDisk()
    End Sub

    Private Sub cmdUpdateRegionCounts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateRegionCounts.Click
        UpdateThreeCircleDistinctCounts()
    End Sub

#End Region

#Region "Combobox And Trackbar Handlers"
    Private Sub cboRegionCountMode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRegionCountMode.SelectedIndexChanged
        AutoUpdateRegionCountValue()
    End Sub

    Private Sub tbarMessageDisplayTimeSeconds_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbarMessageDisplayTimeSeconds.ValueChanged
        lblMessageDisplayTimeSeconds.Text = tbarMessageDisplayTimeSeconds.Value.ToString & " sec."
    End Sub

    Private Sub tbarDuplicateMessageIgnoreWindowSeconds_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbarDuplicateMessageIgnoreWindowSeconds.ValueChanged
        lblMessageDuplicateIgnoreWindow.Text = tbarDuplicateMessageIgnoreWindowSeconds.Value.ToString & " sec."
    End Sub

    Private Sub tbarImgXOffset_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbarImgXOffset.ValueChanged
        UpdateImagePositionValues()
    End Sub

    Private Sub tbarImgYOffset_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbarImgYOffset.ValueChanged
        UpdateImagePositionValues()
    End Sub

    Private Sub tbarImgRotation_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbarImgRotation.ValueChanged
        UpdateImagePositionValues()
    End Sub

    Private Sub tbarImgZoom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbarImgZoom.ValueChanged
        UpdateImagePositionValues()
    End Sub

#End Region

#Region "Checkbox and Option Button Handlers"

    Private Sub chkCircleC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCircleC.CheckedChanged
        EnableDisableControls()
        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        AutoSizeWindow()
        RefreshVennDiagrams(False)
    End Sub

    Private Sub chkFillCirclesWithColor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not vdgTwoCircles Is Nothing Then
            vdgTwoCircles.VennDiagram.PaintSolidColorCircles = chkFillCirclesWithColor.Checked
        End If
        If Not vdgThreeCircles Is Nothing Then
            vdgThreeCircles.VennDiagram.PaintSolidColorCircles = chkFillCirclesWithColor.Checked
        End If
        RefreshVennDiagrams(False)
    End Sub

    Private Sub chkHideMessagesOnSuccessfulUpdate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideMessagesOnSuccessfulUpdate.CheckedChanged
        If chkHideMessagesOnSuccessfulUpdate.Checked Then
            ClearMessages()
        End If
    End Sub

    Private Sub optTotal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTotal.CheckedChanged
        EnableDisableControls()
    End Sub

    Private Sub optCountDistinct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCountDistinct.CheckedChanged
        EnableDisableControls()
    End Sub

#End Region

#Region "Menu Handlers"

    Private Sub mnuEditCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditCopy.Click
        CopyVennToClipboard()
    End Sub

    Private Sub mnuEditRefreshPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditRefreshPlot.Click
        RefreshVennDiagrams(False)
    End Sub

    Private Sub mnuEditResetValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditResetValues.Click
        ResetValues(False)
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        ShowAboutBox()
    End Sub

    Private Sub mnuLoadDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLoadDefaults.Click
        LoadDefaults()
    End Sub

    Private Sub mnuReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReset.Click
        ResetValues(True)
    End Sub

    Private Sub mnuSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveFile.Click
        SaveGraphicToDisk()
    End Sub

    Private Sub mnuSaveDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveDefaults.Click
        SaveDefaults()
    End Sub
#End Region

#Region "Textbox Handlers"
    Private Sub txtDistinctA_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDistinctA.LostFocus
        UpdateSizeTotalA()
        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtDistinctB_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDistinctB.LostFocus
        UpdateSizeTotalB()
        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeA_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeA.LostFocus
        UpdateCountDistinctA()
        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeB_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeB.LostFocus
        UpdateCountDistinctB()
        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeC_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeC.LostFocus
        UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeOverlap_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeOverlapAB.LostFocus
        If optCountDistinct.Checked Then
            UpdateSizeTotals()
        Else
            UpdateCountDistinct()
        End If

        If chkCircleC.Checked Then UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeOverlapBC_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeOverlapBC.LostFocus
        If optCountDistinct.Checked Then
            UpdateSizeTotals()
        Else
            UpdateCountDistinct()
        End If

        UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

    Private Sub txtSizeOverlapAC_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSizeOverlapAC.LostFocus
        If optCountDistinct.Checked Then
            UpdateSizeTotals()
        Else
            UpdateCountDistinct()
        End If

        UpdateThreeCircleDistinctCounts()
        If ValidDimensionsPresent() Then RefreshVennDiagrams(False)
    End Sub

#End Region

End Class
