Option Strict On

Imports System.Drawing

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'

Public Class ThreeCircleVennDiagram
     Inherits VennDiagramBaseClass

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
        'ThreeCircleVennDiagram
        '
        Me.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ThreeCircleVennDiagram"

    End Sub

#End Region

#Region "Events"
    'Raised when smoothingmode, FillFactor, or colors are changed
    Public Event DrawingChangeCircleC(ByVal sender As ThreeCircleVennDiagram)

    'Never raised
    Public Event LabelChangeCircleC(ByVal sender As ThreeCircleVennDiagram)

    'Raised when size of circles or overlap is changed.
    Public Event SizeChangeCircleC(ByVal sender As ThreeCircleVennDiagram)
#End Region

#Region "Structures"
    Public Structure udtThreeCircleRegionsType
        Public CircleA As Double
        Public CircleB As Double
        Public CircleC As Double
        Public OverlapAB As Double
        Public OverlapBC As Double
        Public OverlapAC As Double
        Public Ax As Double
        Public Bx As Double
        Public Cx As Double
        Public ABx As Double
        Public BCx As Double
        Public ACx As Double
        Public ABC As Double
        Public TotalUniqueCount As Double         ' Unique item count across all circles
    End Structure

    Protected Structure udtArcAngleAdjustmentType
        Public AlphaAdd As Double               ' Angle adjustment amount (in degrees)
        Public BetaAdd As Double                ' Angle adjustment amount (in degrees)
    End Structure
#End Region

#Region "Member Variables"
    ' Note: Location, size, etc. for each circle is held in the DrawingPath property of the VennDiagramAreaInfo
    Protected m_circleC As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorCircleC)

    Protected m_overlapBC As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorOverlapBC)
    Protected m_overlapAC As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorOverlapAC)
    Protected m_overlapABC As VennDiagramAreaInfo = New VennDiagramAreaInfo(Me.DefaultColorOverlapABC)

    Protected mOverlapBCArcAdj As udtArcAngleAdjustmentType
    Protected mOverlapACArcAdj As udtArcAngleAdjustmentType

    Protected mCirclesAB As TwoCircleVennDiagram
    Protected mCirclesBC As TwoCircleVennDiagram
    Protected mCirclesAC As TwoCircleVennDiagram

    'Computation World Coordinates
    Protected m_CircleC_Loc As udtPointXYType
    Protected m_CircleC_Radius_Compute As Double

    Protected m_ZoomPct As Integer         ' Allows fine-tuning of the zoom value of the Venn Diagram within the window; 100 means 100%
    Protected m_XOffset As Integer         ' Allows fine-tuning of the center of the Venn Diagram within the window; percentage of window width (-100 to 100)
    Protected m_YOffset As Integer         ' Allows fine-tuning of the center of the Venn Diagram within the window; percentage of window width (-100 to 100)
    Protected m_Rotation As Integer        ' Rotates the image by the given number of degrees (0 to 360)

    Protected m_CircleC_ScreenLoc As udtPointXYType
    Protected m_CircleC_ScreenRadius As Double

    Protected m_DrawOverlapRegionsOnError As Boolean = True
#End Region

#Region "Properties "
    Public Property CircleCSize() As Double
        Get
            Return Me.m_circleC.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_circleC.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChangeCircleC(Me)
        End Set
    End Property

    Public ReadOnly Property DefaultColorCircleC() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(255, 255, 192)
        End Get
    End Property

    Public ReadOnly Property DefaultColorOverlapBC() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(192, 192, 0)
        End Get
    End Property

    Public ReadOnly Property DefaultColorOverlapAC() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(128, 128, 255)
        End Get
    End Property

    Public ReadOnly Property DefaultColorOverlapABC() As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(192, 192, 192)
        End Get
    End Property

    Public Property OverlapBCSize() As Double
        Get
            Return Me.m_overlapBC.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_overlapBC.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapACSize() As Double
        Get
            Return Me.m_overlapAC.Size
        End Get
        Set(ByVal Value As Double)
            Me.m_overlapAC.Size = Value
            Me.InvalidateComputeWorldCoordinates()
            RaiseEvent SizeChangeCircleC(Me)
        End Set
    End Property

    'Labels are not currently used.  They were intended to provide
    'the ability to label the diagram with intelligently placed labels
    'on the actual diagram
    Public Property CircleCLabel() As String()
        Get
            Return Me.m_circleC.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_circleC.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapBCLabel() As String()
        Get
            Return Me.m_overlapBC.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_overlapBC.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapACLabel() As String()
        Get
            Return Me.m_overlapAC.Label
        End Get
        Set(ByVal Value() As String)
            Me.m_overlapAC.Label = Value
            Me.InvalidateScreenCoordinates()
            RaiseEvent LabelChangeCircleC(Me)
        End Set
    End Property

    Public Property CircleCColor() As Color
        Get
            Return Me.m_circleC.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_circleC.Color = Value
            'Me.UpdateOverlapColor()
            Me.Invalidate()
            RaiseEvent DrawingChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapBCColor() As Color
        Get
            Return Me.m_overlapBC.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_overlapBC.Color = Value
            Me.Invalidate()
            RaiseEvent DrawingChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapACColor() As Color
        Get
            Return Me.m_overlapAC.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_overlapAC.Color = Value
            Me.Invalidate()
            RaiseEvent DrawingChangeCircleC(Me)
        End Set
    End Property

    Public Property OverlapABCColor() As Color
        Get
            Return Me.m_overlapABC.Color
        End Get
        Set(ByVal Value As Color)
            Me.m_overlapABC.Color = Value
            Me.Invalidate()
            RaiseEvent DrawingChangeCircleC(Me)
        End Set
    End Property

    Public Property Rotation() As Integer
        Get
            Return m_Rotation
        End Get
        Set(ByVal Value As Integer)
            Do While Value < 0
                Value += 360
            Loop

            Do While Value > 360
                Value -= 360
            Loop

            m_Rotation = Value
        End Set
    End Property

    Public Property XOffset() As Integer
        Get
            Return m_XOffset
        End Get
        Set(ByVal Value As Integer)
            m_XOffset = Value
        End Set
    End Property

    Public Property YOffset() As Integer
        Get
            Return m_YOffset
        End Get
        Set(ByVal Value As Integer)
            m_YOffset = Value
        End Set
    End Property

    Public Property ZoomPct() As Integer
        Get
            Return m_ZoomPct
        End Get
        Set(ByVal Value As Integer)
            If Value < 1 Then Value = 1
            If Value > 200 Then Value = 200
            m_ZoomPct = Value
        End Set
    End Property
#End Region

    Protected Shared Function ComputeRotationAngle(ByVal udtPointA As udtPointXYType, ByVal udtPointB As udtPointXYType) As Double
        Dim dblDeltaX As Double
        Dim dblDeltaY As Double
        Dim dblAngleDegrees As Double

        dblDeltaX = udtPointB.X - udtPointA.X
        dblDeltaY = udtPointB.Y - udtPointA.Y

        If dblDeltaX = 0 Then
            If dblDeltaY > 0 Then
                dblAngleDegrees = 90
            Else
                dblAngleDegrees = 270
            End If
        Else
            dblAngleDegrees = ConvertRadiansToDegrees(Math.Atan(Math.Abs(dblDeltaY) / Math.Abs(dblDeltaX)))

            ' Possibly adjust dblAngleDegrees based on the location of udtPointB vs. udtPointA
            If dblDeltaY > 0 Then
                If dblDeltaX < 0 Then
                    dblAngleDegrees = 180 - dblAngleDegrees
                End If
            Else
                If dblDeltaX < 0 Then
                    dblAngleDegrees += 180
                Else
                    dblAngleDegrees = 360 - dblAngleDegrees
                End If
            End If
        End If

        Return dblAngleDegrees

    End Function

    Public Shared Sub ComputeRotationAngleTest()

        Dim x As Integer, y As Integer
        Dim udtPointA As udtPointXYType
        Dim udtPointB As udtPointXYType
        Dim dblAngle As Double

        udtPointA.X = 0
        udtPointA.Y = 0

        For x = -4 To 4 Step 2
            For y = -4 To 4 Step 2
                udtPointB.X = x
                udtPointB.Y = y

                dblAngle = ComputeRotationAngle(udtPointA, udtPointB)
            Next y
        Next x
    End Sub

    Public Function ComputeThreeCircleAreasGivenOverlapABC(ByVal dblUserSuppliedABC As Double) As Boolean
        Dim udtCircleRegions As udtThreeCircleRegionsType

        With udtCircleRegions
            .CircleA = m_circleA.Size
            .CircleB = m_circleb.Size
            .CircleC = m_circleC.Size
            .OverlapAB = m_overlapAB.Size
            .OverlapBC = m_overlapBC.Size
            .OverlapAC = m_overlapAC.Size
        End With

        Return ComputeThreeCircleAreasGivenOverlapABC(udtCircleRegions, dblUserSuppliedABC)
    End Function

    Public Shared Function ComputeThreeCircleAreasGivenOverlapABC(ByRef udtCircleRegions As udtThreeCircleRegionsType, ByVal dblUserSuppliedABC As Double) As Boolean

        Const DIM_COUNT As Integer = 6

        ' If the user provides a value for ABC, then the following equations can be used
        ' Ax+      ABx+    ACx = A-ABC
        '    Bx+   ABx+BCx     = B-ABC
        '       Cx+    BCx+ACx = C-ABC
        '          ABx         = AB-ABC
        '              BCx     = BC-ABC
        '                  ACx = AC-ABC
        '
        ' The following coefficients correspond to these equations (array is DIM_COUNT rows and DIM_COUNT columns:
        Dim dblCoefMatrix(,) As Double = {{1, 0, 0, 1, 0, 1}, _
                                          {0, 1, 0, 1, 1, 0}, _
                                          {0, 0, 1, 0, 1, 1}, _
                                          {0, 0, 0, 1, 0, 0}, _
                                          {0, 0, 0, 0, 1, 0}, _
                                          {0, 0, 0, 0, 0, 1}}

		' dblXValues is initialized by objMatrixSolver.SolveEquations
		' Will be of length DIM_COUNT
		Dim dblXValues() As Double = Nothing
        Dim objMatrixSolver As New clsMatrixSolver

        ' dblPotentialXValues() will have DIM_COUNT+1 rows and intPotentialXValsCount columns; 
        '  it holds the potential XValues, plus the ABC value used to compute the X Values
        Dim intPotentialXValsCount As Integer
        Dim dblPotentialXValues(,) As Double

		Dim strMessage As String = String.Empty
        Dim blnSuccess As Boolean
        Dim blnUseUserSuppliedABC As Boolean

        Dim dblConstants(DIM_COUNT - 1) As Double
        Dim dblCurrentABC As Double

        Dim intABCIteration As Integer
        Dim intMinimumABC As Integer
        Dim intMaximumABC As Integer
        Dim x As Integer
        Dim intBestIndex As Integer

        If dblUserSuppliedABC <= 0 Then
            intMinimumABC = 0
            intMaximumABC = CInt(Math.Ceiling(Math.Min(Math.Min(udtCircleRegions.OverlapAB, udtCircleRegions.OverlapBC), udtCircleRegions.OverlapAC)))
            blnUseUserSuppliedABC = False
        Else
            intMinimumABC = 0
            intMaximumABC = 0
            blnUseUserSuppliedABC = True
        End If

        ' Initially reserve space for 10 sets of results
        intPotentialXValsCount = 0
        ReDim dblPotentialXValues(DIM_COUNT, 9)

        For intABCIteration = intMinimumABC To intMaximumABC

            If blnUseUserSuppliedABC Then
                dblCurrentABC = dblUserSuppliedABC
            Else
                dblCurrentABC = intABCIteration
            End If

            With udtCircleRegions
                dblConstants(0) = .CircleA - dblCurrentABC
                dblConstants(1) = .CircleB - dblCurrentABC
                dblConstants(2) = .CircleC - dblCurrentABC
                dblConstants(3) = .OverlapAB - dblCurrentABC
                dblConstants(4) = .OverlapBC - dblCurrentABC
                dblConstants(5) = .OverlapAC - dblCurrentABC
            End With

            blnSuccess = objMatrixSolver.SolveEquations(dblCoefMatrix, dblConstants, dblXValues, strMessage)

            If blnSuccess Then
                'objMatrixSolver.PrintArray(dblConstants)
                'objMatrixSolver.PrintArray(dblXValues)

                '' Optionally check the results
                ''Dim dblResults() As Double
                ''objMatrixSolver.MultiplyMatrices(dblCoefMatrix, dblXValues, dblResults)
                ''dblResidual = 0
                ''For x = 0 To DIM_COUNT - 1
                ''    dblResidual += Math.Abs(dblConstants(x) - dblResults(x))
                ''Next x

                If blnUseUserSuppliedABC Then
                    ' dblXValues now contains the areas of each of the regions we're trying to solve for
                    ' Update udtCircleRegions then exit the for loop
                    With udtCircleRegions
                        .Ax = dblXValues(0)
                        .Bx = dblXValues(1)
                        .Cx = dblXValues(2)
                        .ABx = dblXValues(3)
                        .BCx = dblXValues(4)
                        .ACx = dblXValues(5)
                        .ABC = dblUserSuppliedABC
                    End With

                    blnSuccess = True
                    Exit For
                Else
                    ComputeThreeCircleAreasStoreIfAllPositive(dblXValues, intPotentialXValsCount, dblPotentialXValues, DIM_COUNT, dblCurrentABC)
                End If
            End If

        Next intABCIteration

        If Not blnUseUserSuppliedABC Then
            If intPotentialXValsCount > 0 Then
                ' Find the ABC value that gave the value for Ax, Bx, Cx, ABx, BCx, & ACx 
                ' with the smallest residual from the average of each of those values

                intBestIndex = ComputeThreeCircleAreasFindSmallestResidual(intPotentialXValsCount, dblPotentialXValues, DIM_COUNT)

                ' Populate dblXValues with the values located at intBestIndex
                For x = 0 To DIM_COUNT - 1
                    dblXValues(x) = dblPotentialXValues(x, intBestIndex)
                Next x
                dblCurrentABC = dblPotentialXValues(DIM_COUNT, intBestIndex)

                blnSuccess = True
            Else
                ' None of the results were valid; clear dblXValues
                Array.Clear(dblXValues, 0, DIM_COUNT)
                dblCurrentABC = 0
                blnSuccess = False
            End If

            ' Update udtCircleRegions with the best values, as just determined
            With udtCircleRegions
                .Ax = dblXValues(0)
                .Bx = dblXValues(1)
                .Cx = dblXValues(2)
                .ABx = dblXValues(3)
                .BCx = dblXValues(4)
                .ACx = dblXValues(5)
                .ABC = dblCurrentABC
            End With
        End If

        If blnSuccess Then
            ' Update TotalUniqueCount
            With udtCircleRegions
                .TotalUniqueCount = .Ax + .Bx + .Cx + .ABx + .BCx + .ACx + .ABC
            End With
        End If

        Return blnSuccess

    End Function

    Public Function ComputeThreeCircleAreasGivenTotalUniqueCount(ByVal dblUserSuppliedTotalUniqueCount As Double) As Boolean
        Dim udtCircleRegions As udtThreeCircleRegionsType

        With udtCircleRegions
            .CircleA = m_circleA.Size
            .CircleB = m_circleb.Size
            .CircleC = m_circleC.Size
            .OverlapAB = m_overlapAB.Size
            .OverlapBC = m_overlapBC.Size
            .OverlapAC = m_overlapAC.Size
        End With

        Return ComputeThreeCircleAreasGivenTotalUniqueCount(udtCircleRegions, dblUserSuppliedTotalUniqueCount)
    End Function

    Public Shared Function ComputeThreeCircleAreasGivenTotalUniqueCount(ByRef udtCircleRegions As udtThreeCircleRegionsType, ByVal dblUserSuppliedTotalUniqueCount As Double) As Boolean
        Const DIM_COUNT As Integer = 7

        ' If the user provides a value for N (the total distinct items present across all collections), then the following equations can be used
        ' Ax+      ABx+    ACx+ABC = A
        '    Bx+   ABx+BCx    +ABC = B
        '       Cx+    BCx+ACx+ABC = C
        '          ABx        +ABC = AB
        '              BCx    +ABC = BC
        '                  ACx+ABC = AC
        ' Ax+Bx+Cx+ABx+BCx+ACx+ABC = N
        '
        ' The following coefficients correspond to these equations (array is DIM_COUNT rows and DIM_COUNT columns:
        Dim dblCoefMatrix(,) As Double = {{1, 0, 0, 1, 0, 1, 1}, _
                                          {0, 1, 0, 1, 1, 0, 1}, _
                                          {0, 0, 1, 0, 1, 1, 1}, _
                                          {0, 0, 0, 1, 0, 0, 1}, _
                                          {0, 0, 0, 0, 1, 0, 1}, _
                                          {0, 0, 0, 0, 0, 1, 1}, _
                                          {1, 1, 1, 1, 1, 1, 1}}

		' dblXValues is initialized by objMatrixSolver.SolveEquations
		' Will be of length DIM_COUNT
		Dim dblXValues() As Double = Nothing
        Dim objMatrixSolver As New clsMatrixSolver


        ' dblPotentialXValues() will have DIM_COUNT+1 rows and intPotentialXValsCount columns; 
        '  it holds the potential XValues, plus the TotalUniqueCount value used to compute the X Values
        Dim intPotentialXValsCount As Integer
        Dim dblPotentialXValues(,) As Double

		Dim strMessage As String = String.Empty
        Dim blnSuccess As Boolean
        Dim blnUseUserSuppliedTotalUniqueCount As Boolean

        Dim dblConstants(DIM_COUNT - 1) As Double
        Dim dblCurrentTotalUniqueCount As Double

        Dim intTotalUniqueCountIteration As Integer
        Dim intMinimumTotalUniqueCount As Integer
        Dim intMaximumTotalUniqueCount As Integer
        Dim x As Integer
        Dim intBestIndex As Integer


        If dblUserSuppliedTotalUniqueCount = 0 Then
            intMinimumTotalUniqueCount = CInt(Math.Floor(Math.Min(Math.Min(udtCircleRegions.CircleA, udtCircleRegions.CircleB), udtCircleRegions.CircleC)))
            intMaximumTotalUniqueCount = CInt(Math.Ceiling(udtCircleRegions.CircleA + udtCircleRegions.CircleB + udtCircleRegions.CircleC))
            blnUseUserSuppliedTotalUniqueCount = False
        Else
            intMinimumTotalUniqueCount = 0
            intMaximumTotalUniqueCount = 0
            blnUseUserSuppliedTotalUniqueCount = True
        End If

        ' Initially reserve space for 10 sets of results
        intPotentialXValsCount = 0
        ReDim dblPotentialXValues(DIM_COUNT, 9)

        ' Populate everything in dblConstants except for dblConstants(6)
        With udtCircleRegions
            dblConstants(0) = .CircleA
            dblConstants(1) = .CircleB
            dblConstants(2) = .CircleC
            dblConstants(3) = .OverlapAB
            dblConstants(4) = .OverlapBC
            dblConstants(5) = .OverlapAC
        End With

        For intTotalUniqueCountIteration = intMinimumTotalUniqueCount To intMaximumTotalUniqueCount

            If blnUseUserSuppliedTotalUniqueCount Then
                dblCurrentTotalUniqueCount = dblUserSuppliedTotalUniqueCount
            Else
                dblCurrentTotalUniqueCount = intTotalUniqueCountIteration
            End If

            dblConstants(6) = dblCurrentTotalUniqueCount

            blnSuccess = objMatrixSolver.SolveEquations(dblCoefMatrix, dblConstants, dblXValues, strMessage)

            If blnSuccess Then
                ''objMatrixSolver.PrintArray(dblConstants)
                ''objMatrixSolver.PrintArray(dblXValues)

                '' Optionally check the results
                ''Dim dblResults() As Double
                ''Dim dblResidual As Double
                ''objMatrixSolver.MultiplyMatrices(dblCoefMatrix, dblXValues, dblResults)
                ''dblResidual = 0
                ''For x = 0 To DIM_COUNT - 1
                ''    dblResidual += Math.Abs(dblConstants(x) - dblResults(x))
                ''Next x

                If blnUseUserSuppliedTotalUniqueCount Then
                    ' dblXValues now contains the areas of each of the regions we're trying to solve for
                    ' Update udtCircleRegions then exit the for loop
                    With udtCircleRegions
                        .Ax = dblXValues(0)
                        .Bx = dblXValues(1)
                        .Cx = dblXValues(2)
                        .ABx = dblXValues(3)
                        .BCx = dblXValues(4)
                        .ACx = dblXValues(5)
                        .ABC = dblXValues(6)
                        .TotalUniqueCount = dblUserSuppliedTotalUniqueCount
                    End With

                    blnSuccess = True
                    Exit For
                Else
                    ComputeThreeCircleAreasStoreIfAllPositive(dblXValues, intPotentialXValsCount, dblPotentialXValues, DIM_COUNT, dblCurrentTotalUniqueCount)
                End If
            End If

        Next intTotalUniqueCountIteration

        If Not blnUseUserSuppliedTotalUniqueCount Then
            If intPotentialXValsCount > 0 Then
                ' Find the TotalUniqueCount value that gave the value for Ax, Bx, Cx, ABx, BCx, ACx, and ABC
                ' with the smallest residual from the average of each of those values

                intBestIndex = ComputeThreeCircleAreasFindSmallestResidual(intPotentialXValsCount, dblPotentialXValues, DIM_COUNT)

                ' Populate dblXValues with the values located at intBestIndex
                For x = 0 To DIM_COUNT - 1
                    dblXValues(x) = dblPotentialXValues(x, intBestIndex)
                Next x
                dblCurrentTotalUniqueCount = dblPotentialXValues(DIM_COUNT, intBestIndex)

                blnSuccess = True
            Else
                ' None of the results were valid; clear dblXValues
                Array.Clear(dblXValues, 0, DIM_COUNT)
                dblCurrentTotalUniqueCount = 0
                blnSuccess = False
            End If

            ' Update udtCircleRegions with the best values, as just determined
            With udtCircleRegions
                .Ax = dblXValues(0)
                .Bx = dblXValues(1)
                .Cx = dblXValues(2)
                .ABx = dblXValues(3)
                .BCx = dblXValues(4)
                .ACx = dblXValues(5)
                .ABC = dblXValues(6)
                .TotalUniqueCount = dblCurrentTotalUniqueCount
            End With

        End If

        Return blnSuccess

    End Function

    Private Shared Sub ComputeThreeCircleAreasStoreIfAllPositive(ByRef dblXValues() As Double, ByRef intPotentialXValsCount As Integer, ByRef dblPotentialXValues(,) As Double, ByVal DIM_COUNT As Integer, ByVal dblCurrentIterationValue As Double)
        Dim blnAllPositive As Boolean
        Dim x As Integer

        ' Check whether all values in dblXValues are >= 0
        blnAllPositive = True
        For x = 0 To DIM_COUNT - 1
            If dblXValues(x) < 0 Then
                blnAllPositive = False
                Exit For
            End If
        Next x

        If blnAllPositive Then
            ' Add results to array of possible values
            If intPotentialXValsCount >= UBound(dblPotentialXValues, 2) Then
                ReDim Preserve dblPotentialXValues(DIM_COUNT, intPotentialXValsCount * 2 - 1)
            End If

            For x = 0 To DIM_COUNT - 1
                dblPotentialXValues(x, intPotentialXValsCount) = dblXValues(x)
            Next x
            dblPotentialXValues(DIM_COUNT, intPotentialXValsCount) = dblCurrentIterationValue

            intPotentialXValsCount += 1
        End If

    End Sub

    Private Shared Function ComputeThreeCircleAreasFindSmallestResidual(ByVal intPotentialXValsCount As Integer, ByRef dblPotentialXValues(,) As Double, ByVal DIM_COUNT As Integer) As Integer
        Dim dblAverages(DIM_COUNT - 1) As Double
        Dim dblSumResidualSquared() As Double
        Dim dblSum As Double

        Dim x As Integer, y As Integer
        Dim intBestIndex As Integer

        ReDim dblSumResidualSquared(intPotentialXValsCount - 1)

        ' Process each region
        For x = 0 To DIM_COUNT - 1
            ' Compute the average
            dblSum = 0
            For y = 0 To intPotentialXValsCount - 1
                dblSum += dblPotentialXValues(x, y)
            Next y
            dblAverages(x) = dblSum / intPotentialXValsCount
        Next x

        For y = 0 To intPotentialXValsCount - 1
            ' Compute the sum of the squares of the residual values (observed minus average) for this potential X value
            dblSumResidualSquared(y) = 0
            For x = 0 To DIM_COUNT - 1
                ' Compute the residual
                dblSumResidualSquared(y) += (dblPotentialXValues(x, y) - dblAverages(x)) ^ 2
            Next x
        Next y

        ' Find the smallest value in dblSumResidualSquared(x)
        intBestIndex = 0
        For y = 1 To intPotentialXValsCount - 1
            If dblSumResidualSquared(y) < dblSumResidualSquared(intBestIndex) Then
                intBestIndex = y
            End If
        Next y

        Return intBestIndex

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

        ' Determine the the X,Y coordinates of the box that would surround the 3 circles
        Dim udtWorldBounds As udtBoundingRegionType
        Dim udtCircleABounds As udtBoundingRegionType
        Dim udtCircleBBounds As udtBoundingRegionType
        Dim udtCircleCBounds As udtBoundingRegionType

        GetCircleBounds(m_CircleA_Loc, m_CircleA_Radius_Compute, udtCircleABounds)
        udtWorldBounds = udtCircleABounds

        GetCircleBounds(m_CircleB_Loc, m_CircleB_Radius_Compute, udtCircleBBounds)
        ExpandWorldBounds(udtWorldBounds, udtCircleBBounds)

        GetCircleBounds(m_CircleC_Loc, m_CircleC_Radius_Compute, udtCircleCBounds)
        ExpandWorldBounds(udtWorldBounds, udtCircleCBounds)


        ' Define the width of the compute world coordinates, from the outside of upper left-most circle to the outside of the bottom right-most circle
        Dim computeWorldWidth As Double = udtWorldBounds.LowerRight.X - udtWorldBounds.UpperLeft.X
        Dim computeWorldHeight As Double = udtWorldBounds.LowerRight.Y - udtWorldBounds.UpperLeft.Y


        'Factor that world coordinates need to be scaled to make height fit in control
        Dim heightScaleFactor As Double = (Me.Height) * Me.FillFactor / computeWorldHeight
        'Factor that world coordinates need to be scaled to make width fit in control
        Dim widthScaleFactor As Double = Me.Width * Me.FillFactor / computeWorldWidth

        'Factor that world coordinates need to be scaled to make both width and height fit
        Dim scaleFactor As Double = Math.Min(heightScaleFactor, widthScaleFactor)

        Dim circleAWidth As Double = (2 * Me.m_CircleA_Radius_Compute) * scaleFactor
        Dim circleBWidth As Double = (2 * Me.m_CircleB_Radius_Compute) * scaleFactor
        Dim circleCWidth As Double = (2 * Me.m_CircleC_Radius_Compute) * scaleFactor

        Dim screenWidth As Double = computeWorldWidth * scaleFactor
        Dim screenHeight As Double = computeWorldHeight * scaleFactor

        ' The following value is used if the left edge of CircleC becomes further left than the left edge of CircleA
        Dim dblXDirectionOffset As Double
        If udtCircleCBounds.UpperLeft.X < udtCircleABounds.UpperLeft.X Then
            dblXDirectionOffset = (udtCircleABounds.UpperLeft.X - udtCircleCBounds.UpperLeft.X) * scaleFactor
        Else
            dblXDirectionOffset = 0
        End If

        ' The following value is used if the top edge of CircleC becomes further up than the top edge of CircleA
        Dim dblYDirectionOffset As Double
        If udtCircleCBounds.UpperLeft.Y < udtCircleABounds.UpperLeft.Y Then
            dblYDirectionOffset = (udtCircleABounds.UpperLeft.Y - udtCircleCBounds.UpperLeft.Y) * scaleFactor
        Else
            dblYDirectionOffset = 0
        End If

        'The amount that the left edge of circleA has to be indented from the edge of the control.
        Dim xInset As Double = (Me.Width - screenWidth) / 2 + dblXDirectionOffset
        Dim yInset As Double = (Me.Height - screenHeight) / 2 + dblYDirectionOffset

        'Have to compute left and top coordinates for drawing circles
        Dim circleALeft As Double
        Dim circleBLeft As Double
        Dim circleCLeft As Double
        Dim circleATop As Double
        Dim circleBTop As Double
        Dim circleCTop As Double

        circleALeft = xInset
        circleBLeft = circleALeft + (Me.m_CircleA_Radius_Compute + m_CircleB_Loc.X - Me.m_CircleB_Radius_Compute) * scaleFactor
        circleCLeft = circleALeft + (Me.m_CircleA_Radius_Compute + m_CircleC_Loc.X - Me.m_CircleC_Radius_Compute) * scaleFactor

        If m_CircleA_Radius_Compute >= m_CircleB_Radius_Compute Then
            circleATop = yInset
            circleBTop = circleATop + (Me.m_CircleA_Radius_Compute + m_CircleB_Loc.Y - Me.m_CircleB_Radius_Compute) * scaleFactor
        Else
            circleBTop = yInset
            circleATop = circleBTop + (Me.m_CircleB_Radius_Compute + m_CircleB_Loc.Y - Me.m_CircleA_Radius_Compute) * scaleFactor
        End If

        circleCTop = circleATop + (Me.m_CircleA_Radius_Compute + m_CircleC_Loc.Y - Me.m_CircleC_Radius_Compute) * scaleFactor

        ' Update the ScreenLoc member variables (used when creating an SVG file)
        m_CircleA_ScreenRadius = circleAWidth / 2
        m_CircleA_ScreenLoc.X = circleALeft + m_CircleA_ScreenRadius
        m_CircleA_ScreenLoc.Y = circleATop + m_CircleA_ScreenRadius

        m_CircleB_ScreenRadius = circleBWidth / 2
        m_CircleB_ScreenLoc.X = circleBLeft + m_CircleB_ScreenRadius
        m_CircleB_ScreenLoc.Y = circleBTop + m_CircleB_ScreenRadius

        m_CircleC_ScreenRadius = circleCWidth / 2
        m_CircleC_ScreenLoc.X = circleCLeft + m_CircleC_ScreenRadius
        m_CircleC_ScreenLoc.Y = circleCTop + m_CircleC_ScreenRadius


        'Rectangles around circleA and circleB
        Dim circleABoundingBox As RectangleF = New RectangleF(CSng(circleALeft), CSng(circleATop), CSng(circleAWidth), CSng(circleAWidth))
        Dim circleBBoundingBox As RectangleF = New RectangleF(CSng(circleBLeft), CSng(circleBTop), CSng(circleBWidth), CSng(circleBWidth))
        Dim circleCBoundingBox As RectangleF = New RectangleF(CSng(circleCLeft), CSng(circleCTop), CSng(circleCWidth), CSng(circleCWidth))

        'Console.WriteLine("Computing Screen Coordinates")

        'screen coordinates don't exist as class level variables, they are inherently part
        'of the DrawingPath property of VennDiagramAreaInfo
        Me.m_circleA.DrawingPath.Reset()
        Me.m_circleA.DrawingPath.AddEllipse(circleABoundingBox)

        Me.m_circleB.DrawingPath.Reset()
        Me.m_circleB.DrawingPath.AddEllipse(circleBBoundingBox)

        Me.m_circleC.DrawingPath.Reset()
        Me.m_circleC.DrawingPath.AddEllipse(circleCBoundingBox)

        If Me.PaintSolidColorCircles Then
            ComputeScreenOverlapCoordinates(circleABoundingBox, circleBBoundingBox, circleCBoundingBox, m_DrawOverlapRegionsOnError)
        End If

        'Console.WriteLine("calculated screen coordinates")
        Me.m_ScreenCoordinatesValid = True
    End Sub

    Private Sub ComputeScreenOverlapCoordinates(ByVal circleABoundingBox As RectangleF, ByVal circleBBoundingBox As RectangleF, ByVal circleCBoundingBox As RectangleF, ByVal blnContinueOnError As Boolean)
        Dim sngAngleStart As Single
        Dim sngAngleStartB As Single
        Dim sngSweep As Single
        Dim sngSweepB As Single
        Dim blnInvalidSweepComputed As Boolean
        blnInvalidSweepComputed = False

        ' Create overlap region by taking an arc from each circle
        ' Circles A & B
        DrawOverlapRegion(Me.m_overlapAB.DrawingPath, circleABoundingBox, circleBBoundingBox, Me.mCirclesAB.OverlapABAlpha, Me.mCirclesAB.OverlapABBeta, 0, 0)

        ' Circles B & C
        DrawOverlapRegion(Me.m_overlapBC.DrawingPath, circleBBoundingBox, circleCBoundingBox, Me.mCirclesBC, mOverlapBCArcAdj)

        ' Circles A & C
        DrawOverlapRegion(Me.m_overlapAC.DrawingPath, circleABoundingBox, circleCBoundingBox, Me.mCirclesAC, mOverlapACArcAdj)


        ' The ABC overlap region (composed of 3 arcs)
        m_overlapABC.DrawingPath.Reset()

        ' The first arc (goes along circle A)
        sngAngleStart = CSng(-Me.mCirclesAC.OverlapABAlpha / 2 + mOverlapACArcAdj.AlphaAdd)

        sngAngleStartB = CSng(-Me.mCirclesAB.OverlapABAlpha / 2 + 0)
        sngSweepB = CSng(Me.mCirclesAB.OverlapABAlpha)

        sngSweep = sngSweepB - (sngAngleStart - sngAngleStartB)
        If sngSweep < 0 Then
            ' Invalid sweep; skip this arc
            blnInvalidSweepComputed = True
        Else
            m_overlapABC.DrawingPath.AddArc(circleABoundingBox, sngAngleStart, sngSweep)
        End If

        If Not blnInvalidSweepComputed OrElse blnContinueOnError Then
            ' The second arc (goes along circle B)
            sngAngleStart = CSng(180 - Me.mCirclesAB.OverlapABBeta / 2)

            sngAngleStartB = CSng(-Me.mCirclesBC.OverlapABAlpha / 2 + mOverlapBCArcAdj.AlphaAdd)
            sngSweepB = CSng(Me.mCirclesBC.OverlapABAlpha)

            sngSweep = sngSweepB - (sngAngleStart - sngAngleStartB)
            If sngSweep < 0 Then
                ' Invalid sweep; skip this arc
                blnInvalidSweepComputed = True
            Else
                m_overlapABC.DrawingPath.AddArc(circleBBoundingBox, sngAngleStart, sngSweep)
            End If
        End If

        If Not blnInvalidSweepComputed OrElse blnContinueOnError Then
            ' The third arc (goes along circle C)
            sngAngleStart = CSng(180 - Me.mCirclesBC.OverlapABBeta / 2 + mOverlapBCArcAdj.BetaAdd)

            sngAngleStartB = CSng(180 - Me.mCirclesAC.OverlapABBeta / 2 + mOverlapACArcAdj.BetaAdd)
            sngSweepB = CSng(Me.mCirclesAC.OverlapABBeta)

            sngSweep = sngSweepB - (sngAngleStart - sngAngleStartB)
            If sngSweep < 0 Then
                ' Invalid sweep; skip this arc
                blnInvalidSweepComputed = True
            Else
                m_overlapABC.DrawingPath.AddArc(circleCBoundingBox, sngAngleStart, sngSweep)
            End If
        End If
    End Sub

    'The first step in drawing a venn diagram, compute world coordinates with the
    'following properties.
    '   Position of circleA is (0, 0)
    '   Position of circleB is (x, 0)
    '   Position of circleC is (cx, cy) where cx is between 0 and x and cy is a positive number
    '   Areas of all shapes are equal to the sizes set by the user.
    Public Overrides Sub ComputeWorldCoordinates()
        ' Use TwoCircleVennDiagram classes to compute the overlap of A with B, A with C, and B with C, then
        ' determine the location of circle C by triangulating the three circles

        Dim dblAngleA As Double
        Dim dblSidea As Double
        Dim dblSideb As Double
        Dim dblSidec As Double

        Dim dblBCScalingFactor As Double
        Dim dblACScalingFactor As Double

        Dim udtCircleC_LocAlt As udtPointXYType
        Dim dblCircleC_Radius_ComputeAlt As Double

        Me.m_ComputeWorldCoordinatesValid = False

        If mCirclesAB Is Nothing Then
            mCirclesAB = New TwoCircleVennDiagram
            mCirclesBC = New TwoCircleVennDiagram
            mCirclesAC = New TwoCircleVennDiagram
        End If

        ' Check for conditions that can not be painted
        If (Me.m_circleA.Size <= 0 AndAlso Me.m_circleB.Size <= 0 AndAlso Me.m_circleC.Size <= 0) Then
            Me.SizesReasonable = False
            Exit Sub
        End If

        If (Me.m_overlapAB.Size > Math.Min(Me.m_circleA.Size, Me.m_circleB.Size)) Then
            Throw New ApplicationException("AB Overlap greater than Circle A or Circle B")
        End If

        If (Me.m_overlapBC.Size > Math.Min(Me.m_circleB.Size, Me.m_circleC.Size)) Then
            Throw New ApplicationException("BC Overlap greater than Circle B or Circle C")
        End If

        If (Me.m_overlapAC.Size > Math.Min(Me.m_circleA.Size, Me.m_circleC.Size)) Then
            Throw New ApplicationException("AC Overlap greater than Circle A or Circle C")
        End If

        'When size for one circle is equal to 0 this algorithm fails to work.  Solution
        'is to just make the size really small.  This causes it to be unnoticeable 
        'on the printed control.
        If (Me.CircleASize <= 0 AndAlso Me.CircleBSize > 0) Then Me.CircleASize = 0.0000000001
        If (Me.CircleASize > 0 AndAlso Me.CircleBSize <= 0) Then Me.CircleBSize = 0.0000000001

        If (Me.CircleBSize <= 0 AndAlso Me.CircleCSize > 0) Then Me.CircleBSize = 0.0000000001
        If (Me.CircleBSize > 0 AndAlso Me.CircleCSize <= 0) Then Me.CircleCSize = 0.0000000001

        If (Me.CircleASize <= 0 AndAlso Me.CircleCSize > 0) Then Me.CircleASize = 0.0000000001
        If (Me.CircleASize > 0 AndAlso Me.CircleCSize <= 0) Then Me.CircleCSize = 0.0000000001

        'Sizes are all reasonable, so move ahead with the calculation
        Me.SizesReasonable = True


        ' Compute the pairwise two circle overlaps
        ' Circles A & B
        With mCirclesAB
            .CircleASize = Me.CircleASize
            .CircleBSize = Me.CircleBSize
            .OverlapABSize = Me.OverlapABSize
            .ComputeWorldCoordinates()

            ' Copy the center and radius for Circle A
            m_CircleA_Loc = .CoordCircleA                       ' Always 0,0
            m_CircleA_Radius_Compute = .CoordCircleARadius

            ' Copy the X,Y coordinates and radius for Circle B; note that m_CircleB_Loc.Y will be 0 for now
            m_CircleB_Loc = .CoordCircleB
            m_CircleB_Radius_Compute = .CoordCircleBRadius
        End With

        ' Circles B & C
        With mCirclesAC
            .CircleASize = Me.CircleASize
            .CircleBSize = Me.CircleCSize
            .OverlapABSize = Me.OverlapACSize
            .ComputeWorldCoordinates()

            ' Copy the X,Y coordinates and radius for circle C (relative to Circle A)
            m_CircleC_Loc = .CoordCircleB
            m_CircleC_Radius_Compute = .CoordCircleBRadius

            If .CoordCircleARadius > 0 Then
                ' Scale Circle C's dimensions based on Circle A's radius in mCirclesAC vs. mCirclesAB
                dblACScalingFactor = mCirclesAB.CoordCircleARadius / .CoordCircleARadius
            Else
                dblACScalingFactor = 1
            End If

            m_CircleC_Loc.X *= dblACScalingFactor
            m_CircleC_Radius_Compute *= dblACScalingFactor
        End With

        ' Circles A & C
        With mCirclesBC
            .CircleASize = Me.CircleBSize
            .CircleBSize = Me.CircleCSize
            .OverlapABSize = Me.OverlapBCSize
            .ComputeWorldCoordinates()

            ' Copy the center and radius for circle C (relative to Circle B)
            udtCircleC_LocAlt = .CoordCircleB
            dblCircleC_Radius_ComputeAlt = .CoordCircleBRadius

            If .CoordCircleARadius > 0 Then
                ' Scale Circle C's dimensions based on Circle B's radius in mCirclesBC vs. mCirclesAB
                dblBCScalingFactor = mCirclesAB.CoordCircleBRadius / .CoordCircleARadius
            Else
                dblBCScalingFactor = 1
            End If


            udtCircleC_LocAlt.X *= dblBCScalingFactor
            dblCircleC_Radius_ComputeAlt *= dblBCScalingFactor
        End With

        ' Determine the location of Circle C
        ' We know the distance between the centers of each of the circles (A, B, and C)
        ' Thus, we know the length of three sides of the triangle between the circle centers 
        '  and we need to determine the angles, A, B, and C
        '
        '
        '                 C
        '                 /\
        'SideAC = side_b /   \ SideBC = side_a
        '               /      \
        '              /         \
        '            A-------------B
        '            SideAB = side_c

        ' Define the sides to the triangle, as illustrated above
        dblSidea = Math.Abs(dblBCScalingFactor * (mCirclesBC.CoordCircleB.X - mCirclesBC.CoordCircleA.X))
        dblSideb = Math.Abs(dblACScalingFactor * (mCirclesAC.CoordCircleB.X - mCirclesAC.CoordCircleA.X))
        dblSidec = Math.Abs(mCirclesAB.CoordCircleB.X - mCirclesAB.CoordCircleA.X)

        ' Compute Angle A, which is between SideAC and SideAB (the angle will be in radians)
        dblAngleA = AngleAFromThreeSides(dblSidea, dblSideb, dblSidec)

        ' Compute the x and y coordinates for circle C; store in m_CircleC_Loc
        m_CircleC_Loc.X = dblSideb * Math.Cos(dblAngleA)
        m_CircleC_Loc.Y = dblSideb * Math.Sin(dblAngleA)


        ' Compute the overlap angle adjustment value for the BC and AC overlaps
        Dim dblAngle As Double

        dblAngle = ComputeRotationAngle(m_CircleB_Loc, m_CircleC_Loc)
        mOverlapBCArcAdj.AlphaAdd = dblAngle
        mOverlapBCArcAdj.BetaAdd = dblAngle

        dblAngle = ComputeRotationAngle(m_CircleA_Loc, m_CircleC_Loc)
        mOverlapACArcAdj.AlphaAdd = dblAngle
        mOverlapACArcAdj.BetaAdd = dblAngle


        ' Compute the 

        'Console.WriteLine("calculated world coordinates")
        Me.m_ComputeWorldCoordinatesValid = True

    End Sub

    Protected Overloads Sub DrawOverlapRegion(ByRef objDrawingPath As System.Drawing.Drawing2D.GraphicsPath, ByVal BoundingBoxA As RectangleF, ByVal BoundingBoxB As RectangleF, ByRef objVennDiagram As TwoCircleVennDiagram, ByVal udtArcAngleAdjust As udtArcAngleAdjustmentType)
        DrawOverlapRegion(objDrawingPath, BoundingBoxA, BoundingBoxB, objVennDiagram.OverlapABAlpha, objVennDiagram.OverlapABBeta, udtArcAngleAdjust.AlphaAdd, udtArcAngleAdjust.BetaAdd)
    End Sub

    Protected Sub ExpandWorldBounds(ByRef udtWorldBounds As udtBoundingRegionType, ByVal udtCompareBounds As udtBoundingRegionType)

        ' Update udtUpperLeft if necessary
        If udtCompareBounds.UpperLeft.X < udtWorldBounds.UpperLeft.X Then udtWorldBounds.UpperLeft.X = udtCompareBounds.UpperLeft.X
        If udtCompareBounds.UpperLeft.Y < udtWorldBounds.UpperLeft.Y Then udtWorldBounds.UpperLeft.Y = udtCompareBounds.UpperLeft.Y

        ' Update udtLowerRight if necessary
        If udtCompareBounds.LowerRight.X > udtWorldBounds.LowerRight.X Then udtWorldBounds.LowerRight.X = udtCompareBounds.LowerRight.X
        If udtCompareBounds.LowerRight.Y > udtWorldBounds.LowerRight.Y Then udtWorldBounds.LowerRight.Y = udtCompareBounds.LowerRight.Y

    End Sub

    Protected Sub GetCircleBounds(ByVal udtCircleLoc As udtPointXYType, ByVal dblRadius As Double, ByRef udtBounds As udtBoundingRegionType)

        udtBounds.UpperLeft.X = udtCircleLoc.X - dblRadius
        udtBounds.UpperLeft.Y = udtCircleLoc.Y - dblRadius

        udtBounds.LowerRight.X = udtCircleLoc.X + dblRadius
        udtBounds.LowerRight.Y = udtCircleLoc.Y + dblRadius

    End Sub

    ''' <summary>
    ''' Constructs the SVG file text for the overlapping circles
    ''' </summary>
    ''' <param name="blnFillCircles">True to color the circles, false to leave empty</param>
    ''' <param name="sngOpacity">Value between 0 and 1 representing transparency; 1 means fully opaque</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function GetSVG(ByVal blnFillCircles As Boolean, ByVal sngOpacity As Single) As String
        Dim strSVG As System.Text.StringBuilder
        Dim intStrokeWidthPixels As Integer = 1


        ' Create the SVG text
        strSVG = New System.Text.StringBuilder

        strSVG.AppendLine("<?xml version=""1.0"" standalone=""no""?>")
        strSVG.AppendLine("<!DOCTYPE svg PUBLIC ""-//W3C//DTD SVG 1.1//EN"" ""http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"">")
        strSVG.AppendLine()
        strSVG.AppendLine("<svg width=""100%"" height=""100%"" version=""1.1"" xmlns=""http://www.w3.org/2000/svg"">")
        strSVG.AppendLine()

        strSVG.AppendLine("<g id=""OverlappingCircles"">")

        strSVG.AppendLine(MyBase.GetSVGCirclesAB(intStrokeWidthPixels, blnFillCircles, sngOpacity))

        strSVG.AppendLine(ControlChars.Tab & "<desc>Circle C</desc>")
        strSVG.AppendLine(GetSVGCircleText(m_CircleC_ScreenLoc, m_CircleC_ScreenRadius, blnFillCircles, CircleCColor, sngOpacity, intStrokeWidthPixels))

        strSVG.AppendLine("</g>")

        'strSVG.AppendLine("<g id=""Text"">")
        'strSVG.AppendLine("	<text x=""240"" y=""50"" font-family=""Verdana"" font-size=""20"" fill=""black"" text-anchor=""middle"">220</text>")
        'strSVG.AppendLine("	<text x=""350"" y=""75"" font-family=""Verdana"" font-size=""20"" fill=""black"" text-anchor=""middle"">75</text>")
        'strSVG.AppendLine("	<text x=""400"" y=""50"" font-family=""Verdana"" font-size=""20"" fill=""black"" text-anchor=""middle"">150</text>")
        'strSVG.AppendLine("</g>")
        strSVG.AppendLine()
        strSVG.AppendLine("</svg>")

        Return strSVG.ToString

    End Function

    Protected Overrides Sub InitializeVariables()
        MyBase.InitializeVariables()

        m_ZoomPct = 100
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        Dim g As Graphics = e.Graphics
        Dim sngScalingFactor As Single
        Dim sngXOffsetFactor As Single
        Dim sngYOffsetFactor As Single

        'Begin a container so smoothing mode change does not impact others drawing on same graphics
        Dim container As Drawing2D.GraphicsContainer = g.BeginContainer()

        g.SmoothingMode = Me.SmoothingMode

        sngScalingFactor = m_ZoomPct / 100.0F
        sngXOffsetFactor = m_XOffset / 100.0F
        sngYOffsetFactor = m_YOffset / 100.0F

        If sngScalingFactor > 0 AndAlso sngScalingFactor <> 1 Then
            g.TranslateTransform(Me.Width * ((1 - sngScalingFactor) / 2 + sngXOffsetFactor), Me.Height * ((1 - sngScalingFactor) / 2 + sngYOffsetFactor))
            g.ScaleTransform(sngScalingFactor, sngScalingFactor)
        ElseIf sngXOffsetFactor <> 0 Or sngYOffsetFactor <> 0 Then
            g.TranslateTransform(Me.Width * sngXOffsetFactor, Me.Height * sngYOffsetFactor)
        End If

        If m_Rotation <> 0 Then
            ' The following trick will effectively rotate the entire image across the center
            g.TranslateTransform(Me.Width / 2.0F, Me.Height / 2.0F)
            g.RotateTransform(m_Rotation)
            g.TranslateTransform(-Me.Width / 2.0F, -Me.Height / 2.0F)
        End If

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


        'draw circle and overlap, starting with circleA, circleB, and circleC, then with overlap so that it draws over them
        Me.PaintVennDiagramInfo(Me.m_circleA, g)
        Me.PaintVennDiagramInfo(Me.m_circleB, g)
        Me.PaintVennDiagramInfo(Me.m_circleC, g)
        Me.PaintVennDiagramInfo(Me.m_overlapAB, g)
        Me.PaintVennDiagramInfo(Me.m_overlapBC, g)
        Me.PaintVennDiagramInfo(Me.m_overlapAC, g)
        Me.PaintVennDiagramInfo(Me.m_overlapABC, g)

        If m_Rotation <> 0 Then
            g.RotateTransform(-m_Rotation)
        End If

        'deal with labels - TODO
        'g.DrawString(Me.m_circleA.Label, Me.Font, New SolidBrush(Me.ForeColor), 0, 50)
        MyBase.OnPaint(e)
    End Sub

    Protected Shared Sub RotatePoints(ByRef udtLoc As udtPointXYType, ByVal dblAngleDegrees As Double)
        RotatePoints(udtLoc.X, udtLoc.Y, dblAngleDegrees)
    End Sub

    Public Shared Sub RotatePoints(ByRef dblX As Double, ByRef dblY As Double, ByVal dblAngleDegrees As Double)
        Dim dblXNew As Double
        Dim dblYNew As Double
        Dim dblAngleRadians As Double

        dblAngleRadians = ConvertDegreesToRadians(dblAngleDegrees)

        ' Compute the new coordinates
        dblXNew = dblX * Math.Cos(dblAngleRadians) - dblY * Math.Sin(dblAngleRadians)
        dblYNew = dblX * Math.Sin(dblAngleRadians) + dblY * Math.Cos(dblAngleRadians)

        ' Update dblX and dblY
        dblX = dblXNew
        dblY = dblYNew

    End Sub
End Class
