Option Strict On

Public Class clsMatrixSolver
    ' From http://www.vb-helper.com/howto_net_gaussian_elimination.html

    ' This class can be used to solve a system of equations using Gaussian Elimination
    ' The system of equations is of the form: 
    '     A1*x1 + B1*x2 + ... + N1*xn = c1
    '     A2*x1 + B2*x2 + ... + N2*xn = c2
    '     ...
    '     Am*x1 + Bm*x2 + ... + Nm*xn = cm = 0
    '
    ' For example: 
    '     9 * x1 + 4 * x2 = 7
    '     4 * x1 + 3 * x2 = 8
    '
    ' You can write these equations as a matrix multiplied by a vector of variables x1, x2, ..., xn, equals a vector of constants c1, c2, ..., cn. 
    '          Matrix A        Matrix X   Matrix B
    '     | A1  B1  ...  N1 |   | x1 |     | c1 |
    '     | A2  B2  ...  N2 |   | x2 |     | c2 |
    '     |         ...     | * |... |  =  |... |
    '     | Am  Bm  ...  Nm |   | xm |     | cm |
    '
    ' To use SolveEquations, provide MatrixA in dblCoefMatrix and MatrixB in dblConstants; 
    ' the solved values for the variables will be returned in dblXValues

    Public Function SolveEquations(ByRef dblCoefMatrix(,) As Double, ByRef dblConstants() As Double, ByRef dblXValues() As Double, ByRef strMessage As String) As Boolean

        ' This function uses dblCoefMatrix and dblConstants to structure the data into an augmented matrix 
        '  that includes the data matrix, the constants, and one extra column to hold the variables' final values

        Dim intRowCount As Integer
        Dim intColCount As Integer
        Dim blnSuccess As Boolean

        Dim dblAugmentedMatrix(,) As Double
        Dim x As Integer, y As Integer

        If dblCoefMatrix Is Nothing Then
            strMessage = "dblCoefMatrix cannot be Nothing"
            Return False
        ElseIf dblConstants Is Nothing Then
            strMessage = "dblConstants cannot be Nothing"
            Return False
        End If

        intRowCount = UBound(dblCoefMatrix, 1) + 1
        intColCount = UBound(dblCoefMatrix, 2) + 1

        If intRowCount < 2 Then
            strMessage = "Must have at least two rows"
            Return False
        End If

        If intColCount < 2 Then
            strMessage = "Must have at least two columns"
            Return False
        End If

        If UBound(dblConstants) + 1 <> intRowCount Then
            strMessage = "dblConstants must have the same number of rows as dblCoefMatrix"
            Return False
        End If

        ' Clear dblXValues
        If dblXValues Is Nothing OrElse UBound(dblXValues) < intRowCount - 1 Then
            ReDim dblXValues(intRowCount - 1)
        Else
            Array.Clear(dblXValues, 0, UBound(dblXValues) + 1)
        End If

        ' Populate dblAugmentedMatrix
        ReDim dblAugmentedMatrix(intRowCount - 1, intColCount + 1)

        For x = 0 To intRowCount - 1
            For y = 0 To intColCount - 1
                dblAugmentedMatrix(x, y) = dblCoefMatrix(x, y)
            Next y
            dblAugmentedMatrix(x, intColCount) = dblConstants(x)
        Next x

        blnSuccess = SolveEquationsSingleMatrix(dblAugmentedMatrix, strMessage)

        If blnSuccess Then
            ' Populate dblXValues
            For x = 0 To intRowCount - 1
                dblXValues(x) = dblAugmentedMatrix(x, intColCount + 1)
            Next x
        End If

        Return blnSuccess

    End Function

    Public Function MultiplyMatrices(ByRef dblMatrix1(,) As Double, ByRef dblArray() As Double, ByRef dblResults() As Double) As Boolean
        ' This function assumes dblArray is a matrix with many rows, but just 1 column (and is represented by an array)
        ' Auto-determines the row and column count of dblMatrix1 using the rule that 
        ' the number of columns in the first matrix must equal the number of rows in the second matrix
        ' If dblMatrix1 or dblArray has extra rows or columns, then they're ignored

        Dim intRowCount1 As Integer
        Dim intColCount1 As Integer
        Dim intRowCount2 As Integer

        If dblMatrix1 Is Nothing OrElse dblArray Is Nothing Then
            Return False
        End If

        intRowCount1 = UBound(dblMatrix1, 1) + 1
        intColCount1 = UBound(dblMatrix1, 2) + 1
        intRowCount2 = UBound(dblArray) + 1

        If intColCount1 > intRowCount2 Then
            intColCount1 = intRowCount2
        End If

        Return MultiplyMatrices(dblMatrix1, dblArray, dblResults, intRowCount1, intColCount1)

    End Function

    Public Function MultiplyMatrices(ByRef dblMatrix1(,) As Double, ByRef dblMatrix2(,) As Double, ByRef dblResults(,) As Double) As Boolean
        ' Auto-determines the row and column count of dblMatrix1 using the rule that 
        ' the number of columns in the first matrix must equal the number of rows in the second matrix
        ' If dblMatrix1 or dblArray has extra rows or columns, then they're ignored

        Dim intRowCount1 As Integer
        Dim intColCount1 As Integer
        Dim intRowCount2 As Integer
        Dim intColCount2 As Integer

        If dblMatrix1 Is Nothing OrElse dblMatrix2 Is Nothing Then
            Return False
        End If

        intRowCount1 = UBound(dblMatrix1, 1) + 1
        intColCount1 = UBound(dblMatrix1, 2) + 1
        intRowCount2 = UBound(dblMatrix2, 1) + 1
        intColCount2 = UBound(dblMatrix2, 2) + 1

        If intColCount1 > intRowCount2 Then
            intColCount1 = intRowCount2
        End If

        Return MultiplyMatrices(dblMatrix1, dblMatrix2, dblResults, intRowCount1, intColCount1, intRowCount2, intColCount2)

    End Function


    Protected Function MultiplyMatrices(ByRef dblMatrix1(,) As Double, ByRef dblMatrix2() As Double, ByRef dblResults() As Double, _
                                        ByVal intRowCount1 As Integer, ByVal intColCount1 As Integer) As Boolean
        ' This function assumes dblMatrix2 is a matrix with many rows, but just 1 column (and is represented by an array)
        ' Thus, dblResults is returned as a 1D array 

        Dim x As Integer, y As Integer
        Dim dblTemp As Double

        If dblMatrix1 Is Nothing OrElse dblMatrix2 Is Nothing Then
            Return False
        End If

        If dblResults Is Nothing OrElse UBound(dblResults) < intRowCount1 - 1 Then
            ReDim dblResults(intRowCount1 - 1)
        Else
            Array.Clear(dblResults, 0, UBound(dblResults) + 1)
        End If

        ' Multiply the matrices and store in dblResults
        For x = 0 To intRowCount1 - 1
            dblTemp = 0
            For y = 0 To intColCount1 - 1
                dblTemp += dblMatrix1(x, y) * dblMatrix2(y)
            Next y
            dblResults(x) = dblTemp
        Next x

        Return True

    End Function

    Protected Function MultiplyMatrices(ByRef dblMatrix1(,) As Double, ByRef dblMatrix2(,) As Double, ByRef dblResults(,) As Double, _
                                        ByVal intRowCount1 As Integer, ByVal intColCount1 As Integer, _
                                        ByVal intRowCount2 As Integer, ByVal intColCount2 As Integer) As Boolean
        ' If dblMatrix1 has extra columns or dblMatrix2 has extra rows, then the extra information is ignored

        Dim x As Integer, y As Integer, z As Integer
        Dim dblTemp As Double

        If dblMatrix1 Is Nothing OrElse dblMatrix2 Is Nothing Then
            Return False
        End If

        If dblResults Is Nothing OrElse UBound(dblResults, 1) < intRowCount1 - 1 OrElse UBound(dblResults, 2) < intColCount2 - 1 Then
            ReDim dblResults(intRowCount1 - 1, intColCount2 - 1)
        Else
            For x = 0 To UBound(dblResults, 1)
                For y = 0 To UBound(dblResults, 2)
                    dblResults(x, y) = 0
                Next y
            Next x
        End If

        ' Multiply the matrices and store in dblResults
        For z = 0 To intColCount2 - 1
            For x = 0 To intRowCount1 - 1
                dblTemp = 0
                For y = 0 To intColCount1 - 1
                    dblTemp += dblMatrix1(x, y) * dblMatrix2(y, z)
                Next y
                dblResults(x, z) = dblTemp
            Next x
        Next z

        Return True

    End Function

    Protected Function GetTestMatrix(ByVal intRowCount As Integer, ByVal intColCount As Integer, ByVal blnUseFixed As Boolean, ByRef dblConstants() As Double) As Double(,)
        Const MAX_INTEGER As Integer = 10

        Dim objRand As New Random

        Dim x As Integer
        Dim y As Integer

        If intRowCount < 2 Then intRowCount = 2
        If intColCount < 2 Then intColCount = 2

        If blnUseFixed Then
            Dim dblCoefMatrix(,) As Double = {{2, 2, 7, 0, 0, 0}, _
                                          {2, 1, 2, 2, 2, 0}, _
                                          {3, 4, 5, 7, 3, 0}, _
                                          {6, 7, 1, 8, 9, 6}, _
                                          {7, 5, 4, 9, 1, 5}, _
                                          {1, 9, 5, 8, 1, 0}}

            ReDim dblConstants(UBound(dblCoefMatrix, 1))
            For x = 0 To UBound(dblConstants)
                dblConstants(x) = x + 1
            Next x

            Return dblCoefMatrix
        Else
            Dim dblCoefMatrix(,) As Double
            ReDim dblCoefMatrix(intRowCount - 1, intColCount - 1)
            ReDim dblConstants(intRowCount - 1)

            For x = 0 To intRowCount - 1
                For y = 0 To intColCount - 1
                    dblCoefMatrix(x, y) = objRand.Next(0, MAX_INTEGER)
                Next y
                dblConstants(x) = x + 1
            Next x

            Return dblCoefMatrix
        End If

    End Function

    Public Sub TestSolver()
        Dim dblCoefMatrix(,) As Double
		Dim dblConstants() As Double = Nothing
		Dim dblXValues() As Double = Nothing

        Dim blnSuccess As Boolean
		Dim strMessage As String = String.Empty

        Dim intRowCount As Integer = 6
        Dim intColCount As Integer = 6

        dblCoefMatrix = GetTestMatrix(intRowCount, intColCount, True, dblConstants)

        blnSuccess = SolveEquations(dblCoefMatrix, dblConstants, dblXValues, strMessage)

    End Sub

    Protected Function SolveEquationsSingleMatrix(ByRef dblAugmentedMatrix(,) As Double, ByRef strMessage As String) As Boolean

        ' To use this function, structure the data into an augmented matrix that includes the 
        ' original coefficients, the constants, and one extra column to hold the variables' final values
        '     | A1  B1  ...  N1  c1  x1 |
        '     | A2  B2  ...  N2  c2  x2 |
        '     |         ...           |
        '     | Am  Bm  ...  Nm  cm  xm |

        ' In other words, dblAugmentedMatrix must contain the constants plus one extra column
        '  Column num_cols + 1 will hold the variables' final values.
        '
        ' Note that This function uses the dimensions of dblAugmentedMatrix to determine the number of rows and columns

        Const TINY As Double = 0.00001
        Dim intRowCount As Integer
        Dim intColCount As Integer
        Dim dblTemp As Double
        Dim factor As Double
        Dim sOriginalMatrix(,) As Double
        Dim x As Integer, y As Integer
        Dim r As Integer, r2 As Integer
        Dim c As Integer

        Dim blnSuccess As Boolean

        strMessage = String.Empty
        blnSuccess = False

        ' Determine the number of rows and columns in dblAugmentedMatrix (the matrix must have 3 extra columns in addition to the coefficients)
        intRowCount = UBound(dblAugmentedMatrix, 1) + 1
        intColCount = UBound(dblAugmentedMatrix, 2) - 1

        If intRowCount < 2 Then
            strMessage = "dblAugmentedMatrix must have at least two rows"
            Return False
        End If

        If intColCount < intRowCount Then
            strMessage = "dblAugmentedMatrix must have at least as many columns as rows, and must also contain two additional columns, one for the constants and one to hold the results"
            Return False
        End If

        ' Copy the values from dblAugmentedMatrix to sOriginalMatrix
        ReDim sOriginalMatrix(intRowCount - 1, intColCount)
        For x = 0 To intRowCount - 1
            For y = 0 To intColCount
                sOriginalMatrix(x, y) = dblAugmentedMatrix(x, y)
            Next y
        Next x

        ' Display the initial array.
        'PrintArray(dblAugmentedMatrix)
        'PrintArray(sOriginalMatrix)

        ' Start solving.
        For r = 0 To intRowCount - 2
            ' Zero out all entries in column r after this row.
            ' See if this row has a non-zero entry in column r.
            If Math.Abs(dblAugmentedMatrix(r, r)) < TINY Then
                ' Not a non-zero value. Try to swap with a
                ' later row.
                For r2 = r + 1 To intRowCount - 1
                    If Math.Abs(dblAugmentedMatrix(r2, r)) > TINY Then
                        ' This row will work. Swap them.
                        For c = 0 To intColCount
                            dblTemp = dblAugmentedMatrix(r, c)
                            dblAugmentedMatrix(r, c) = dblAugmentedMatrix(r2, c)
                            dblAugmentedMatrix(r2, c) = dblTemp
                        Next c
                        Exit For
                    End If
                Next r2
            End If

            ' If this row has a non-zero entry in column r,
            ' skip this column.
            If Math.Abs(dblAugmentedMatrix(r, r)) > TINY Then
                ' Zero out this column in later rows.
                For r2 = r + 1 To intRowCount - 1
                    factor = -dblAugmentedMatrix(r2, r) / dblAugmentedMatrix(r, r)
                    For c = r To intColCount
                        dblAugmentedMatrix(r2, c) = dblAugmentedMatrix(r2, c) + factor * dblAugmentedMatrix(r, c)
                    Next c
                Next r2
            End If
        Next r

        ' Display the upper-triangular array.
        'PrintArray(dblAugmentedMatrix)

        ' See if we have a solution.
        If dblAugmentedMatrix(intRowCount - 1, intColCount - 1) = 0 Then
            ' We have no solution.
            strMessage = "No solution"
        Else
            ' Back solve.
            For r = intRowCount - 1 To 0 Step -1
                dblTemp = dblAugmentedMatrix(r, intColCount)
                For r2 = r + 1 To intRowCount - 1
                    dblTemp -= dblAugmentedMatrix(r, r2) * dblAugmentedMatrix(r2, intColCount + 1)
                Next r2
                dblAugmentedMatrix(r, intColCount + 1) = dblTemp / dblAugmentedMatrix(r, r)
            Next r

            ' Display the results.
            strMessage = " Solution found"
            For r = 0 To intRowCount - 1
                strMessage &= vbCrLf & "x" & r.ToString & " = " & dblAugmentedMatrix(r, intColCount + 1).ToString
            Next r

            'If True Then
            '    ' Verify the results.
            '    Dim dblXValues() As Double
            '    Dim dblResults() As Double

            '    ReDim dblXValues(intRowCount - 1)
            '    For x = 0 To intRowCount - 1
            '        dblXValues(x) = dblAugmentedMatrix(x, intColCount + 1)
            '    Next x

            '    MultiplyMatrices(sOriginalMatrix, dblXValues, dblResults)
            'End If

            'PrintArray(dblAugmentedMatrix)

            blnSuccess = True
        End If

        Return blnSuccess
    End Function

    Public Shared Sub PrintArray(ByRef dblArray() As Double)
        Dim intDataCount As Integer
        Dim x As Integer

        intDataCount = UBound(dblArray) + 1

        Console.WriteLine()
        Console.WriteLine("Contents of the " & intDataCount.ToString & " member array")

        For x = 0 To intDataCount - 1
            Console.WriteLine(Math.Round(dblArray(x), 3).ToString)
        Next x

    End Sub

    Public Shared Sub PrintArray(ByRef dblMatrix(,) As Double)
        Dim intRowCount As Integer, intColCount As Integer
        Dim x As Integer, y As Integer
        Dim strRow As String

        intRowCount = UBound(dblMatrix, 1) + 1
        intColCount = UBound(dblMatrix, 2) + 1

        Console.WriteLine()
        Console.WriteLine("Contents of the " & intRowCount.ToString & " x " & intColCount.ToString & " matrix")

        For x = 0 To intRowCount - 1
            strRow = "| "
            For y = 0 To intColCount - 1
                strRow &= Math.Round(dblMatrix(x, y), 3).ToString
                If y < intColCount - 1 Then
                    strRow &= " , "
                Else
                    strRow &= " |"
                End If
            Next y
            Console.WriteLine(strRow)
        Next x

    End Sub
End Class
