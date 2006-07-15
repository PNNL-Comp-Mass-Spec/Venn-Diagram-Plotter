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
Public Class ExperimentationDiagram
    Inherits System.Windows.Forms.UserControl


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region


    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim g As Graphics = e.Graphics
        Dim r As Rectangle = New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim b As Brush = New Drawing2D.LinearGradientBrush(r, Color.WhiteSmoke, Color.Violet, 0, False)

        MyBase.OnPaint(e)
        'g.SetClip(r, Drawing2D.CombineMode.Replace)
        'g.Clear(Color.White)
        'Console.WriteLine(r)
        'Console.WriteLine("Bounds " & g.ClipBounds.ToString)
        'g.DrawLine(System.Drawing.Pens.BlanchedAlmond, r.Left, r.Top, r.Right, r.Bottom)
        g.FillRectangle(b, New Rectangle(Me.Left + 10, Me.Top + 10, Me.Width - 20, Me.Height - 20))
        If (LineOn) Then
            g.DrawLine(System.Drawing.Pens.Plum, r.Left, r.Bottom, r.Right, r.Top)
        End If

    End Sub

    Private m_lineOn As Boolean = True

    Public Property LineOn() As Boolean
        Get
            Return m_lineOn
        End Get
        Set(ByVal Value As Boolean)
            m_lineOn = Value
        End Set
    End Property
End Class
