Option Strict On

Imports System.Windows.Forms
Imports System.Drawing

'Prints controls to Graphics objects in a way that vaguely resembles how the control appears onscreen
'(but it's actually quite good)
'Currently handles
'   IPrintableControlGroup
'   IPrintableControl
'   Panel
'   Label (Ignores vertical part of .TextAlign property, but does handle horizontal part) - Now handles both.
'   GroupBox (Slightly differing label alignment, also unknown (and probably undesirable behavior)
'       when used with long labels.  This could be corrected with proper use of the StringFormat
'       class and the Graphics.MeasureString methods.)
'   Button (Coloring of border is not 100% the same, due to limits of ControlPaint class, which always uses
'       default form gray.)
'   Borders (on labels and panels) are always drawn as a solid line, because the DrawBorder3D method
'       does not use the current transform of the Graphics object
'
'Does not handle any pictures used as control backgrounds.
'Handling of Enabled property of controls is non-existent

'Now using GraphicsContainer's instead of setting then reversing clip/smoothing mode, etc.

'The basic idea for this class was found at http://www.c-sharpcorner.com/Code/2003/March/FormPrinting.asp

'The behavior of this class when overlapping controls are encountered is not always consistent with
'the onscreen display.  So don't overlap controls (controls that contain others, such as group boxes
'and panels are ok).

'The other way of printing controls is to draw them to screen, then use some dll calls to copy what
'is onscreen.  This way is easier to understand and the controls do not have to actually be drawn to
'the screen.

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

Public Class ControlPrinter
    'Control is drawn at 0, 0 to width, height
    Public Shared Sub DrawControl(c As Control, g As Graphics, drawBackground As Boolean)
        'Console.WriteLine("Printing control: " & c.GetType().ToString & " """ & c.Text & """")
        Dim controlBounds = New Rectangle(0, 0, c.Width, c.Height)
        Dim controlBoundsF = New RectangleF(0, 0, c.Width, c.Height)
        Dim container As Drawing2D.GraphicsContainer = g.BeginContainer()

        'set the current clipping so that even we attempt to draw something outside the
        'boundaries of the control, it will not actually be drawn on the output
        g.SetClip(controlBounds, Drawing2D.CombineMode.Intersect)


        If (drawBackground) Then
            g.FillRectangle(New SolidBrush(c.BackColor), New RectangleF(0, 0, c.Width, c.Height))
        End If

        If (TypeOf c Is IPrintableControl) Then
            'it knows how to draw itself - too bad everything isn't this intelligent
            DirectCast(c, IPrintableControl).DrawOnGraphics(g, drawBackground)

        ElseIf (TypeOf c Is Button) Then
            Dim b = DirectCast(c, Button)
            Dim format = New StringFormat
            Dim oldClip As Region = g.Clip
            Dim innerButtonBounds = New RectangleF(2, 2, b.Width - 4, b.Height - 4)
            Dim buttonContainer As Drawing2D.GraphicsContainer = g.BeginContainer

            'since ControlPaint always draws a gray button, set a clip that excludes the inner portion
            'so that whatever background is currently there remains.
            g.SetClip(innerButtonBounds, Drawing2D.CombineMode.Exclude)
            ControlPaint.DrawButton(g, controlBounds, ButtonState.Normal)

            'set clip back so text can be drawn
            g.EndContainer(buttonContainer)

            'This differs from what buttons actually do, but I like the look better
            format.Trimming = StringTrimming.EllipsisWord

            'get text alignment
            format.LineAlignment = GetVerticalStringAlignment(b.TextAlign)
            format.Alignment = GetHorizontalStringAlignment(b.TextAlign)

            'controlBounds are too wide, restrict by the same amount that background was
            'g.DrawString(c.Text, c.Font, New SolidBrush(c.ForeColor), controlBoundsF, format)

            g.DrawString(b.Text, b.Font, New SolidBrush(GetFontColor(c)), innerButtonBounds, format)

        ElseIf (TypeOf c Is Label) Then
            Dim l = DirectCast(c, Label)
            Dim format = New StringFormat
            Dim fontColor As Color = l.ForeColor

            'This differs from what buttons actually do, but I like the look better
            format.Trimming = StringTrimming.EllipsisWord

            'set horizontal alignment to same as the label
            format.LineAlignment = GetVerticalStringAlignment(l.TextAlign)
            format.Alignment = GetHorizontalStringAlignment(l.TextAlign)

            'draw border if it has one
            If (l.BorderStyle = BorderStyle.FixedSingle Or l.BorderStyle = BorderStyle.Fixed3D) Then
                'see GroupBox drawing for why the DrawBorder method is always used
                ControlPaint.DrawBorder(g, controlBounds, l.ForeColor, ButtonBorderStyle.Solid)
                'ControlPaint.DrawBorder3D(g, controlBounds, Border3DStyle.Sunken)
            End If

            g.DrawString(c.Text, c.Font, New SolidBrush(GetFontColor(c)), controlBoundsF, format)

        ElseIf (TypeOf c Is Panel) Then
            Dim p = DirectCast(c, Panel)
            If (p.BorderStyle = BorderStyle.FixedSingle Or p.BorderStyle = BorderStyle.Fixed3D) Then
                ControlPaint.DrawBorder(g, controlBounds, c.ForeColor, ButtonBorderStyle.Solid)
            End If
        ElseIf (TypeOf c Is GroupBox) Then
            Dim stringBounds As SizeF = g.MeasureString(c.Text, c.Font)
            Dim stringClip = New Rectangle(5, 0, CInt(stringBounds.Width), CInt(stringBounds.Height))
            Dim oldClip As Region = g.Clip
            Dim groupBoxContainer As Drawing2D.GraphicsContainer = g.BeginContainer()

            'remove the area where the label will go from the part of the graphics that can be drawn to
            'so that the border will not be drawn in that part
            g.SetClip(stringClip, Drawing2D.CombineMode.Exclude)

            'draw a border slightly indented from the true edges of the control
            ControlPaint.DrawBorder(g, New Rectangle(1, CInt(stringBounds.Height * 0.4), c.Width - 2, CInt(c.Height - (stringBounds.Height * 0.4) - 1)), c.ForeColor, ButtonBorderStyle.Solid)

            'can not draw true windows form style border using the DrawBorder3D method
            'because it ignores current transform of the graphics, instead drawing at the absolute
            'origin instead of the current transform origin.
            'ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, c.Width, c.Height), Border3DStyle.Flat)

            'g.FillRectangle(New SolidBrush(c.BackColor), 5, 0, stringBounds.Width, stringBounds.Height)
            'set old clip back so that label cn be drawn
            g.EndContainer(groupBoxContainer)
            g.DrawString(c.Text, c.Font, New SolidBrush(GetFontColor(c)), 5, 0)
        End If

        If (ControlHasPaintableChildren(c)) Then
            DrawChildControls(c, g, drawBackground)
        End If
        'return to previous clip
        g.EndContainer(container)
    End Sub

    Private Shared Function GetFontColor(c As Control) As Color
        If (Not c.Enabled) Then
            Return Color.Gray
        End If
        Return c.ForeColor
    End Function

    Private Shared Function ControlHasPaintableChildren(c As Control) As Boolean
        Return TypeOf c Is IPrintableControlContainer Or TypeOf c Is Panel Or TypeOf c Is GroupBox
    End Function

    'Use StringFormat class instead
    'Private Shared Function GetStringTopLeft(ByVal label As Label, ByVal g As Graphics) As PointF
    '    Dim stringBounds As SizeF = g.MeasureString(label.Text, label.Font)
    '    Dim left As Single
    '    Dim top As Single

    '    Select Case label.TextAlign
    '        Case ContentAlignment.BottomCenter
    '            left = (label.Width - stringBounds.Width) / 2
    '            top = label.Height - stringBounds.Height
    '        Case ContentAlignment.BottomLeft
    '            left = 0
    '            top = label.Height - stringBounds.Height
    '        Case ContentAlignment.BottomRight
    '            left = (label.Width - stringBounds.Width)
    '            top = label.Height - stringBounds.Height
    '        Case ContentAlignment.MiddleCenter
    '            left = (label.Width - stringBounds.Width) / 2
    '            top = (label.Height - stringBounds.Height) / 2
    '        Case ContentAlignment.MiddleLeft
    '            left = 0
    '            top = (label.Height - stringBounds.Height) / 2
    '        Case ContentAlignment.MiddleRight
    '            left = (label.Width - stringBounds.Width)
    '            top = (label.Height - stringBounds.Height) / 2
    '        Case ContentAlignment.TopCenter
    '            left = (label.Width - stringBounds.Width) / 2
    '            top = 0
    '        Case ContentAlignment.TopLeft
    '            left = 0
    '            top = 0
    '        Case ContentAlignment.TopRight
    '            left = (label.Width - stringBounds.Width)
    '            top = 0
    '    End Select
    '    Return New PointF(left, top)
    'End Function

    'These functions work with English and other left-to-right languages, they probably don't work
    'for right-to-left languages
    Private Shared Function GetHorizontalStringAlignment(align As ContentAlignment) As StringAlignment
        Select Case align
            Case ContentAlignment.BottomCenter
                Return StringAlignment.Center
            Case ContentAlignment.BottomLeft
                Return StringAlignment.Near
            Case ContentAlignment.BottomRight
                Return StringAlignment.Far
            Case ContentAlignment.MiddleCenter
                Return StringAlignment.Center
            Case ContentAlignment.MiddleLeft
                Return StringAlignment.Near
            Case ContentAlignment.MiddleRight
                Return StringAlignment.Far
            Case ContentAlignment.TopCenter
                Return StringAlignment.Center
            Case ContentAlignment.TopLeft
                Return StringAlignment.Near
            Case ContentAlignment.TopRight
                Return StringAlignment.Far
            Case Else
                Return StringAlignment.Near
        End Select
    End Function

    Private Shared Function GetVerticalStringAlignment(align As ContentAlignment) As StringAlignment
        Select Case align
            Case ContentAlignment.BottomCenter
                Return StringAlignment.Far
            Case ContentAlignment.BottomLeft
                Return StringAlignment.Far
            Case ContentAlignment.BottomRight
                Return StringAlignment.Far
            Case ContentAlignment.MiddleCenter
                Return StringAlignment.Center
            Case ContentAlignment.MiddleLeft
                Return StringAlignment.Center
            Case ContentAlignment.MiddleRight
                Return StringAlignment.Center
            Case ContentAlignment.TopCenter
                Return StringAlignment.Near
            Case ContentAlignment.TopLeft
                Return StringAlignment.Near
            Case ContentAlignment.TopRight
                Return StringAlignment.Near
            Case Else
                Return StringAlignment.Near
        End Select
    End Function

    Private Shared Sub DrawChildControls(c As Control, g As Graphics, drawBackground As Boolean)
        Dim childControl As Control

        For Each childControl In c.Controls
            Dim container As Drawing2D.GraphicsContainer = g.BeginContainer()

            'translate the effective origin of the graphics so that the
            'child control can be drawn at 0,0
            g.TranslateTransform(childControl.Left, childControl.Top)

            DrawControl(childControl, g, drawBackground)

            'translate origin back to upper-left of this control, by ending current container
            g.EndContainer(container)

            'g.TranslateTransform(-childControl.Left, -childControl.Top)
        Next
    End Sub
End Class
