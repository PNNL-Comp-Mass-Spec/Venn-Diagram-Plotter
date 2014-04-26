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

Imports System.Drawing

Public Interface IPrintableControl
    Sub DrawOnGraphics(ByVal g As Graphics, ByVal drawBackground As Boolean)
End Interface

'This interface tags controls that are merely containers for other controls, and are drawn 
'solely by drawing their child controls and possibly the background
'Since there is nothing to draw, no methods are needed.
Public Interface IPrintableControlContainer

End Interface
