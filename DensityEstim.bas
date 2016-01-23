Attribute VB_Name = "DensityEstim"
'
'
'   Copyright (C) 2001  Prof. Jayanth R. Varma, jrvarma@iimahd.ernet.in,
'   Indian Institute of Management, Ahmedabad 380 015, INDIA
'
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program (see file COPYING); if not, write to the
'   Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'   Boston, MA  02111-1307  USA
'
'
'Functions relating to kernal density estimation
'The function KernelDensity is for use in formulas while the other
'functions are for use with the form EstimDensityForm

'
'
Static Sub EstimDensity()
'This is required only for the EstimDensity form
EstimDensityForm.Show
End Sub
'
'
Static Function EstimDensityDefaults(Data, x, out, H, Demean, _
                Gaussian, Rescale, store As Boolean)
'This is required only for the EstimDensity form
'It is defined here to allow the use of static variables
'whose values are remembered from invocation to invocation
'within the same Excel session
Static Data0 As Range, X0 As Range, out0 As Range
Static h0
Static Demean0, Gaussian0, Rescale0
'When called with store=True, the arguments are stored as new defaults
'Else the relevant global variables are set to their current defaults
If store Then
    Set Data0 = Data
    Set X0 = x
    Set out0 = out
    h0 = H
    Demean0 = Demean
    Gaussian0 = Gaussian
    Rescale0 = Rescale
Else
    Set Data = Data0
    Set x = X0
    Set out = out0
    H = h0
    Demean = Demean0
    Gaussian = Gaussian0
    Rescale = Rescale0
End If
End Function
'
'
Function KernelDensity(n As Integer, H As Double, x As Double, AddString As String) As Double
Attribute KernelDensity.VB_ProcData.VB_Invoke_Func = " \n14"
'This is for use in formulas (instead of using the form)
'The density is estimated for the given value of x (abscissa)
'Using the given bandwidth H
'AddString is a string containing the address of the data
'A Gaussian kernel is used
Dim DataRange As Range
Set DataRange = Range(AddString)
hs = H * Application.StDev(DataRange)
p = 0
For i = 1 To n
    z = (x - DataRange.Cells(i)) / H
    p = p + Exp(-0.5 * z * z)
Next i
KernelDensity = p / Sqr(2 * Application.Pi) / (n * H)
End Function

