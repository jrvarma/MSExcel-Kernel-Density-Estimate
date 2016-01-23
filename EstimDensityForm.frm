VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EstimDensityForm 
   Caption         =   "Estimate Density using Triangular or Gaussian Kernel"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   OleObjectBlob   =   "EstimDensityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EstimDensityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
'This form provides a convenient user interface to the Kernel method
'of density estimation. The actual computation is a few lines of code
'in the subroutine estimate. Rest is the user interface
'
'
'Global Variables
'Data is the data - the set of observations
'x is the range of values (abscissa) for which the density is required
'out is where the estimated densities are written
Dim Data As Range, x As Range, out As Range
'H is the bandwidth of the kernel
Dim H
'Boolean variables:
'whether the mean is to be removed
'whether the Gaussian kernel is to be used or the Triangular one
'whether the data is to be rescaled to unit variance
'whether parameters are to be customised or defaulted
Dim Demean, Gaussian, Rescale, Custom
'
'
Private Sub estimate(Data As Object, x As Object, out As Object, H As Double, Gaussian As Boolean, Demean As Boolean, Rescale As Boolean)
'This is the computational routine
Sigma = Application.StDev(Data)
If Not Rescale Then
    H = H * Sigma
    Sigma = 1
End If
If Demean Then mu = Application.Average(Data) Else mu = 0
If Gaussian Then Factor = Sqr(2 * Application.Pi) Else Factor = 1
For k = 1 To x.Count
    'This provides a percentage completion status
    PctComplete.Caption = Int((k - 1) * 100 / x.Count) & "%"
    EstimDensityForm.Repaint
    p = 0
    For i = 1 To Data.Count
        z = (x.Cells(k) - (Data.Cells(i) - mu) / Sigma) / H
        If Gaussian Then
            'Gaussian kernel is exp(-0.5*z*z)/Factor
            'We divide by Factor outside the loop
            p = p + Exp(-0.5 * z * z)
        Else
            'Triangular kernel is
            '1-Abs(z) for -1 <= z <= 1 and
            '0 elsewhere
            If Abs(z) < 1 Then p = p + (1 - Abs(z))
        End If
    Next i
    out.Cells(k) = p / Factor / (Data.Count * H)
Next k
Exit Sub
End Sub
'
'
Private Sub CancelButton_Click()
Unload EstimDensityForm
EstimDensityCancel = True
End Sub
'
'
Private Function BandWidth()
'This is the default bandwidth which is essentially inverse of n^(1/5)
H = Data.Count ^ (-0.2)
hlog = Int(Application.Log10(H))
BandWidth = 10 ^ hlog * Application.Round(2 * H / (10 ^ hlog), 0) / 2
End Function
'
'
Private Sub DataRE_Change()
Set Data = Range(DataRE.Value)
BandwidthTB.Value = BandWidth()
End Sub
'
'
Private Sub MoreOrLessCB_Click()
'Toggles between detailed list of options and simpler screen
Custom = Not Custom
If Custom Then
    MoreOrLessCB.Caption = "Less <<"
    AbscissaRE.Visible = True
    OutputRE.Visible = True
    BandwidthTB.Visible = True
    GaussianCB.Visible = True
    DemeanCB.Visible = True
    RescaleCB.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
Else
    MoreOrLessCB.Caption = "More >>"
    AbscissaRE.Visible = False
    OutputRE.Visible = False
    BandwidthTB.Visible = False
    GaussianCB.Visible = False
    DemeanCB.Visible = False
    RescaleCB.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
End If
End Sub
'
'
Private Sub OKButton_Click()
Dim MySheet
Set Data = Range(DataRE.Value)
If Custom Then
    Set x = Range(AbscissaRE.Value)
    Set out = Range(OutputRE.Value)
    H = BandwidthTB.Value
    Gaussian = GaussianCB.Value
    Demean = DemeanCB.Value
    Rescale = RescaleCB.Value
Else
    'Create new sheet
    Set MySheet = Worksheets.Add
    MySheet.Range("A1") = "Abscissa"
    MySheet.Range("B1") = "Density"
    Set x = MySheet.Range("A2:A32")
    Set out = MySheet.Range("B2:B32")
    'fill up x values
    Call FillXRange
    H = BandWidth()
    Gaussian = True
    Demean = True
    Rescale = True
End If
'We store the current values of all user parameters as the
'new defaults for this Excel session
'This function is in the DensityEstim module to allow static
'variables to be used
Call EstimDensityDefaults(Data, x, out, H, Demean, Gaussian, _
     Rescale, True)
'set up a wait message
WaitMessage.Visible = True
WaitMessage.ZOrder
PctComplete.Visible = True
PctComplete.ZOrder
EstimDensityForm.Repaint
'estimate the density
Call estimate(Data, x, out, (H), (Gaussian), (Demean), (Rescale))
'plot the estimated density
Call MakeChart
WaitMessage.Visible = False
PctComplete.Visible = False
Unload EstimDensityForm
End Sub
'
'
Private Sub UserForm_Initialize()
'Read default values. In case this form has been invoked earlier
'in this Excel session, the values are remembered from that invocation
Call EstimDensityDefaults(Data, x, out, H, Demean, Gaussian, _
     Rescale, False)
'Data, x and output ranges are also remembered from earlier
'invocation if any. Else these are all empty
Call ValidateRange(Data)
If Data Is Nothing Then H = Empty
Call ValidateRange(x)
Call ValidateRange(out)
DataRE.Value = RangeAddress(Data, "")
OutputRE.Value = RangeAddress(out, "")
AbscissaRE.Value = RangeAddress(x, "")
If IsEmpty(Demean) Then Demean = True
If IsEmpty(Gaussian) Then Gaussian = True
If IsEmpty(Rescale) Then Rescale = True
If IsEmpty(Custom) Then Custom = False
DemeanCB.Value = Demean
GaussianCB.Value = Gaussian
RescaleCB.Value = Rescale
BandwidthTB.Value = H
End Sub
'
'
Private Function FillXRange()
'This fills in a default x range which runs
'-6, -5, -4, -3, -2.75, -2.5, ..., 2.5, 2.75, 3, 4, 5, 6
Dim i As Integer
For i = 1 To 4
x.Cells(i) = i - 7
x.Cells(i + 27) = i + 2
Next i
For i = 5 To 27
x.Cells(i) = 0.25 * (i - 4) - 3
Next i
End Function
'
'
Private Sub MakeChart()
Charts.Add
'smoothed line XY chart
ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
ActiveChart.SetSourceData Source:=Union(x, out), PlotBy:=xlColumns
'located as an object on the same sheet as the output range
ActiveChart.Location Where:=xlLocationAsObject, Name:=out.Worksheet.Name
With ActiveChart
    'set up axes, no legends
    .HasAxis(xlCategory, xlPrimary) = True
    .HasAxis(xlValue, xlPrimary) = True
    .HasLegend = False
End With
ActiveChart.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
If Not Custom Then
    'For default chart, axis runs from -6 to 6
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = -6
        .MaximumScale = 6
        .MinorUnit = 0.25
        .MajorUnit = 1
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
    End With
End If
ActiveChart.Axes(xlValue).Select
With Selection.Border
    .Weight = xlHairline
    .LineStyle = xlAutomatic
End With
With Selection
    .MajorTickMark = xlOutside
    .MinorTickMark = xlNone
    .TickLabelPosition = xlNone
End With
ActiveChart.PlotArea.Select
'white background
Selection.Interior.ColorIndex = xlNone
'ActiveChart.ChartArea.Select
ActiveChart.Axes(xlCategory).Select
With Selection
    .MajorTickMark = xlCross
    .MinorTickMark = xlOutside
    .TickLabelPosition = xlNextToAxis
End With
End Sub


