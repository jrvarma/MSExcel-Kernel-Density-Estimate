# MSExcel-Kernel-Density-Estimate
This Visual Basic code computes a kernel based density estimate from data contained in an MS Excel spreadsheet. The default bandwidth can also be changed.

The software consists of a form and some VBA code. The form allows the user to specify the range of cells in the spreadsheet containing the data. The following additional optional inputs can also be given:
* The set of points at which the density is to be estimated
* Whether the data is to be demeaned
* Whether the data is to be rescaled to unit standard deviation 
* Use a triangular kernel instead of the default Gaussian kernel
* Change the default bandwidth
* The output range (By default a new sheet is created)
