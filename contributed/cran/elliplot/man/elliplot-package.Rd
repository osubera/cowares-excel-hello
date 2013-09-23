\name{elliplot-package}
\alias{elliplot-package}
\alias{elliplot}
\docType{package}
\title{
Ellipse Summary Plot of Quantiles
}
\description{
Correlation chart of two set (x and y) of data. 
Using Quantiles.  Visualize the effect of factor. 
}
\details{
\tabular{ll}{
Package: \tab elliplot\cr
Type: \tab Package\cr
Version: \tab 1.0.0\cr
Date: \tab 2013-09-15\cr
License: \tab \href{http://paidforeveryone.wordpress.com}{paid4e license}: Fortitudinous, Free, Fair. \cr
}
This package contains quantile functions and ellipse 
plot functions.  These functions are to calculate 
quantile summaries and visualize them with ellipses.  

The ellipseplot works both for 1 set and 2 sets of data.  
When used for 2 sets data, it visualize the correlation 
of x and y axis.  
}
\author{
Shinichiro Tomizono

Maintainer: Shinichiro Tomizono <cowares@gmail.com>
}
\references{
Quantiles: \url{http://tomizonor.wordpress.com/2013/04/28/quantiles-octiles/}\cr
Ellipse Plot: \url{http://tomizonor.wordpress.com/2013/04/29/ellipse-plot/}
}
\keyword{ package }
\keyword{ hplot }
\keyword{ utilities }
\seealso{
  \code{\link{midpoints}},
  \code{\link{ellipseplot}}.
}
\examples{
ninenum(1:999)
ellipseplot(iris[c(5,1)], iris[c(5,2)])
}
