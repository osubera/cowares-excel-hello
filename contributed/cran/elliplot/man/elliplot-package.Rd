\name{elliplot-package}
\alias{elliplot-package}
\alias{elliplot}
\docType{package}
\title{
\packageTitle{elliplot}
}
\description{
\packageDescription{elliplot}
}
\details{

The DESCRIPTION file:
\packageDESCRIPTION{elliplot}
\packageIndices{elliplot}
This package contains quantile functions and ellipse 
plot functions.  These functions are to calculate 
quantile summaries and visualize them with ellipses.  

The ellipseplot works both for 1 set and 2 sets of data.  
When used for 2 sets data, it visualize the correlation 
of x and y axis.  
}
\author{
\packageAuthor{elliplot}

Maintainer: \packageMaintainer{elliplot}
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
