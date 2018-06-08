\name{boxplotdbl-package}
\alias{boxplotdbl-package}
\alias{boxplotdbl}
\docType{package}
\title{
\packageTitle{boxplotdbl}
}
\description{
\packageDescription{boxplotdbl}
}
\details{

The DESCRIPTION file:
\packageDESCRIPTION{boxplotdbl}
\packageIndices{boxplotdbl}
This package contains boxplotdou function. 
It is used for 2 sets data, to visualize the correlation 
of x and y axis.  
}
\author{
\packageAuthor{boxplotdbl}

Maintainer: \packageMaintainer{boxplotdbl}
}
\references{
Double Box Plot: \url{http://tomizonor.wordpress.com/2013/03/15/double-box-plot/}\cr
Double Box Plot 1.2: \url{http://tomizonor.wordpress.com/2013/11/24/double-box-plot-1-2/}
}
\keyword{ package }
\keyword{ hplot }
\seealso{
  \code{\link{boxplotdou}},
  \code{\link{boxplot}}.
}
\examples{
boxplotdou(Sepal.Width~Species, iris, Petal.Width~Species, iris)
}
