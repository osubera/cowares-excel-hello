\name{boxplotdbl-package}
\alias{boxplotdbl-package}
\alias{boxplotdbl}
\docType{package}
\title{
Double Box Plot
}
\description{
Correlation chart of two set (x and y) of data. 
Extend the box plot chart for two-axes to show the
correlation of independent two set of observations
with common factor.  Quartiles and outliers are 
calculated by standard boxplot function. 
}
\details{
\tabular{ll}{
Package: \tab boxplotdbl\cr
Type: \tab Package\cr
Version: \tab 1.2.0\cr
Date: \tab 2013-11-23\cr
License: \tab MIT | \href{http://paidforeveryone.wordpress.com}{paid4e}: Copyright (c) 2013, Shinichiro Tomizono\cr
}
This package contains boxplotdou function. 
It is used for 2 sets data, to visualize the correlation 
of x and y axis.  
}
\author{
Shinichiro Tomizono

Maintainer: Shinichiro Tomizono <cowares@gmail.com>
}
\references{
Double Box Plot: \url{http://tomizonor.wordpress.com/2013/03/15/double-box-plot/}\cr
Double Box Plot: \url{http://tomizonor.wordpress.com/2013/03/15/double-box-plot/}
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
