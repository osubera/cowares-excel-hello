\name{midpoint}
\alias{midpoint}
\title{
Center of Indexes
}
\description{
Divide a given range of index into two of exact halves. 
}
\usage{
midpoint(x)
}
\arguments{
  \item{x}{
range of index. \code{c(min.index, max.index)}
}
}
\details{
This function implements the concept of \code{\link{median}}, 
and is used in the \code{\link{midpoints}} to calculate 
quantiles. 
}
\value{
A list of two numeric vectors is returned. 
\item{[[1]]}{a range of index for the lower half}
\item{[[2]]}{a range of index for the higher half}
Lengths of the two halves are exactly same.
If the parent range has an odd length, the exact center index 
is used both at the end of the lower half and at the start of 
the higher half.
}
\references{
Quantiles: median, quartiles, octiles, hexadeciles, \ldots 
\url{http://tomizonor.wordpress.com/2013/04/28/quantiles-octiles/}
}
\author{
Shinichiro Tomizono
}

\seealso{
  \code{\link{midpoints}},
  \code{\link{median}},
  \code{\link{range}}.
}
\examples{
midpoint(c(2,8))
# results are shown below.
# [[1]]
# [1] 2 5
# 
# [[2]]
# [1] 5 8

midpoint(c(2,9))
# results are shown below.
# [[1]]
# [1] 2 5
# 
# [[2]]
# [1] 6 9
}
\keyword{ utilities }
