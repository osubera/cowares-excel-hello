\name{ninenum}
\alias{ninenum}
\alias{octile}
\title{
Nine Number Summaries
}
\description{
Return nine number summary (minimum, 1st-3rd octiles, median, 5th-7th octiles, maximum) for the input data.
}
\usage{
ninenum(x, na.rm=TRUE)
}
\arguments{
  \item{x}{
numeric, maybe including \code{NA}s and +/-\code{Inf}s.
}
  \item{na.rm}{
logical value indicating whether \code{NA}s should be stripped before the computation proceeds.
}
}
\details{
This function is similar to \code{\link{fivenum}}.  This returns octiles instead of quartiles of the \code{\link{fivenum}}. 
}
\value{
A numeric vector of length 9 containing the summary information. 
}
\references{
Quantiles: median, quartiles, octiles, hexadeciles, \ldots 
\url{http://tomizonor.wordpress.com/2013/04/28/quantiles-octiles/}
}
\author{
Shinichiro Tomizono
}
\note{
Internally calling \code{\link{midpoints}} with \code{n=3}.
}

\seealso{
  \code{\link{midpoints}},
  \code{\link{seventeennum}},
  \code{\link{fivenum}},
  \code{\link{median}}.
}
\examples{
ninenum(1:100)
ninenum(c(rnorm(100), -Inf, Inf))
}
\keyword{ utilities }
