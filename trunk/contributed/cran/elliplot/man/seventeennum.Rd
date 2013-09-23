\name{seventeennum}
\alias{seventeennum}
\alias{hexadecile}
\title{
Seventeen Number Summaries
}
\description{
Return seventeen number summary (minimum, 1st-7th hexadeciles, median, 9th-15th hexadeciles, maximum) for the input data.
}
\usage{
seventeennum(x, na.rm=TRUE)
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
This function is similar to \code{\link{fivenum}}.  This returns hexadeciles instead of quartiles of the \code{\link{fivenum}}. 
}
\value{
A numeric vector of length 17 containing the summary information. 
}
\references{
Quantiles: median, quartiles, octiles, hexadeciles, \ldots 
\url{http://tomizonor.wordpress.com/2013/04/28/quantiles-octiles/}
}
\author{
Shinichiro Tomizono
}
\note{
Internally calling \code{\link{midpoints}} with \code{n=4}.
}

\seealso{
  \code{\link{midpoints}},
  \code{\link{ninenum}},
  \code{\link{fivenum}},
  \code{\link{median}}.
}
\examples{
seventeennum(1:100)
seventeennum(c(rnorm(100), -Inf, Inf))
}
\keyword{ utilities }
