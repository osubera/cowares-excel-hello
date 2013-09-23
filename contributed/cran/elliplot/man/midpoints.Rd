\name{midpoints}
\alias{midpoints}
\alias{quantile}
\title{
Quantile Summaries
}
\description{
Return quantile summary (minimum, quantiles, maximum) for the input data.
}
\usage{
midpoints(x, n = 1, na.rm = TRUE)
}
\arguments{
  \item{x}{
numeric, maybe including \code{NA}s and +/-\code{Inf}s.
}
  \item{n}{
positive integer, to determine which quantiles to calculate.  
return values are \eqn{2 ^ n + 1} number summary.
\code{n=1} (default) is for three number, namely, minimum, median 
and maximum.  \code{n=2} is for five number, such as quartiles. 
\code{n=3} is for nine number, such as octiles.  
}
  \item{na.rm}{
logical value indicating whether \code{NA}s should be stripped before the computation proceeds.
}
}
\details{
This function is calling \code{\link{midpoint}} with \code{n} depth.
}
\value{
A numeric vector of length \eqn{2 ^ n + 1} containing the summary information. 
}
\references{
Quantiles: median, quartiles, octiles, hexadeciles, \ldots 
\url{http://tomizonor.wordpress.com/2013/04/28/quantiles-octiles/}
}
\author{
Shinichiro Tomizono
}

\seealso{
  \code{\link{midpoint}},
  \code{\link{seventeennum}},
  \code{\link{ninenum}},
  \code{\link{fivenum}},
  \code{\link{median}}.
}
\examples{
midpoints(1:100, 4)
midpoints(c(rnorm(100), -Inf, Inf), 3)

# define 33 number summary
thirtythreenum <- function(x, ...) midpoints(x, 5, ...)
thirtythreenum(1:100)
}
\keyword{ utilities }
