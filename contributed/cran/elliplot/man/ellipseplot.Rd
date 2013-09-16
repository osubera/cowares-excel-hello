\name{ellipseplot}
\alias{ellipseplot}
\alias{ellipseplot.default}
\alias{ellipseplot.data.frame}
\alias{ellipseplot.numeric}
\alias{ellipseplot.matrix}
\alias{ellipseplot.list}
\title{
Draw Ellipse Summary Plot
}
\description{
Correlation chart of two set (x and y) of data. 
Using Quantiles.  Visualize the effect of factor. 
}
\usage{
ellipseplot(x, y=NULL, SUMMARY=ninenum, SHEER=sheer.color,
            plot=TRUE, verbose=FALSE, ...)
}
\arguments{
  \item{x}{
An x-axis data, such as data frame of factors (1st column) and 
observations (2nd column).  
A vector, a matrix or a list is also acceptable.  
If a vector is given, a single ellipse without factors are 
drawn. 
A matrix is as same format as the data frame.  
A list is formed by factors with observatoin vectors as each item.  
}
  \item{y}{
A y-axis data, such as data frame of factors (1st column) and 
observations (2nd column).  
Same types as the x-axis data are also acceptable.  
It can be a NULL (default), to draw a single axis chart.  
}
  \item{SUMMARY}{
A function generating summaries to write contours of ellipses.  
The default is ninenum to use nine number summary.  
}
  \item{SHEER}{
A function adjusting color levels of ellipses. 
The default is sheer.color function shown below.  
    \preformatted{
sheer.color <- function(col, level) \{
  sheer <- level^2 * 0.5
  adjustcolor(col, alpha.f=sheer)
\}
    }
}
  \item{plot}{
If FALSE is given, it disable to plot and print a summary.  
The default is TRUE.  
}
  \item{verbose}{
If TRUE is given, it print verbose debugging information. 
The default is FALSE.  
}
  \item{\dots}{
Plot parameters are acceptable.  
}
}
\details{
}
\value{
A summary list is explicitly printed when plot=FALSE is given, 
and is invisibly returned when plot=TRUE.  
}
\references{
Ellipse Summary Plot
\url{http://tomizonor.wordpress.com/2013/04/29/ellipse-plot/}
}
\author{
Shinichiro Tomizono
}
\note{
}

\seealso{
  \code{\link{ninenum}}
  \code{\link{seventeennum}}
  \code{\link{midpoints}}
}
\examples{
ellipseplot(iris[c(5,1)], iris[c(5,2)])
\keyword{ hplot }
