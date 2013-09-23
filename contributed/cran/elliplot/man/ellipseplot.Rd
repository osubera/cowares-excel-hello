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
ellipseplot(x, ...)

\method{ellipseplot}{default}(x, y=NULL, SUMMARY=ninenum, SHEER=sheer.color,
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
A function generating quantile summaries to write contours of ellipses.  
The default is ninenum to use nine number summary.  
The function must return an odd length numerical vector, 
because a center, such as median, is required.  
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
This function is designed to visualize a correlation 
between 2 sets of independent observation with common 
factors. 
Such as, the plant height v.s. the soil pH by location.
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

\seealso{
  \code{\link{ninenum}},
  \code{\link{seventeennum}},
  \code{\link{midpoints}}.
}
\examples{
# iris data: Sepal.Length v.s. Sepal.Width by Species
ellipseplot(iris[c(5,1)], iris[c(5,2)])

# PlantGrowth data: weight by group : single axis
# five number summaries are used.  
# similar to boxplot(weight~group,PlantGrowth)
ellipseplot(PlantGrowth[2:1], SUMMARY=fivenum)

# iris data: Sepal.Length v.s. Sepal.Width without factor
ellipseplot(iris[,1], iris[,2], xlab='Sepal.Length', ylab='Sepal.Width')

# list example
ellipseplot(list(untreated=rnorm(30,3,1), treated=rnorm(30,5,2)), 
            list(untreated=rnorm(20,6,3), treated=rnorm(20,4,2)))

# using cutomized sheer function
my.sheer.color <- function(col, level) adjustcolor(rainbow(100)[runif(1,1,100)], alpha.f=0.4)
ellipseplot(iris[c(5,1)], iris[c(5,2)], SHEER=my.sheer.color)

}
\keyword{ hplot }
