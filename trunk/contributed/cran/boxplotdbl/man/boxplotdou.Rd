\name{boxplotdou}
\alias{boxplotdou}
\alias{boxplotdou.default}
\alias{boxplotdou.data.frame}
\alias{boxplotdou.factor}
\alias{boxplotdou.formula}
\alias{boxplotdou.list}
\title{
Double Box Plot
}
\description{
Extend box plot chart into two-axes (x and y) to visualize correlation.
}
\usage{
boxplotdou(x, ...)

\method{boxplotdou}{default}(x, y, 
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=factor.labels, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           STAT=bxpdou.boxplot.stat, 
           verbose=FALSE, plot=TRUE, ...) 

\method{boxplotdou}{data.frame}(x, y, 
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=factor.labels, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           STAT=bxpdou.boxplot.stat, 
           verbose=FALSE, plot=TRUE, ...) 

\method{boxplotdou}{factor}(x, obs.x, f.y, obs.y,  
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=factor.labels, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           STAT=bxpdou.boxplot.stat, 
           verbose=FALSE, plot=TRUE, ...) 

\method{boxplotdou}{formula}(formula.x, data.x, formula.y, data.y,  
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=factor.labels, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           STAT=bxpdou.boxplot.stat, 
           verbose=FALSE, plot=TRUE, ...) 

\method{boxplotdou}{list}(x, 
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=factor.labels, factor.labels=NULL, draw.legend=NA,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           verbose=FALSE, plot=TRUE, ...) 
}
\arguments{
  \item{x}{
data frame, 
contains two columns as factor and observation to x-axis
(for \code{boxplotdou.default}, that is,  = \code{boxplotdou.data.frame}). \cr
factor vector, as factor to x-axis (for \code{boxplotdou.factor}). \cr
list, output values of boxplotdou, 
eg. previously saved stat to redraw a chart, 
(for \code{boxplotdou.list}). 
  }
  \item{y}{
data frame, 
contains two columns as factor and observation to y-axis
(for \code{boxplotdou.default}, that is,  = \code{boxplotdou.data.frame}). 
  }
  \item{obs.x}{
numeric vector, as observation to x-axis (for \code{boxplotdou.factor}). 
  }
  \item{f.y}{
factor vector, as factor to y-axis (for \code{boxplotdou.factor}). 
  }
  \item{obs.y}{
numeric vector, as observation to y-axis (for \code{boxplotdou.factor}). 
  }
  \item{formula.x}{
formula, a model formula to x-axis, eg. \code{obs ~ factor} 
(for \code{boxplotdou.formula}). 
  }
  \item{data.x}{
data.frame, contains variables in formula.x
(for \code{boxplotdou.formula}). 
  }
  \item{formula.y}{
formula, a model formula to y-axis, eg. \code{obs ~ factor} 
(for \code{boxplotdou.formula}). 
  }
  \item{data.y}{
data.frame, contains variables in formula.y
(for \code{boxplotdou.formula}). 
  }
  \item{boxed.whiskers}{ logical, default is FALSE,
TRUE to draw rectangular range rather than whisker and staple. 
  }
  \item{outliers.has.whiskers}{ logical, default is FALSE, 
extend whisker and staple through outliers. 
  }
  \item{name.on.axis}{ control labels on each group on axes, 
default is \code{factor.labels}, 
NULL to use factor data, 
TRUE to abbreviate by alphabet letters, 
FALSE to draw no labels, 
character vector to give explicit labels, 
single character to use identical character. 
  }
  \item{factor.labels}{ control labels on each group on factor, 
default is NULL, using factor data, 
TRUE to abbreviate by alphabet letters, 
FALSE to draw no labels, 
character vector to give explicit labels, 
single character to use identical character, 
NA in vector to exclude any groups. 
  }
  \item{draw.legend}{ logical, draw legend or not, 
default is NA, enable legend only when labels abbreviated. 
  }
  \item{condense}{ logical, default is FALSE, 
TRUE to unify near groups into one box. 
  }
  \item{condense.severity}{ character, default is \code{"iqr"},
one of \code{c('iqr','whisker','iqr.xory','whisker.xory')},
which is the border to condense or not,
used only when \code{condense=TRUE}. 
  }
  \item{condense.once}{ logical, default is FALSE,
TRUE to disable recursive condenses,
used only when \code{condense=TRUE}. 
  }
  \item{col}{ character vector, colors for each group, 
default is NULL, automatic colors. 
  }
  \item{COLOR.SHEER}{ function, to convert color to sheer color,
default is \code{bxpdou.sheer.color}, internally defined as,
\code{function(col) adjustcolor(col, alpha.f=0.2)}.
sheer colors are used for inside box, 
or for \code{outliers.has.whiskers=TRUE}. 
  }
  \item{shading}{ numeric vector, 
as shading density to draw inside of box.
default is NA, means automatic, 
no shadings when both \code{shading} and \code{shading.angle} are NA. 
the density value means lines per inch. 
the generator depends the vector length. 
    \describe{
      \item{1}{a sigle value is used to all densities}
      \item{2}{values are generated between the pair}
      \item{k}{when k is number of factor levels, values are used to each level}
    }
logical value TRUE has a special meaning to enable shading with automatic densities.  
  }
  \item{shading.angle}{ numeric vector, 
as shading angle to draw inside of box.
default is NA, means automatic, 
no shadings when both \code{shading} and \code{shading.angle} are NA. 
the angle value means degree of line direction to horizon. 
the generator depends the vector length with same manner of density.
  }
  \item{blackwhite}{ logical, default is FALSE,
TRUE to draw black and white chart,
equivalent to set following 3 parameters,
\code{col='black', shading=TRUE, COLOR.SHEER=(function(a) a)}
  }
  \item{STAT}{ function, default is \code{bxpdou.boxplot.stat},
internally defined as,
\code{function(formula) boxplot(formula=formula, plot=FALSE)}, 
delegates to standard \code{boxplot} function. 
specify user function to calculate summary. 
  }
  \item{plot}{
if FALSE is given, it disable to plot and print a summary.  
default is TRUE.  
  }
  \item{verbose}{
if TRUE is given, it print verbose debugging information. 
default is FALSE.  
  }
  \item{\dots}{
plot parameters and boxplot color parameters are acceptable.  
  }
}
\section{boxplot color parameters}{
  \describe{
    \item{medcol}{ default is NULL, to use black, 
    colors for median labels.
    }
    \item{whiskcol}{ default is NULL, to use \code{col}, 
    colors for whiskers.
    }
    \item{staplecol}{ default is NULL, to use \code{col}, 
    colors for staples.
    }
    \item{boxcol}{ default is NULL, to use black, 
    colors for box borders.
    }
    \item{outcol}{ default is NULL, to use \code{col}, 
    colors for outliers.
    }
    \item{outbg}{ default is NULL, to use transparent, 
    colors inside outliers.
    }
    \item{outcex}{ default is 2, size of outliers.
    }
    \item{outpch}{ default is 1, to use a transparent circle, 
    symbol number of outliers, as graphic \code{\link{par}} \code{pch}.
    }
  }
}
\details{
This function is designed to visualize a correlation 
between 2 sets of independent observation with common 
factors. 
Such as, the plant height v.s. the soil pH by location.\cr
This function depends on \code{\link{boxplot}} function 
to calculate summaries such as IQRs. 
This dependency can be overridden by \code{STAT} argument. 
}
\value{
A summary list is explicitly printed when plot=FALSE is given, 
and is invisibly returned when plot=TRUE.  
  \item{stat}{
    \describe{
      \item{x}{ x-axis summary, same as boxplot statistics }
      \item{y}{ y-axis summary, same as boxplot statistics }
    }
  }
  \item{name}{
    \describe{
      \item{x}{ character, x-axis label }
      \item{y}{ character, y-axis label }
    }
  }
  \item{level}{
    character vecotr of factor names
  }
  Each summary of x and y is identical to \code{\link{boxplot}} statistics,
  \item{stats}{
    matrix, each column contains the extreme of the lower whisker, 
    the lower hinge, the median, the upper hinge and 
    the extreme of the upper whisker.
  }
  \item{n}{
    numerical vector, sample numbers of each factor level.
  }
  \item{conf}{
    matrix, each column contains the lower and upper extremes of the notch.
  }
  \item{out}{
    numerical vector, outliers
  }
  \item{group}{
    numerical vector of same length as \code{out}, 
    indicates which factor level the outlier belongs.
  }
  \item{names}{
    character vector, each name of factor levels.
  }
}
\references{
Double Box Plot: \url{http://tomizonor.wordpress.com/2013/03/15/double-box-plot/}\cr
Double Box Plot: \url{http://tomizonor.wordpress.com/2013/03/15/double-box-plot/}
}
\author{
Shinichiro Tomizono
}

\seealso{
  \code{\link{boxplot}},
  \code{\link{fivenum}}.
}
\examples{
# iris data: Sepal.Length v.s. Sepal.Width by Species
stat <- boxplotdou(iris[c(5,1)], iris[c(5,2)])
boxplotdou(iris[,5], iris[,1], iris[,5], iris[,2])
boxplotdou(Sepal.Length~Species, iris, Sepal.Width~Species, iris)
boxplotdou(stat, main='redraw by saved stat')

# color and shading
boxplotdou(iris[c(5,1)], iris[c(5,2)], col=c('wheat','wheat','black'),
           boxcol='springgreen')
boxplotdou(iris[c(5,1)], iris[c(5,2)], shading=c(3,5))
boxplotdou(iris[c(5,1)], iris[c(5,2)], shading=5, shading.angle=c(0,90))
boxplotdou(iris[c(5,1)], iris[c(5,2)], blackwhite=TRUE)

# customized sheer funtion
mysheer <- function(x) 
             adjustcolor(x, alpha.f=0.2, red.f=0.3, green.f=0.3, blue.f=0.3)
boxplotdou(iris[c(5,1)], iris[c(5,2)], COLOR.SHEER=mysheer)

# whisker
boxplotdou(iris[c(5,1)], iris[c(5,2)], boxed.whiskers=TRUE)
boxplotdou(iris[c(5,1)], iris[c(5,2)], outliers.has.whiskers=TRUE)

# condense
boxplotdou(iris[c(5,1)], iris[c(5,2)], condense=TRUE)

# labels
boxplotdou(iris[c(5,1)], iris[c(5,2)], factor.labels=FALSE)
boxplotdou(iris[c(5,1)], iris[c(5,2)], factor.labels=TRUE)
boxplotdou(iris[c(5,1)], iris[c(5,2)], factor.labels=TRUE, 
                                       draw.legend=FALSE)
boxplotdou(iris[c(5,1)], iris[c(5,2)], factor.labels=c('Se','Ve','Vi'))
boxplotdou(iris[c(5,1)], iris[c(5,2)], factor.labels='+', 
                                       name.on.axis=FALSE)

# customized summary function
mystat <- function(x) boxplot(formula=x, range=1, plot=FALSE)
boxplotdou(iris[c(5,1)], iris[c(5,2)], STAT=mystat)

# graphic parameters
boxplotdou(iris[c(5,1)], iris[c(5,2)], xlim=c(4.8, 7.0), ylim=c(2.0, 3.5))

# print summary
boxplotdou(iris[c(5,1)], iris[c(5,2)], plot=FALSE)
}
\keyword{ hplot }
