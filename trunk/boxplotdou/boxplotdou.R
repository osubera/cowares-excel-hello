# double boxplot
# http://code.google.com/p/cowares-excel-hello/source/browse/trunk/boxplotdou/
#
# Copyright (C) 2013 Tomizono
# Fortitudinous, Free, Fair, http://cowares.nobody.jp
#
# boxplotdou(cbind(factor, data1), cbind(factor, data2))

boxplotdou <- function(x, ...) UseMethod("boxplotdou")

boxplotdou.default <- 
function(x, y, 
         color=NULL, colorSheer=NULL, 
         boxedWhiskers=FALSE, outliersHasWhiskers=FALSE, nameOnAxis=TRUE, 
         pars=NULL, verbose=FALSE, plot=TRUE, ...) {

  # both x and y expect data frame with 2 columns:  factor, observation

  data <- list(x=x[,2], y=y[,2])
  name <- list(x=names(x[2]), y=names(y[2]))
  stat <- list()

  # merge and unite the factor levels of x and y
  factorLevels <- unique(c(unique(as.character(x[,1])),unique(as.character(y[,1]))))
  factor <- list(x=factor(x[,1],levels=factorLevels), 
                 y=factor(y[,1],levels=factorLevels))

  for(v in c("x","y")) {
    # delegate the calculation to the standard boxplot
    stat[[v]] <- boxplot(formula=formula(paste("data$", v, "~factor$", v, sep="")), 
                         plot=FALSE)
    if(is.null(name[[v]])) name[[v]]=v
  }

  if(verbose) {
    print(name)
    print(data)
    print(stat)
  }

  if(plot) {
    bxpdou(stat$x, stat$y, factorLevels,
           xlab=name$x, ylab=name$y, pars=par(), 
           color=color, colorSheer=colorSheer, 
           boxedWhiskers=boxedWhiskers, outliersHasWhiskers=outliersHasWhiskers,
           nameOnAxis=nameOnAxis, verbose=verbose, ...)
    invisible(stat)
  } else {
    stat
  }
}

bxpdou <- 
function(xStats, yStats, factorLevels, 
         color=NULL, colorSheer=NULL, 
         boxedWhiskers=FALSE, outliersHasWhiskers=FALSE, nameOnAxis=TRUE, 
         pars=NULL, verbose=FALSE, ...) {
  
  pars <- c(list(...), pars)
  # the first overrides the later
  pars <- pars[unique(names(pars))]

  xMin <- min(xStats$stats, na.rm=TRUE)
  xMax <- max(xStats$stats, na.rm=TRUE)
  yMin <- min(yStats$stats, na.rm=TRUE)
  yMax <- max(yStats$stats, na.rm=TRUE)

  if(is.null(pars$xlim)) xlim <- c(xMin, xMax)
  if(is.null(pars$ylim)) ylim <- c(yMin, yMax)

  levelsNum <- length(factorLevels)
  levelsCol <- rainbow(levelsNum)
  ##FIXME color and colorSheer is not used.
  
  # open a plot area and draw axis
  plot(NULL, xlim=xlim, ylim=ylim, ...)

  # draw boxes
  for(i in 1L:levelsNum) {
    bxpdou.abox(xStats, yStats, 
                columnNum=i, columnChar=as.character(factorLevels)[i], 
                color=levelsCol[i], colorSheer=NULL, nameOnAxis=nameOnAxis, 
                boxedWhiskers=boxedWhiskers, outliersHasWhiskers=outliersHasWhiskers,
                verbose=verbose)
  }
}

bxpdou.abox <- 
function(x, y, columnNum, columnChar, 
         color, colorSheer=NULL, 
         boxedWhiskers=FALSE, outliersHasWhiskers=FALSE, nameOnAxis=TRUE, 
         verbose=FALSE) {

  # declare five numbers explicitly

  xLowest  <- x$stats[1, columnNum]
  xHighest <- x$stats[5, columnNum]
  yLowest  <- y$stats[1, columnNum]
  yHighest <- y$stats[5, columnNum]
  
  xLower  <- x$stats[2, columnNum]
  xHigher <- x$stats[4, columnNum]
  yLower  <- y$stats[2, columnNum]
  yHigher <- y$stats[4, columnNum]
  
  xCenter  <- x$stats[3, columnNum]
  yCenter  <- y$stats[3, columnNum]
  
  if(is.null(colorSheer)) {
    colorSheer <- paste(substring(color, 1, 7), "33", sep="")
  }
  
  hasX <- !is.na(xCenter)
  hasY <- !is.na(yCenter)

  if(verbose) {
    print(c("column", columnNum, columnChar))
    print(c("color", color, colorSheer))
    print(c("x", xLowest, xLower, xCenter, xHigher, xHighest))
    print(c("y", yLowest, yLower, yCenter, yHigher, yHighest))
    print(c("has data", hasX, hasY))
  }

  # draw factor character on top and right axis
  if(nameOnAxis) {
    if(hasX) mtext(columnChar,side=3,at=xCenter,col=color)
    if(hasY) mtext(columnChar,side=4,at=yCenter,col=color)
  }

  # both X and Y are required to draw followings
  if(!hasX || !hasY) return(FALSE)

  # draw a box of 2nd and 3rd quantiles
  rect(xLower, yLower, xHigher, yHigher, col=colorSheer)
 
  # draw whiskers for 1st and 4th quantiles
  # bosedWhiskers=TRUE draws a large box as whiskers

  if(boxedWhiskers) {
    xBarLow  <- xLowest
    xBarHigh <- xHighest
    yBarLow  <- yLowest
    yBarHigh <- yHighest
  } else {
    xBarLow  <- xLower
    xBarHigh <- xHigher
    yBarLow  <- yLower
    yBarHigh <- yHigher
  }
  
  segments(xLowest, yBarLow, xLowest, yBarHigh, col=color)
  segments(xHighest, yBarLow, xHighest, yBarHigh, col=color)
  segments(xBarLow, yLowest, xBarHigh, yLowest, col=color)
  segments(xBarLow, yHighest, xBarHigh, yHighest, col=color)
  
  segments(xLowest, yCenter, xHighest, yCenter, col=color)
  segments(xCenter, yLowest, xCenter, yHighest, col=color)
 
  # draw outliers

  xOut <- x$out
  xOutGroup <- x$group
  yOut <- y$out
  yOutGroup <- y$group
  
  if(verbose) {
    print(c("x out", xOutGroup,xOut))
    print(c("y out", yOutGroup,yOut))
  }
    
  for(x in xOut[xOutGroup==columnNum]) points(x, yCenter, col=color, pch=1, cex=2)
  for(y in yOut[yOutGroup==columnNum]) points(xCenter, y, col=color, pch=1, cex=2)
 
  # outliersHasWhiskers=TRUE add whiskers at each outlier

  if(outliersHasWhiskers) {
    for(x in xOut[xOutGroup==columnNum]) {
      segments(x, yCenter, xCenter, yCenter, col=colorSheer)
      segments(x, yLower, x, yHigher, col=colorSheer)
    }
    for(y in yOut[yOutGroup==columnNum]) {
      segments(xCenter, y, xCenter, yCenter, col=colorSheer)
      segments(xLower, y, xHigher, y, col=colorSheer)
    }
  }
 
  # draw the center as factor character
  text(xCenter, yCenter, columnChar)
  
  return(TRUE)
}

