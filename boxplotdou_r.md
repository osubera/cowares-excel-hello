

# Introduction #

  * double boxplot (two axes boxplot) for R

## 概要 ##
  * Rで二重箱ひげ図 (2軸箱ひげ図) を描く

# Details #

## 説明 ##

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=boxplotdou_r)
  * [download from developer source tree / 開発用レポジトリの最新のソース](http://code.google.com/p/cowares-excel-hello/source/browse/trunk/boxplotdou/)

# How to use #

  * [how to use the double box plot](http://tomizonor.wordpress.com/2013/03/15/double-box-plot/)

## 使い方 ##

  * [使い方の解説 ](http://cowares.blogspot.jp/2013/03/r-2.html)

# Snapshots #

**Double box plot chart shows both the (x, y) correlation and the distribution of each data.
> ![http://2.bp.blogspot.com/-BIrTSHxgsp0/UTvZpl1Dk8I/AAAAAAAAANQ/yBUc-0l1Yj0/s1600/fig9.png](http://2.bp.blogspot.com/-BIrTSHxgsp0/UTvZpl1Dk8I/AAAAAAAAANQ/yBUc-0l1Yj0/s1600/fig9.png)**

# Code #

### boxplotdou.r ###

```
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
           color=NULL, color.sheer=NULL, 
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=TRUE,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           pars=NULL, verbose=FALSE, plot=TRUE, ...) {

  # both x and y expect data frame with 2 columns:  factor, observation

  #stat <- bxpdou.stats(x, y, verbose)
  stat <- bxpdou.stats.condense(x, y,
                                condense=condense, 
                                severity=condense.severity, 
                                once=condense.once,
                                verbose=verbose)

  if(plot) {
    bxpdou(stat$stat$x, stat$stat$y, stat$level,
           xlab=stat$name$x, ylab=stat$name$y, 
           pars=par(), color=color, color.sheer=color.sheer, 
           boxed.whiskers=boxed.whiskers, 
           outliers.has.whiskers=outliers.has.whiskers,
           name.on.axis=name.on.axis, 
           verbose=verbose, ...)
    invisible(stat$stat)
  } else {
    stat$stat
  }
}

bxpdou <- 
function(x.stats, y.stats, factor.levels, 
         color=NULL, color.sheer=NULL, 
         boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, name.on.axis=TRUE, 
         pars=NULL, verbose=FALSE, ...) {
  
  pars <- c(list(...), pars)
  # the first overrides the later
  pars <- pars[unique(names(pars))]

  x.min <- min(x.stats$stats, x.stats$out, na.rm=TRUE)
  x.max <- max(x.stats$stats, x.stats$out, na.rm=TRUE)
  y.min <- min(y.stats$stats, y.stats$out, na.rm=TRUE)
  y.max <- max(y.stats$stats, y.stats$out, na.rm=TRUE)

  if(is.null(pars$xlim)) xlim <- c(x.min, x.max)
  if(is.null(pars$ylim)) ylim <- c(y.min, y.max)

  levels.num <- length(factor.levels)
  levels.col <- rainbow(levels.num)
  ##FIXME color and color.sheer is not used.
  
  if(verbose) {
    print(list(xlim=xlim, ylim=ylim))
  }

  # open a plot area and draw axis
  plot(NULL, xlim=xlim, ylim=ylim, ...)

  # draw boxes
  for(i in 1L:levels.num) {
    bxpdou.abox(x.stats, y.stats, 
                column.num=i, column.char=as.character(factor.levels)[i], 
                color=levels.col[i], color.sheer=NULL, name.on.axis=name.on.axis, 
                boxed.whiskers=boxed.whiskers, outliers.has.whiskers=outliers.has.whiskers,
                verbose=verbose)
  }
}

bxpdou.abox <- 
function(x, y, column.num, column.char, 
         color, color.sheer=NULL, 
         boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
         name.on.axis=TRUE, 
         verbose=FALSE) {

  # declare five numbers explicitly

  x.lowest  <- x$stats[1, column.num]
  x.highest <- x$stats[5, column.num]
  y.lowest  <- y$stats[1, column.num]
  y.highest <- y$stats[5, column.num]
  
  x.lower  <- x$stats[2, column.num]
  x.higher <- x$stats[4, column.num]
  y.lower  <- y$stats[2, column.num]
  y.higher <- y$stats[4, column.num]
  
  x.center  <- x$stats[3, column.num]
  y.center  <- y$stats[3, column.num]
  
  if(is.null(color.sheer)) {
    color.sheer <- paste(substring(color, 1, 7), "33", sep="")
  }
  
  has.x <- !is.na(x.center)
  has.y <- !is.na(y.center)

  if(verbose) {
    print(c("column", column.num, column.char))
    print(c("color", color, color.sheer))
    print(c("x", x.lowest, x.lower, x.center, x.higher, x.highest))
    print(c("y", y.lowest, y.lower, y.center, y.higher, y.highest))
    print(c("has data", has.x, has.y))
  }

  # draw factor character on top and right axis
  if(name.on.axis) {
    if(has.x) mtext(column.char,side=3,at=x.center,col=color)
    if(has.y) mtext(column.char,side=4,at=y.center,col=color)
  }

  # both X and Y are required to draw followings
  if(!has.x || !has.y) return(FALSE)

  # draw a box of 2nd and 3rd quantiles
  rect(x.lower, y.lower, x.higher, y.higher, col=color.sheer)
 
  # draw whiskers for 1st and 4th quantiles
  # bosed.whiskers=TRUE draws a large box as whiskers

  if(boxed.whiskers) {
    x.bar.low  <- x.lowest
    x.bar.high <- x.highest
    y.bar.low  <- y.lowest
    y.bar.high <- y.highest
  } else {
    x.bar.low  <- x.lower
    x.bar.high <- x.higher
    y.bar.low  <- y.lower
    y.bar.high <- y.higher
  }
  
  segments(x.lowest, y.bar.low, x.lowest, y.bar.high, col=color)
  segments(x.highest, y.bar.low, x.highest, y.bar.high, col=color)
  segments(x.bar.low, y.lowest, x.bar.high, y.lowest, col=color)
  segments(x.bar.low, y.highest, x.bar.high, y.highest, col=color)
  
  segments(x.lowest, y.center, x.highest, y.center, col=color)
  segments(x.center, y.lowest, x.center, y.highest, col=color)
 
  # draw outliers

  x.out <- x$out
  x.out.group <- x$group
  y.out <- y$out
  y.out.group <- y$group
  
  if(verbose) {
    print(c("x out", x.out.group,x.out))
    print(c("y out", y.out.group,y.out))
  }
    
  for(x in x.out[x.out.group==column.num]) points(x, y.center, col=color, pch=1, cex=2)
  for(y in y.out[y.out.group==column.num]) points(x.center, y, col=color, pch=1, cex=2)
 
  # outliers.has.whiskers=TRUE add whiskers at each outlier

  if(outliers.has.whiskers) {
    for(x in x.out[x.out.group==column.num]) {
      segments(x, y.center, x.center, y.center, col=color.sheer)
      segments(x, y.lower, x, y.higher, col=color.sheer)
    }
    for(y in y.out[y.out.group==column.num]) {
      segments(x.center, y, x.center, y.center, col=color.sheer)
      segments(x.lower, y, x.higher, y, col=color.sheer)
    }
  }
 
  # draw the center as factor character
  text(x.center, y.center, column.char)
  
  return(TRUE)
}

bxpdou.stats <- function(x, y, verbose=FALSE) {

  # both x and y expect data frame with 2 columns:  factor, observation

  data <- list(x=x[,2], y=y[,2])
  name <- list(x=names(x[2]), y=names(y[2]))
  stat <- list()

  # merge and unite the factor levels of x and y
  factor.levels <- unique(c(unique(as.character(x[,1])),
                            unique(as.character(y[,1]))))
  factor <- list(x=factor(x[,1],levels=factor.levels), 
                 y=factor(y[,1],levels=factor.levels))

  for(v in c("x","y")) {
    # delegate the calculation to the standard boxplot
    stat[[v]] <- boxplot(formula=
                         formula(paste("data$", v, "~factor$", v, sep="")), 
                         plot=FALSE)
    if(is.null(name[[v]])) name[[v]] <- v
  }

  if(verbose) {
    print(name)
    print(data)
    print(stat)
  }

  list(stat=stat, name=name, level=factor.levels)
}

bxpdou.stats.condense <- 
  function(x, y,
           condense=FALSE, severity="iqr", once=FALSE,
           verbose=FALSE) {
    
    stat <- bxpdou.stats(x, y, verbose=verbose)
    to.be.condensed <- 
      bxpdou.check.condense(stat=stat$stat,
                            condense=condense, severity=severity,
                            verbose=verbose)

    if(to.be.condensed$result) {
      condensed <- 
        bxpdou.exec.condense(x, y,
                             overlap=to.be.condensed, verbose=verbose)
      return(Recall(condensed$x, condensed$y,
                    condense=condense && !once, severity=severity,
                    verbose=verbose))
    } else {
      return(stat)
    }
}

bxpdou.check.condense <- 
  function(stat,
           condense=FALSE, severity="iqr", 
           verbose=FALSE) {
   
    if(!condense || length(stat$x$names)<=1) return(list(result=FALSE))

    overlap <- has.overlap.stat(stat, severity=severity)
    if(verbose) print(overlap)
    overlap
}

bxpdou.exec.condense <- function(x, y, overlap, verbose=FALSE) {

  condensed.level <- current.level <- overlap$names
  level.indexes <- 1L:length(current.level)

  for(i in 1L:overlap$n) {
    location <- c(overlap$col[i], overlap$row[i])
    pair <- condensed.level[location]
    if(pair[1]==pair[2]) next

    expanded.location <- 
      level.indexes[!is.na(match(condensed.level, pair))]
    condensed.char <- paste(pair, collapse="+")
    condensed.level[expanded.location] <- condensed.char

    if(verbose) {
      print(list(pair=pair, location=location, expanded=expanded.location))
    }
  }

  if(verbose) {
    print(list(current=current.level, condensed=condensed.level))
  }

  current.fs <- list(x=x[,1], y=y[,1])
  condensed.fs <- 
    lapply(current.fs, 
           function(f) {
             factor(condensed.level[match(f, current.level)],
                    levels=unique(condensed.level))
           })
  if(verbose) {
    print(mapply(function(current, condensed) {
                   cbind(as.character(current), current, 
                         as.character(condensed), condensed)
                 }, 
                 current.fs, condensed.fs))
  }

  result <- list(x=x, y=y)
  result$x[,1] <- condensed.fs$x
  result$y[,1] <- condensed.fs$y

  if(verbose) {
    print(result)
  }
  result
}

```

### test.boxplotdou.r ###

```
test1.data.independent <- function() {
  nf <- 5
  nx <- 100
  ny <- 50
  obs.x <- rnorm(nx, mean=10, sd=3)
  obs.y <- rnorm(ny, mean=30, sd=8)
  factor.x <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=nx/nf)]
  factor.y <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=ny/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test3.data.linear <- function() {
  nf <- 5
  nx <- 100
  obs.x <- rnorm(nx, mean=10, sd=3)
  obs.y <- obs.x + rnorm(nx, mean=0, sd=0.1)
  factor.y <- factor.x <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=nx/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test4.data.linear <- function() {
  nf <- 4
  nx <- 100
  obs.x <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  obs.y <- obs.x + rnorm(nx, mean=0, sd=0.1)
  factor.y <- factor.x <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=nx/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test5.data.independent <- function() {
  nf <- 4
  nx <- 100
  ny <- 40
  obs.x <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  obs.y <- rnorm(ny, mean=30, sd=8)
  factor.x <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=nx/nf)]
  factor.y <- strsplit("abcdefghijklmnopqrstuvwxyz",split=character(0))[[1]][rep(1:nf, each=ny/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test14.data.single <- function() {
  nx <- 100
  ny <- 50
  obs.x <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  obs.x <- c(rnorm(nx-1, mean=100, sd=10),147)
  obs.y <- c(rnorm(ny-2, mean=30, sd=5),45,48)
  factor.x <- rep("o", nx)
  factor.y <- rep("o", ny)
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test15.data.independent <- function() {
  n <- 10
  obs.x <- c(rnorm(n, mean=120, sd=18), rnorm(n, mean=100, sd=12), rnorm(n, mean=100, sd=8))
  obs.y <- c(rnorm(n, mean=30, sd=8), rnorm(n, mean=40, sd=5), rnorm(n, mean=32, sd=10))
  factor.y <- factor.x <- rep(c("a","b","c"), each=n)
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test1 <- function() {
  testdata <- test1.data.independent()
  boxplotdou(testdata$x,testdata$y,plot=F,verbose=T)
}

test2 <- function() {
  testdata <- test1.data.independent()
  boxplotdou(testdata$x,testdata$y)
}

test3 <- function() {
  testdata <- test3.data.linear()
  boxplotdou(testdata$x,testdata$y)
}

test4 <- function() {
  testdata <- test4.data.linear()
  boxplotdou(testdata$x,testdata$y)
}

test5 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y)
}

test6 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,boxed.whiskers=T)
}

test7 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,outliers.has.whiskers=T)
}

test8 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,name.on.axis=F)
}

test9 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,verbose=T)
}

test10 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,verbose=T,
             condense=T)
}

test11 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,verbose=T,
             condense=T,condense.severity="whisker")
}

test12 <- function() {
  testdata <- test4.data.linear()
  boxplotdou(testdata$x,testdata$y,verbose=T,
             condense=T,condense.severity="iqr")
}

test13 <- function() {
  testdata <- test5.data.independent()
  boxplotdou(testdata$x,testdata$y,name.on.axis=F)
}

test14 <- function() {
  testdata <- test14.data.single()
  boxplotdou(testdata$x,testdata$y,name.on.axis=F)
}

test15 <- function(...) {
  testdata <- test15.data.independent()
  boxplotdou(testdata$x,testdata$y,...)
}

```

### has.overlap.r ###

```
# has overlap
# http://code.google.com/p/cowares-excel-hello/source/browse/trunk/boxplotdou/
#
# Copyright (C) 2013 Tomizono
# Fortitudinous, Free, Fair, http://cowares.nobody.jp

# rect <- c(xleft, ybottom, xright, ytop)
# contacts are not counted to overlaps.

is.recta.apart.right <- function(recta, rectb) {
  rectb[3] < recta[1]
}

is.recta.apart.left <- function(recta, rectb) {
  recta[3] < rectb[1]
}

is.recta.apart.top <- function(recta, rectb) {
  rectb[4] < recta[2]
}

is.recta.apart.bottom <- function(recta, rectb) {
  recta[4] < rectb[2]
}

has.overlap.rect <- function(recta, rectb) {
  !is.recta.apart.right(recta, rectb) &&
  !is.recta.apart.left(recta, rectb) &&
  !is.recta.apart.top(recta, rectb) &&
  !is.recta.apart.bottom(recta, rectb)
}

has.overlap.xory.rect <- function(recta, rectb) {
  (!is.recta.apart.right(recta, rectb) &&
  !is.recta.apart.left(recta, rectb)) ||
  (!is.recta.apart.top(recta, rectb) &&
  !is.recta.apart.bottom(recta, rectb))
}

# stats is an output of boxplotdou(...,plot=F)
# factor is one of column numbers of stats
# severity is one of,
#    iqr : Inter Quartile Range
#    whisker : Inter Whiskers
#
# $stats example
#           [,1]      [,2]     [,3]     [,4]
# [1,]  8.267469  4.635134 17.31795 16.27995
# [2,]  9.910078  7.193855 18.70520 18.09055
# [3,] 10.616723  8.886337 19.64145 18.97840
# [4,] 11.619414 11.410812 20.77409 19.46167
# [5,] 13.109597 14.787117 23.38350 21.34703

as.rect.of.stats <- function(stats, factor, severity="iqr") {
  row.minmax <- switch(severity, 
                       c(2,4),    # "iqr" is default
                       "whisker"=c(1,5)
                       )
  c(stats$x$stats[row.minmax[1], factor], 
    stats$y$stats[row.minmax[1], factor], 
    stats$x$stats[row.minmax[2], factor], 
    stats$y$stats[row.minmax[2], factor]
    )
}

has.overlap.factor <- function(stats, factora, factorb, severity="iqr") {
  rules <- strsplit(severity, "\\.")[[1]]
  switch(rules[2],
         has.overlap.rect(as.rect.of.stats(stats, factora, rules[1]),
                          as.rect.of.stats(stats, factorb, rules[1])),
         "xory"=
         has.overlap.xory.rect(as.rect.of.stats(stats, factora, rules[1]),
                               as.rect.of.stats(stats, factorb, rules[1]))
         )
}

has.overlap.stat <- function(stats, severity="iqr") {
  columns.num <- dim(stats$x$stats)[2]
  columns <- 1L:columns.num

  row.indexes <- rep(columns, times=columns.num)
  col.indexes <- rep(columns, each=columns.num)
  result.matrix <- matrix(nrow=columns.num, ncol=columns.num)

  for(r in columns) for(c in columns) {
    if(c<r) result.matrix[r,c] <- has.overlap.factor(stats, r, c, severity)
  }

  locate.true <- which(result.matrix)
  count.true <- length(locate.true)

  list(result=(count.true!=0),
       n=count.true,
       row=row.indexes[locate.true],
       col=col.indexes[locate.true],
       overlap=result.matrix,
       names=stats$x$names
       )
}

```

### test.has.overlap.r ###

```
test1 <- function() {
  has.overlap.rect(c(1,1,3,4),c(-5,1,-3,4)) == F
}

test2 <- function() {
  has.overlap.rect(c(1,1,3,4),c(15,1,23,4)) == F
}

test3 <- function() {
  has.overlap.rect(c(1,1,3,4),c(1,-6,3,-4)) == F
}

test4 <- function() {
  has.overlap.rect(c(1,1,3,4),c(1,11,3,14)) == F
}

test5 <- function() {
  has.overlap.rect(c(1,1,3,4),c(1,1,3,4)) == T
}

test6 <- function() {
  has.overlap.rect(c(1,1,3,4),c(2,2,4,5)) == T
}

test7 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-5,1,-3,4)) == T
}

test8 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(15,1,23,4)) == T
}

test9 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,-6,3,-4)) == T
}

test10 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,11,3,14)) == T
}

test11 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,1,3,4)) == T
}

test12 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(2,2,4,5)) == T
}

test13 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(12,12,14,15)) == F
}

test14 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(12,-8,14,-5)) == F
}

test15 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-7,12,-4,15)) == F
}

test16 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-7,-8,-4,-5)) == F
}

```
