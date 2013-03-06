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
   
    if(!condense) return(list(result=FALSE))

    overlap <- has.overlap.stat(stat, severity=severity)
    if(verbose) print(overlap)
    overlap
}

bxpdou.exec.condense <- function(x, y, overlap, verbose=FALSE) {

  fs <- list(x=x[,1], y=y[,1])

  for(i in 1L:overlap$n) {
    pair <- overlap$names[c(overlap$col[i], overlap$row[i])]
    condensed.char <- paste(pair[1], pair[2], sep="+")
    #condensed.level <- c(overlap$names, condensed.char)
    condensed.level <- c(unique(as.character(fs$x)), condensed.char)

    if(verbose) {
      print(i)
      print(pair)
      print(condensed.char)
      print(condensed.level)
    }

    fs <- lapply(fs, 
                 function(f) {
                   f <- factor(f, levels=condensed.level)
                   str(f)
                   f[!is.na(match(f, pair))] <- condensed.char
                   str(f)
                   f
                 })
    if(verbose) {
      str(fs)
    }
#    for(f in fs) {
#      f <- factor(f, levels=condensed.level)
#      f[!is.na(match(f, pair))] <- condensed.char
#    }
#    x[,1] <- factor(x[,1], levels=condensed.level)
#    x[,1][!is.na(match(x[,1], pair))] <- condensed.char
#    y[,1] <- factor(y[,1], levels=condensed.level)
#    y[,1][!is.na(match(y[,1], pair))] <- condensed.char
  }

  result <- list(x=x, y=y)
  result$x[,1] <- fs$x
  result$y[,1] <- fs$y
  result
}
