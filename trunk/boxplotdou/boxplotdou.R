# double boxplot
# http://code.google.com/p/cowares-excel-hello/source/browse/trunk/boxplotdou/
#
# Copyright (C) 2013 Tomizono
# Fortitudinous, Free, Fair, http://cowares.nobody.jp
#                            http://paidforeveryone.wordpress.com/

# boxplotdou : S3 Method
#  default = data.frame
#  factor

# boxplotdou.data.frame(cbind(factor1, data1), cbind(factor2, data2))
# boxplotdou.factor(factor1, data1, factor2, data2)

# boxplotdou.data.frame(
#  x = data.frame, factor and observation to x-axis
#  y = data.frame, factor and observation to y-axis
#  boxed.whiskers = logical, default=FALSE,
#                   TRUE to draw rectangular range rather than whisker
#  outliers.has.whiskers = logical, default=FALSE, 
#                          extend whisker and staple through outliers
#  name.on.axis = logical, default=TRUE, 
#                          TRUE to draw group names on axes
#  factor.labels = control labels on each group on factor
#                  default=NULL, using factor data
#                  TRUE to abbreviate by alphabet letters
#                  FALSE to draw no labels
#                  character vector to give explicit labels
#                  single character to use identical character
#                  NA in vector to exclude any groups
#  draw.legend = logical, draw legend or not
#                default=NA, enable legend only when labels abbreviated
#  condense = logical, default=FALSE, 
#             TRUE to unify near groups into one box
#  condense.severity = character, default="iqr",
#                      one of c('iqr','whisker','iqr.xory','whisker.xory'),
#                      which is the border to condense or not,
#                      used only when condense=TRUE
#  condense.once = logical, default=FALSE,
#                  TRUE to disable recursive condenses,
#                  used only when condense=TRUE
#  col = colors for each group
#        default=NULL, automatic colors
#  COLOR.SHEER = function, to convert color to sheer color,
#                default=bxpdou.sheer.color, internally defined,
#                sheer colors are used for inside box and outlier-whisker
#  shading = shading density to draw inside of box,
#            default=NA is automatic, usually no shadings
#  shading.angle = shading angle to draw inside of box,
#                  default=NA is automatic, usually no shadings
#  blackwhite = logical, default=FALSE,
#               TRUE to draw black and white chart,
#               equivalent to set following 3 parameters,
#                 col='black', shading=TRUE,
#                 COLOR.SHEER=(function(a) a)
#  verbose = logical, default=FALSE, TRUE is to show debug information
#  plot = logical, default=TRUE, to draw a chart
#  ... = accepts graphical parameters and boxplot color parameters

# boxplotdou.factor(
#  f.x = factor vector to x-axis
#  obs.x = observation vector to x-axis
#  f.y = factor vector to y-axis
#  obs.y = observation vector to y-axis

# boxplot color parameters
#
#  medcol = default=NULL, is black, colors for median labels
#  whiskcol = default=NULL, is =col, colors for whiskers
#  staplecol = default=NULL, is =col, colors for staples
#  boxcol = default=NULL, is black, colors for box borders
#  outcol = default=NULL, is =col, colors for outliers
#  outbg = default=NULL, is transparent, colors inside outliers
#  outcex = default=2, size of outliers
#  outpch = default=1, is a transparent circle, symbol number of outliers

# dependencies
#
# using boxplot to calculate boxplot statistics

# data structure (output values)
#
# list of 2 items (x, y)
# each item is identical to boxplot statistics


boxplotdou <- function(x, ...) UseMethod("boxplotdou")

boxplotdou.data.frame <- 
  function(x, y, 
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=TRUE, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           verbose=FALSE, plot=TRUE, ...) {

  # both x and y expect data frame with 2 columns:  factor, observation

  stat <- bxpdou.stats.condense(x, y,
                                condense=condense, 
                                severity=condense.severity, 
                                once=condense.once,
                                verbose=verbose)

  if(plot) {
    if(blackwhite) {
      if(is.null(col)) col <- 'black'
      COLOR.SHEER <- function(a) a
      if(is.na(shading) && is.na(shading.angle)) shading=TRUE
    }

    bxpdou(stat$stat$x, stat$stat$y, stat$level,
           xlab=stat$name$x, ylab=stat$name$y, 
           boxed.whiskers=boxed.whiskers, 
           outliers.has.whiskers=outliers.has.whiskers,
           name.on.axis=name.on.axis,
           factor.labels=factor.labels, draw.legend=draw.legend,
           col=col,
           COLOR.SHEER=COLOR.SHEER, 
           shading.density=shading, shading.angle=shading.angle,
           verbose=verbose, ...)
    invisible(stat$stat)
  } else {
    stat$stat
  }
}

boxplotdou.factor <- 
  function(f.x, obs.x, f.y, obs.y,  
           boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
           name.on.axis=TRUE, factor.labels=NULL, draw.legend=NA,
           condense=FALSE, condense.severity="iqr",
           condense.once=FALSE,
           col=NULL,
           COLOR.SHEER=bxpdou.sheer.color, 
           shading=NA, shading.angle=NA, blackwhite=FALSE,
           verbose=FALSE, plot=TRUE, ...) {

  # f.x and f.y are factor vectors
  # obs.x and obs.y are observation vectors
  x <- data.frame(f=f.x, x=obs.x)
  y <- data.frame(f=f.y, y=obs.y)
  boxplotdou.data.frame(x, y,
           boxed.whiskers=boxed.whiskers,
           outliers.has.whiskers=outliers.has.whiskers,
           name.on.axis=name.on.axis, factor.labels=factor.labels,
           draw.legend=draw.legend, condense=condense,
           condense.severity=condense.severity, condense.once=condense.once,
           col=col, COLOR.SHEER=COLOR.SHEER, shading=shading, 
           shading.angle=shading.angle, blackwhite=blackwhite,
           verbose=verbose, plot=plot, ...) 
}

boxplotdou.default <- boxplotdou.data.frame


bxpdou <- 
function(x.stats, y.stats, factor.levels, 
         boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
         name.on.axis=TRUE, factor.labels=NULL, draw.legend=NA,
         col=NULL,
         COLOR.SHEER=bxpdou.sheer.color,
         shading.density=NA, shading.angle=NA,
         verbose=FALSE, ...) {
  
  xlim <- range(x.stats$stats, x.stats$out, na.rm=TRUE)
  ylim <- range(y.stats$stats, y.stats$out, na.rm=TRUE)

  levels.num <- length(factor.levels)
  levels.col <- 
    if(is.null(col)) { rainbow(levels.num)
    } else { rep(col, levels.num) }
  levels.col.sheer <- COLOR.SHEER(levels.col)

  levels.shade <- 
    make.shadings(n=levels.num, density=shading.density, 
                  angle=shading.angle, verbose=verbose)

  use.strict.factor.labels <- is.null(factor.labels)
  factor.labels <-
    if(length(factor.labels) == 1) rep(factor.labels, levels.num) else
      factor.labels[1L:levels.num]

  abbr.factor.labels <-
    if(use.strict.factor.labels) {
      rep('s', levels.num) # strict
    } else if(is.logical(factor.labels)) {
      ifelse(factor.labels, 
             'g',          # generates internally
             'n')          # no labels
    } else {
      ifelse(is.na(factor.labels), 
             NA,           # ignore, not draw
             'e')          # specified explicitly
    }
  if(levels.num > 26) {
    abbr.is.g <- abbr.factor.labels %in% 'g'
    abbr.is.g[1:26] <- FALSE
    abbr.factor.labels[abbr.is.g] <- 's'
  }

  column.char <- 
    sapply(as.list(1L:levels.num),
           function(i) switch(abbr.factor.labels[i],
                              s=as.character(factor.levels[i]),
                              g=letters[i],
                              e=as.character(factor.labels[i]),
                              '')
           )

  if(is.na(draw.legend))
    draw.legend <- !use.strict.factor.labels

  medcol <- rep(extract.from.dots('medcol', NULL, ...), levels.num)
  whiskcol <- rep(extract.from.dots('whiskcol', NULL, ...), levels.num)
  staplecol <- rep(extract.from.dots('staplecol', NULL, ...), levels.num)
  boxcol <- rep(extract.from.dots('boxcol', NULL, ...), levels.num)
  outcol <- rep(extract.from.dots('outcol', NULL, ...), levels.num)
  outbg <- rep(extract.from.dots('outbg', NULL, ...), levels.num)
  outcex <- rep(extract.from.dots('outcex', 2, ...), levels.num)
  outpch <- rep(extract.from.dots('outpch', 1, ...), levels.num)

  if(is.null(whiskcol)) whiskcol <- levels.col
  if(is.null(staplecol)) staplecol <- levels.col
  if(is.null(outcol)) outcol <- levels.col

  if(verbose) {
    print(list(factor.labels=factor.labels,
               abbr.factor.labels=abbr.factor.labels,
               xlim=xlim, ylim=ylim))
  }

  # open a plot area and draw axis
  do.call.with.par('plot.default', list(...), x=NA, xlim=xlim, ylim=ylim)

  # legend for abbreviates at top
  if(draw.legend) {
    not.na <- !is.na(abbr.factor.labels)
    abbr <- column.char[not.na]
    strict <- factor.levels[not.na]
    legend.col <- levels.col[not.na]

    legend <- paste(abbr, strict, sep=': ')

    n <- length(legend)
    pt <- seq(from=xlim[1], to=xlim[2], length.out=n)
    for(i in 1L:n)
      do.call.with.par('mtext', list(...), dots.win=T,
                       text=legend[i], side=3, line=1, at=pt[i],
                       col=legend.col[i])
  }

  # draw boxes
  for(i in 1L:levels.num) {
    if(!is.na(abbr.factor.labels[i]))
      bxpdou.abox(x.stats, y.stats, 
                  column.num=i, column.char=column.char[i],
                  color=levels.col[i], color.sheer=levels.col.sheer[i],
                  density=levels.shade$density[i],
                  angle=levels.shade$angle[i],
                  name.on.axis=name.on.axis, 
                  boxed.whiskers=boxed.whiskers, 
                  outliers.has.whiskers=outliers.has.whiskers,
                  median.color=medcol[i],
                  whisker.color=whiskcol[i],
                  staple.color=staplecol[i],
                  box.color=boxcol[i],
                  outlier.color=outcol[i],
                  outlier.bgcolor=outbg[i],
                  outlier.cex=outcex[i],
                  outlier.pch=outpch[i],
                  verbose=verbose, ...)
  }
}

bxpdou.sheer.color <- function(col) {
  adjustcolor(col, alpha.f=0.2)
}

bxpdou.abox <- 
function(x, y, column.num, column.char, 
         color, color.sheer, 
         density=NULL, angle=NULL,
         boxed.whiskers=FALSE, outliers.has.whiskers=FALSE, 
         name.on.axis=TRUE, 
         median.color=NULL,
         whisker.color=NULL,
         staple.color=NULL,
         box.color=NULL,
         outlier.color=NULL,
         outlier.bgcolor=NULL,
         outlier.cex=2,
         outlier.pch=1,
         verbose=FALSE, ...) {

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
    if(has.x) 
      do.call.with.par('mtext', list(...), dots.win=T, 
                       text=column.char, side=3, at=x.center, col=color)
    if(has.y) 
      do.call.with.par('mtext', list(...), dots.win=T,
                       text=column.char, side=4, at=y.center, col=color)
  }

  # both X and Y are required to draw followings
  if(!has.x || !has.y) return(FALSE)

  # draw a box of 2nd and 3rd quantiles
  rect(x.lower, y.lower, x.higher, y.higher, col=color.sheer, density=density, angle=angle, border=box.color)
 
  # draw whiskers for 1st and 4th quantiles
  # boxed.whiskers=TRUE draws a large box as staples

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
  
  segments(x.lowest, y.bar.low, x.lowest, y.bar.high, col=staple.color)
  segments(x.highest, y.bar.low, x.highest, y.bar.high, col=staple.color)
  segments(x.bar.low, y.lowest, x.bar.high, y.lowest, col=staple.color)
  segments(x.bar.low, y.highest, x.bar.high, y.highest, col=staple.color)
  
  segments(x.lowest, y.center, x.highest, y.center, col=whisker.color)
  segments(x.center, y.lowest, x.center, y.highest, col=whisker.color)
 
  # draw outliers

  x.out <- x$out
  x.out.group <- x$group
  y.out <- y$out
  y.out.group <- y$group
  
  if(verbose) {
    print(c("x out", x.out.group,x.out))
    print(c("y out", y.out.group,y.out))
  }
    
  for(x in x.out[x.out.group==column.num]) 
    points(x, y.center, col=outlier.color, bg=outlier.bgcolor,
           pch=outlier.pch, cex=outlier.cex)
  for(y in y.out[y.out.group==column.num]) 
    points(x.center, y, col=outlier.color, bg=outlier.bgcolor,
           pch=outlier.pch, cex=outlier.cex)
 
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
  do.call.with.par('text', list(...), dots.win=T,
                   x=x.center, y=y.center, labels=column.char,
                   col=median.color)
  
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

make.shadings <- function(n, density=NA, angle=NA, verbose=FALSE) {
  shadings <- list(density=NULL, angle=NULL)

  if(is.na(density) && is.na(angle)) return(shadings)

  label <- c('density', 'angle')
  start <- c(12, 10)
  end <- c(36, 160)
  shuffling <- c(TRUE, FALSE)
  par <- list(density, angle)

  for(i in 1:2) {
    x <- par[[i]]

    if(length(x) >= n) {
      shadings[[i]] <- x[1L:n]
    } else {
      rx <- if(is.numeric(x)) rep(x, 2)[1:2] else rep(NA, 2)
      if(is.na(rx[1])) rx[1] <- start[i]
      if(is.na(rx[2])) rx[2] <- end[i]

      sx <- seq(from=rx[1], to=rx[2], length.out=n)

      if(shuffling[i]) {
        # shuffle by halves
        # 1, k, 2, k+1, 3, k+2,,, n
        shuffle <- seq(from=1, to=n, by=2)
        shuffle <- c(shuffle, shuffle + 1)[1L:n]
      } else {
        shuffle <- 1L:n
      }
      shadings[[i]] <- sx[order(shuffle)]
    }
  }

  if(verbose) {
    cat('# shading enabled\n')
    print(shadings)
  }

  shadings
}

only.graphic.pars <- function(pars, what='plot.default') {
  if(length(pars) == 0) return(list())

  gnames <- names(par(no.readonly=T))
  pnames <- names(formals(args(what)))
  gpnames <- unique(c(pnames[pnames != '...'], gnames))

  selectgname <- pars[gpnames]
  to.rm.na <- names(selectgname)

  if(is.null(to.rm.na)) list() else 
    selectgname[!is.na(to.rm.na)]
}

do.call.with.par <- function(what, args, dots.win=FALSE, ...) {
  pars <- 
    if(dots.win) {
      modifyList(only.graphic.pars(args, what), list(...))
    } else {
      modifyList(list(...), only.graphic.pars(args, what))
    }
  do.call(what, pars)
}

extract.from.dots <- function(which, default=NULL, ...) {
  x <- list(...)[[which]]
  if(is.null(x)) default else x
}

