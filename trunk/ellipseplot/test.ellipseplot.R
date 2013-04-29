# test functions

testdata.onebox <- data.frame(low=1:0,center=3:4,high=6:5,row.names=c('x','y'))

test.anellipse <- function(x=testdata.onebox, verbose=F) {
  plot(t(x['x',]), t(x['y',]))
  anellipse(x, verbose, col='red', border='blue', lty='dotted')
}

test.ninenum <- function() {
  print(ninenum(1:9))
  print(ninenum(1:999))
  print(ninenum(1:1000))
  print(ninenum(c(9:1,NA)))
  print(ninenum(rep(NA,9)))
  invisible(T)
}

test.manyellipses <- function(...) {
  SUMMARY=ninenum
  stats <- list(
                data.frame(x=SUMMARY(rnorm(10)), y=SUMMARY(rnorm(10))),
                data.frame(x=SUMMARY(rnorm(10,1)), y=SUMMARY(rnorm(10))),
                data.frame(x=SUMMARY(rnorm(10,2)), y=SUMMARY(rnorm(10,1))),
                data.frame(x=SUMMARY(rnorm(10,3)), y=SUMMARY(rnorm(10,1))),
                data.frame(x=SUMMARY(rnorm(10,4)), y=SUMMARY(rnorm(10,4)))
                )
  many.ellipses(stats, list(x='x', y='y'), ...) 
}

test.ellipseplot.single <- function(n=10, SUMMARY=ninenum, 
                                  plot=T, verbose=F, ...) {
  x <- rnorm(n)
  y <- rnorm(n)
  ellipseplot.single(x, y, SUMMARY=SUMMARY, 
              plot=plot, verbose=verbose, ...)
}

test.ellipseplot <- function(series=7, n=10, SUMMARY=ninenum, 
                             plot=T, verbose=F, ...) {
  x <- rnorm(n * series)
  y <- rnorm(n * series)
  f <- rep(1:series, each=n)
  ellipseplot(data.frame(f=LETTERS[f],x=x+f), 
              data.frame(f=LETTERS[f],y=y+f),
              SUMMARY=SUMMARY, 
              plot=plot, verbose=verbose, ...)
}

