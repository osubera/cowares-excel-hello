test1.normdata <- function() {
  nf <- 5
  nx <- 100
  ny <- 50
  datax <- rnorm(nx, mean=10, sd=3)
  datay <- rnorm(ny, mean=30, sd=8)
  factorx <- rep(1:nf, each=nx/nf)
  factory <- rep(1:nf, each=ny/nf)
  x <- data.frame(factorx, datax)
  y <- data.frame(factory, datay)
  invisible(list(x=x, y=y))
}

test3.lineardata <- function() {
  nf <- 5
  nx <- 100
  datax <- rnorm(nx, mean=10, sd=3)
  datay <- datax + rnorm(nx, mean=0, sd=0.1)
  factory <- factorx <- rep(1:nf, each=nx/nf)
  x <- data.frame(factorx, datax)
  y <- data.frame(factory, datay)
  invisible(list(x=x, y=y))
}

test4.lineardata <- function() {
  nf <- 4
  nx <- 100
  datax <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  datay <- datax + rnorm(nx, mean=0, sd=0.1)
  factory <- factorx <- rep(1:nf, each=nx/nf)
  x <- data.frame(factorx, datax)
  y <- data.frame(factory, datay)
  invisible(list(x=x, y=y))
}

test5.normdata <- function() {
  nf <- 4
  nx <- 100
  ny <- 40
  datax <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  datay <- rnorm(ny, mean=30, sd=8)
  factorx <- rep(1:nf, each=nx/nf)
  factory <- rep(1:nf, each=ny/nf)
  x <- data.frame(factorx, datax)
  y <- data.frame(factory, datay)
  invisible(list(x=x, y=y))
}

test1 <- function() {
  testdata <- test1.normdata()
  boxplotdou(testdata$x,testdata$y,plot=F,verbose=T)
}

test2 <- function() {
  testdata <- test1.normdata()
  boxplotdou(testdata$x,testdata$y)
}

test3 <- function() {
  testdata <- test3.lineardata()
  boxplotdou(testdata$x,testdata$y)
}

test4 <- function() {
  testdata <- test4.lineardata()
  boxplotdou(testdata$x,testdata$y)
}

test5 <- function() {
  testdata <- test5.normdata()
  boxplotdou(testdata$x,testdata$y)
}
