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
