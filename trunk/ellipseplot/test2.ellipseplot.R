source("/opt/src/r-dev/beta/midpoints.R")
source("/opt/src/r-dev/beta/has.overlap.R")
source("/opt/src/r-dev/beta/ellipseplot.R")
source("/opt/src/r-dev/beta/boxplotdou.R")

random.box <- function(testdata=test1.data.independent(),
              ...) {
  boxplotdou(testdata$x, testdata$y, ...)
}

random.ell <- function(testdata=test1.data.independent(),
              ...) {
  ellipseplot(testdata$x, testdata$y, ...)
}

test1.data.independent <- function(nx=100, ny=50, nf=5) {
  obs.x <- rnorm(nx, mean=10, sd=3)
  obs.y <- rnorm(ny, mean=30, sd=8)
  factor.x <- letters[rep(1:nf, each=nx/nf)]
  factor.y <- letters[rep(1:nf, each=ny/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

single.box <- function(...) {
  testdata <- test14.data.single()
  boxplotdou(testdata$x, testdata$y, ...)
}

single.ell <- function(...) {
  testdata <- test14.data.single()
  ellipseplot(testdata$x, testdata$y, ...)
}

single.boxell <- function(SUMMARY=ninenum, ...) {
  testdata <- test14.data.single()
  xlim <- range(testdata$x[,2])
  ylim <- range(testdata$y[,2])
  boxplotdou(testdata$x, testdata$y, xlim=xlim, ylim=ylim, col='red', ...)
  par(new=T)
  ellipseplot(testdata$x, testdata$y, xlim=xlim, ylim=ylim, col='blue', SUMMARY=SUMMARY, ...)
}

single.boxell5 <- function(...) {
  single.boxell(SUMMARY=fivenum, ...)
}

single.boxell17 <- function(...) {
  single.boxell(SUMMARY=seventeennum, ...)
}

single.ellhist <- function(
                    testdata=test14.data.single(100, 50),
                    ...) {
  xh <- hist(testdata$x[,2], plot=F)
  yh <- hist(testdata$y[,2], plot=F)

  parkeeper <- par(no.readonly=T)
  layout(matrix(c(2,0,1,3),2,2,byrow=T),c(4,1),c(1,4))
  par(mar=c(3,3,1,1))
  ellipseplot(testdata$x, testdata$y, col='darkblue', ...)
  par(mar=c(0,3,1,1))
  barplot(xh$counts, axes=F, space=0, col='wheat')
  par(mar=c(3,0,1,1))
  barplot(yh$counts, axes=F, space=0, col='wheat', horiz=T)

  par(parkeeper)
}

test14.data.single <- function(nx=100, ny=50) {
  obs.x <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  obs.x <- c(rnorm(nx-1, mean=100, sd=10),147)
  obs.y <- c(rnorm(ny-2, mean=30, sd=5),45,48)
  factor.x <- rep("o", nx)
  factor.y <- rep("o", ny)
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

testdata.unify <- function(nx=1000, ny=1000) {
  obs.x <- rnorm(nx, mean=20, sd=3)
  obs.y <- runif(ny, min=100, max=200)
  factor.x <- rep("o", nx)
  factor.y <- rep("o", ny)
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

single.ellhist.unif <- function() single.ellhist(testdata.unify())

single.ellscat <- function(...) {
  parkeeper <- par(no.readonly=T)
  par(mfrow=c(2,2), mar=c(4,4,1,1))

  for(n in c(10,10000)) {
    testdata <- test14.data.single(n,n)
    xh <- hist(testdata$x[,2], plot=F)
    yh <- hist(testdata$y[,2], plot=F)
    plot(testdata$x[,2], testdata$y[,2], col='tomato', 
          xlab='obs.x', ylab='obs.y', ...)
    mtext(n, side=3, padj=3, cex=3)
    ellipseplot(testdata$x, testdata$y, col='forestgreen', ...)
  }

  par(parkeeper)
}

testdata.onebox <- data.frame(low=1:0,center=3:4,high=6:5,row.names=c('x','y'))

test.anellipse <- function(x=testdata.onebox, verbose=F) {
  plot(t(x['x',]), t(x['y',]),
       type='n',bty='n',xaxt='n',yaxt='n',xlab='',ylab='')
  anellipse(x, verbose, col='#FF000022', border='red')
  rect(x$low[1],x$low[2],x$high[1],x$high[2],border='blue',lty='dotdash')
  lines(x[1,c('low','high')],rep(x$center[2],2),col='black',lty='longdash')
  lines(rep(x$center[1],2),x[2,c('low','high')],col='black',lty='longdash')
  rect(x$center[1],x$low[2],x$high[1],x$center[2],lty='blank',col='#22c88c22')
  text(x$center[1],x$center[2],'Median',col='darkseagreen',cex=3)
  text(x$center[1],x$high[2],'6th Octile',col='darkseagreen',cex=3)
  text(x$center[1],x$low[2],'2nd Octile',col='darkseagreen',cex=3)
  text(x$center[1],(x$center[2]+x$low[2])/2,'Median',srt='90',col='darkorchid',cex=3)
  text(x$high[1],(x$center[2]+x$low[2])/2,'6th Octile',srt='90',col='darkorchid',cex=3)
  text(x$low[1],(x$center[2]+x$low[2])/2,'2nd Octile',srt='90',col='darkorchid',cex=3)
}

test15.data.independent <- function(n=10) {
  obs.x <- c(rnorm(n, mean=120, sd=18), rnorm(n, mean=100, sd=12), rnorm(n, mean=100, sd=8))
  obs.y <- c(rnorm(n, mean=30, sd=8), rnorm(n, mean=40, sd=5), rnorm(n, mean=32, sd=10))
  factor.y <- factor.x <- rep(c("a","b","c"), each=n)
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

test4.data.linear <- function(nf=4,nx=100) {
  obs.x <- c(rnorm(nx/2, mean=10, sd=3), rnorm(nx/2, mean=20, sd=2))
  obs.y <- obs.x + rnorm(nx, mean=0, sd=0.1)
  factor.y <- factor.x <- letters[rep(1:nf, each=nx/nf)]
  x <- data.frame(factor.x, obs.x)
  y <- data.frame(factor.y, obs.y)
  invisible(list(x=x, y=y))
}

linear.boxell <- function(nf=4, nx=100, ...) {
  parkeeper <- par(no.readonly=T)
  par(mfrow=c(2,2), mar=c(4,4,1,1))

  testdata <- test4.data.linear(nf,nx)

  boxplotdou(testdata$x, testdata$y, ...)
  ellipseplot(testdata$x, testdata$y, SUMMARY=fivenum, SHEER=my.sheer.color, ...)
  mtext('five number', side=3, padj=2)
  ellipseplot(testdata$x, testdata$y, SHEER=my.sheer.color, ...)
  mtext('nine number', side=3, padj=2)
  ellipseplot(testdata$x, testdata$y, SUMMARY=seventeennum, SHEER=my.sheer.color, ...)
  mtext('seventeen number', side=3, padj=2)

  par(parkeeper)
}

my.sheer.color <- function(col, level) {
  sheer <- level^2 * 0.5 + 0.2
  adjustcolor(col, alpha.f=sheer)
}



Fig.1 <- function() random.ell(testdata=test15.data.independent())
Fig.2 <- function() single.boxell5()
Fig.3 <- function() test.anellipse()
Fig.4 <- function() single.ellscat()
Fig.5 <- function() linear.boxell()
Fig.6 <- function() single.ellhist()
Fig.7 <- function() single.ellhist.unif()


pnger <- function() {
  lapply(as.list(1:7), function(i) {
    png(paste('fig',i,'.png',sep=''))
    eval(parse(text=paste('Fig.',i,'()',sep='')))
    dev.off()
  })
}

