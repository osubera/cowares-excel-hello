# test midpoints

# test five number summary of quartiles
test1 <- function(r=1, n=100) {
  myfivenum <- function(x, ...) midpoints(x, 2, ...)
  res <- sapply(as.list(1:r),
                function(a) {
                  data <- rnorm(n)
                  control <- fivenum(data)
                  my5 <- myfivenum(data)
                  my9 <- ninenum(data)
                  my17 <- seventeennum(data)
                  all(my5 == control) &&
                  all(my9[seq(1,9,2)] == control) &&
                  all(my17[seq(1,17,4)] == control)
                })
  all(res)
}

# test nine number summary of octiles
#  and seventeen number summary of hexadeciles
test2 <- function(r=1, n=100) {
  res <- sapply(as.list(1:r),
                function(a) {
                  data <- rnorm(n)
                  control9 <- testninenum(data)
                  control17 <- testseventeennum(data)
                  my9 <- ninenum(data)
                  my17 <- seventeennum(data)
                  all(my9 == control9) &&
                  all(my17 == control17) 
                })
  all(res)
}

# check differences to zarfivenum
test3 <- function(r=1, n=9) {
  myfivenum <- function(x, ...) midpoints(x, 2, ...)
  res <- sapply(as.list(1:r),
                function(a) {
                  data <- rnorm(n)
                  control <- zarfivenum(data)
                  my5 <- myfivenum(data)
                  my5 - control
                })
  rowMeans(res)
}

# check differences to zarninenum
test4 <- function(r=1, n=9) {
  res <- sapply(as.list(1:r),
                function(a) {
                  data <- rnorm(n)
                  control <- zarninenum(data)
                  my9 <- ninenum(data)
                  my9 - control
                })
  rowMeans(res)
}

# bulk test
test5 <- function() {
  print(rowSums(sapply(as.list(1:100), function(a) testninenum(1:a) - ninenum(1:a))))
  print(rowSums(sapply(as.list(1:100), function(a) testseventeennum(1:a) - seventeennum(1:a))))
  print(all(sapply(as.list(1:100), function(n) test1(30,n))))
  print(all(sapply(as.list(1:100), function(n) test2(30,n))))

  print(min(sapply(as.list(1:100), function(a) zarninenum(1:a) - ninenum(1:a))))
  print(max(sapply(as.list(1:100), function(a) zarninenum(1:a) - ninenum(1:a))))
  print(rowMeans(sapply(as.list(1:100), function(a) zarninenum(1:a) - ninenum(1:a))))
  print(rowMeans(sapply(as.list(1:100), function(n) test3(30,n))))
  print(rowMeans(sapply(as.list(1:100), function(n) test4(30,n))))
}

# visualize differences between ninenum and zarninenum
test6 <- function(n1=1, n2=25) {
  d <- data.frame(t(
         sapply(as.list(n1:n2), function(a) zarninenum(1:a) - ninenum(1:a))
       ))
  names(d) <- paste(0:8, c('','st','nd','rd',rep('th',5)), ' O', sep='')
  names(d)[1] <- 'Min'
  names(d)[5] <- 'Med'
  names(d)[9] <- 'Max'
  parkeeper <- par(c('mfrow','mar'))
  par(mfrow=c(7,1), mar=c(2,4,1,1))
  lapply(as.list(c(2:4,6:8)), function(i)
    plot(d[,i],type='b',ylab='diff',xlab='',main=names(d)[i],
    ylim=c(-0.5,0.5),yaxp=c(-0.5,0.5,2))
  )
  matplot(d[,c(1,5,9)],type='l',ylab='',xlab='length',
    main='Min, Med, Max',mgp=c(1,1,0),
    ylim=c(-0.5,0.5),yaxp=c(-0.5,0.5,2))
  par(parkeeper)
  d
}

test7 <- function(n1=1, n2=100) {
  d <- sweep(t(
         sapply(as.list(n1:n2), function(a) zarninenum(1:a) - ninenum(1:a))
       ), 1, (n1:n2)-1, '/')
  matplot(d,type='l',ylab='diff',xlab='length')
}


### other calculations

# direct calculation

testninenum <- function(x, na.rm=TRUE)
{
  xna <- is.na(x)
  if(na.rm) x <- x[!xna]
  else if(any(xna)) return(rep.int(NA,9))
  x <- sort(x)
  n <- length(x)
  if(n == 0) {
    rep.int(NA,9)
  } else {
    n2 <- (n+1) / 2
    n4 <- floor((n+3) / 2) / 2
    n8 <- floor((n+7) / 4) / 2
    d <- c(1, n8, n4, n4 * 2 - n8, 
           n2, 
           n + 1 - n4 * 2 + n8, n + 1 - n4, n + 1 - n8, n)
    0.5 * (x[floor(d)] + x[ceiling(d)])
  }
}

testseventeennum <- function(x, na.rm=TRUE)
{
  xna <- is.na(x)
  if(na.rm) x <- x[!xna]
  else if(any(xna)) return(rep.int(NA,17))
  x <- sort(x)
  n <- length(x)
  if(n == 0) {
    rep.int(NA,17)
  } else {
    n2 <- (n+1) / 2
    n4 <- floor((n+3) / 2) / 2
    n8 <- floor((n+7) / 4) / 2
    n16 <- floor((n+15) / 8) / 2
    d <- c(1, n16, n8, n8 * 2 - n16,
           n4,
           n4 * 2 - n8 * 2 + n16, 
           n4 * 2 - n8, n4 * 2 - n16, 
           n2, 
           n + 1 - n4 * 2 + n16, n + 1 - n4 * 2 + n8, 
           n + 1 - n4 * 2 + n8 * 2 - n16,
           n + 1 - n4,
           n + 1 - n8 * 2 + n16, n + 1 - n8, n + 1 - n16, n)
    0.5 * (x[floor(d)] + x[ceiling(d)])
  }
}


# according to Zar

zarfivenum <- function(x) {
  x <- sort(x)
  n <- length(x)
  n2 <- (n + 1) / 2
  n4 <- ceiling((n + 1) / 2) / 2
  d <- pmax(1, pmin(n, c(1, n4, n2, n + 1 - n4, n)))
  0.5 * (x[floor(d)] + x[ceiling(d)])
}

zarninenum <- function(x) {
  x <- sort(x)
  n <- length(x)
  n2 <- (n + 1) / 2
  n4 <- ceiling((n + 1) / 2) / 2
  n8 <- ceiling((n + 1) / 4) / 2
  d <- pmax(1, pmin(n, 
         c(1, n8, n4, n4 * 2 - n8, n2, 
           n + 1 - n4 * 2 + n8, n + 1 - n4, n + 1 - n8, n)
       ))
  0.5 * (x[floor(d)] + x[ceiling(d)])
}

