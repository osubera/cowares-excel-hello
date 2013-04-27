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

