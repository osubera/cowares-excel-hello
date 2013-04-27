# test midpoints

test1 <- function(n=1) {
  myfivenum <- function(x, ...) midpoints(x, 2, ...)
  res <- sapply(as.list(1:n),
                function(a) {
                  data <- rnorm(100)
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
test2 <- function(n=1) {
  res <- sapply(as.list(1:n),
                function(a) {
                  data <- rnorm(100)
                  control9 <- testninenum(data)
                  control17 <- testseventeennum(data)
                  my9 <- ninenum(data)
                  my17 <- seventeennum(data)
                  all(my9 == control9) &&
                  all(my17 == control17) 
                  c((my9 == control9), (my17 == control17))
                  list(rbind(my9,control9),rbind(my17,control17))
                })
  #all(res)
  res
}

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
    n4 <- floor(n2+1) / 2
    n8 <- floor(n4+1) / 2
    d <- c(1, n8, n4, n2 + 1 - n8, 
           n2, 
           n2 - 1 + n8, n + 1 - n4, n + 1 - n8, n)
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
    n4 <- floor(n2+1) / 2
    n8 <- floor(n4+1) / 2
    n16 <- floor(n8+1) / 2
    d <- c(1, n16, n8, n4 + 1 - n16, 
           n4, 
           n4 - 1 + n16, n4 -1 + n8, n2 + 1 - n16, 
           n2, 
           n2 - 1 + n16, n2 - 1 + n8, n2 - 2 + n8 + n16, 
           n2 - 1 + n4,
           n + 2 - n8 - n16, n + 1 - n8, n + 1 - n16, n)
    0.5 * (x[floor(d)] + x[ceiling(d)])
  }
}


