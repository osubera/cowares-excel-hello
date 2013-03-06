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

