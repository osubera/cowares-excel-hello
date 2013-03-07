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

test7 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-5,1,-3,4)) == T
}

test8 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(15,1,23,4)) == T
}

test9 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,-6,3,-4)) == T
}

test10 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,11,3,14)) == T
}

test11 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(1,1,3,4)) == T
}

test12 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(2,2,4,5)) == T
}

test13 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(12,12,14,15)) == F
}

test14 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(12,-8,14,-5)) == F
}

test15 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-7,12,-4,15)) == F
}

test16 <- function() {
  has.overlap.xory.rect(c(1,1,3,4),c(-7,-8,-4,-5)) == F
}

