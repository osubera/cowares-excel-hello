midpoint <-
function(x) {
  n <- (x[1] + x[2]) / 2
  list(c(x[1], floor(n)), c(ceiling(n), x[2]))
}
