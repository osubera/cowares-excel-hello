midpoints <-
function(x, n=1, na.rm=TRUE) {
  xna <- is.na(x)
  if(na.rm) x <- x[!xna]
  if((!na.rm && any(xna)) || (length(x) == 0)) 
    return(rep.int(NA, 2^n + 1))

  midpoints.calc(sort(x), n)
}
