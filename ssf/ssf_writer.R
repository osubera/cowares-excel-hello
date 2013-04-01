# ssfwriter for R
# convert list of data.frame into multiple excel worksheet
# http://code.google.com/p/cowares-excel-hello/wiki/ssf_writer_primitive_r
# http://tomizonor.wordpress.com/
#
# Copyright (C) 2013 Tomizono - kobobau.mocvba.com
# Fortitudinous, Free, Fair, http://cowares.nobody.jp
#
writessf <- function(x, ...) UseMethod('writessf')

writessf.default <- function(x, file=stdout(), ...) {
  writessf.data.frame(data.frame(x), file=file)
}

writessf.list <- function(x, file=stdout(), ...) {
  for(i in 1L:length(x)) {
    key <- names(x[i])
    value <- x[[i]]
    ssfblock.worksheet(key, file=file)
    writessf(value, file=file)
  }
}

writessf.data.frame <- function(x, file=stdout(), ...) {
  ssfblock.cellsformula.data.frame(x, file=file)
}

# map 1:26 into A:Z
AtoZ <- c(chartr('[0-9]', '[A-J]', 0:9), 
          chartr('[0-9]', '[K-T]', 0:9), 
          chartr('[0-5]','[U-Z]',0:5))

num.to.col <- function(number) {
  col <- c()
  while(number > 0) {
    x <- ((number - 1) %% 26) + 1
    col <- c(AtoZ[x], col)
    number <- (number - x) / 26
  }
  paste(col, collapse='')
}

ssfline.address.data.frame <- function(x, file=stdout(),
  bch="'", ech='\n', delimiter=';', 
  ...) {
  address <- paste('A1:', num.to.col(ncol(x)), nrow(x)+1,
                   sep='')
  cat(bch, 'address', delimiter, address, ech,
      sep='', file=file)
  invisible(address)
}

ssfline.formula.data.frame <- function(x, file=stdout(),
  bch="'", ech='\n', delimiter=';', 
  ...) {
  collapse.chars <- paste(ech, bch, delimiter, sep='')
  cat(bch, delimiter,
      paste(colnames(x), collapse=collapse.chars),
      ech,
      sep='', file=file)
  apply(x, 1, function(a) {
        ssfline.vector(a, file=file,
          bch=bch, ech=ech, delimiter=delimiter,
          ...)
#        cat(bch, delimiter,
#            paste(a, collapse=collapse.chars),
#            ech,
#            sep='', file=file)
      })
}

ssfline.vector <- function(x, file=stdout(),
  bch="'", ech='\n', delimiter=';', 
  ...) {
  apply(as.matrix(x), 1, function(a) {
    ssfline.character(a, file=file,
      bch=bch, ech=ech, delimiter=delimiter,
      ...)
  })
}

ssfline.character <- function(x, file=stdout(), 
  bch="'", ech='\n', delimiter=';', 
  ...) {
  if(length(grep(ech, x)) > 0) {
    ssfline.escape(x, file=file,
      bch=bch, ech=ech, delimiter=delimiter,
      ...)
  } else {
    cat(bch, delimiter, x, ech,
        sep='', file=file)
  }
}

ssfline.escape <- function(x, file=stdout(),
  escapebegin='{{{', escapeend='}}}',
  bch="'", ech='\n', delimiter=';', 
  ...) {
  cat(bch, escapebegin, ech,
      x, ech,
      bch, escapeend, ech,
      sep='', file=file)
}

ssfblock.ssfbegin <- function(file=stdout(), 
  bch="'", ech='\n', delimiter=';', 
  magicbegin='ssf-begin',
  ...) {
  cat(bch, magicbegin, ech,
      bch, delimiter, ech,
      ech,
      sep='', file=file)
}

ssfblock.ssfend <- function(file=stdout(), 
  bch="'", ech='\n', delimiter=';', 
  magicend='ssf-end',
  ...) {
  cat(bch, magicend, ech,
      ech,
      sep='', file=file)
}

ssfblock.worksheet <- function(name, file=stdout(),
  bch="'", ech='\n', delimiter=';', 
  blockname='worksheet',
  ...) {
  cat(bch, blockname, ech,
      bch, 'name', delimiter, name, ech,
      ech,
      sep='', file=file)
}

ssfblock.cellsformula.data.frame <- function(x, file=stdout(),
  bch="'", ech='\n', delimiter=';', 
  blockname='cells-formula',
  ...) {
  cat(bch, blockname, ech,
      sep='', file=file)
  ssfline.address.data.frame(x, file=file)
  ssfline.formula.data.frame(x, file=file)
  cat(ech,
      sep='', file=file)
}


