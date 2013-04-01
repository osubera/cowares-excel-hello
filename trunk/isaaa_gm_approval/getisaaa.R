# download isaaa gm approval database into R
# http://www.isaaa.org/gmapprovaldatabase/
# http://code.google.com/p/cowares-excel-hello/
# http://tomizonor.wordpress.com/
#
# Copyright (C) 2013 Tomizono - kobobau.mocvba.com
# Fortitudinous, Free, Fair, http://cowares.nobody.jp
#

# level2urls
level2urls <- c(
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=AR&Country=Argentina',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=AU&Country=Australia',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=BO&Country=Bolivia',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=BR&Country=Brazil',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=BF&Country=Burkina+Faso',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CA&Country=Canada',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CL&Country=Chile',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CN&Country=China',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CO&Country=Colombia',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CR&Country=Costa+Rica',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=EG&Country=Egypt',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=SV&Country=El+Salvador',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=EU&Country=European+Union',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=HN&Country=Honduras',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=IN&Country=India',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=ID&Country=Indonesia',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=IR&Country=Iran',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=JP&Country=Japan',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=MY&Country=Malaysia',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=MX&Country=Mexico',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=MM&Country=Myanmar',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=NZ&Country=New+Zealand',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=PK&Country=Pakistan',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=PY&Country=Paraguay',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=PH&Country=Philippines',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=RU&Country=Russian+Federation',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=ZA&Country=South+Africa',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=KR&Country=South+Korea',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=CH&Country=Switzerland',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=TW&Country=Taiwan',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=TR&Country=Turkey',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=US&Country=United+States+of+America',
'http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=UY&Country=Uruguay'
)

# level2urls.use
level2urls.use <- level2urls[c(32,4,1)]

# level2urls.crop
level2urls.crop <- c(
'http://www.isaaa.org/gmapprovaldatabase/crop/default.asp?CropID=2&Crop=Argentine+Canola',
'http://www.isaaa.org/gmapprovaldatabase/crop/default.asp?CropID=14&Crop=Polish+canola',
'http://www.isaaa.org/gmapprovaldatabase/crop/default.asp?CropID=7&Crop=Cotton',
'http://www.isaaa.org/gmapprovaldatabase/crop/default.asp?CropID=6&Crop=Maize',
'http://www.isaaa.org/gmapprovaldatabase/crop/default.asp?CropID=19&Crop=Soybean'
)

# preload()
preload <- function() {
  library(XML)
  library(RCurl)
}

# trim(x)
trim <- function(x) gsub(pattern='\\s*$|^\\s*', replacement='', x=x)

# nodeText(node)
nodeText <- function(x) trim(toString.XMLNode(x[['text']]))

# makedatanew()
makedatanew <- function() data.frame(id=NA, stringsAsFactors=F)

# makelog()
makelog <- function() data.frame(log=c(''),stringsAsFactors=F)

# addlog()
addlog <- function(log, x) {
  n <- nrow(log)
  log[(n+1):(n+length(x)),] <- x
  log
}

# xmlAttrs(node, ...)
xmlAttrs.NULL <- function(node, ...) NULL

# simplify2data.frame(x)
simplify2data.frame <- function(x) {
  data.frame(key=names(x), value=simplify2array(x), stringsAsFactors=F)
}

# getlevel3(url)
# http://www.isaaa.org/gmapprovaldatabase/event/default.asp?EventID=173
#
getlevel3 <- function(url) {
  id <- strsplit(tolower(url), 'eventid=')[[1]][2]
  time <- as.character(Sys.time())
  src <- getURL(url)
  doc <- htmlParse(src)

  basicinfo <- getNodeSet(doc, 
    'id("TabbedPanels1")//div[@class="TabbedPanelsContent"][1]//p')[1:4]
  basicinfo.parsed <- lapply(basicinfo, function(x) {
      sapply(x['a'], function(a) nodeText(a))
    })
  developer <- basicinfo.parsed[[1]]
  method <- basicinfo.parsed[[2]]
  gm <- basicinfo.parsed[[3]]
  commercial <- basicinfo.parsed[[4]]
  rm(basicinfo)
  
  genetictable <- getNodeSet(doc, '//table')
  genetic <- readHTMLTable(genetictable[[1]], stringsAsFactors=F, header=T)
  rm(genetictable)
  
  free(doc)

  list(id=id, time=time, url=url, src=src,
    developer=developer, method=method, gm=gm, 
    commercial=commercial, genetic=genetic)
}

# getlevel2(url, id)
# http://www.isaaa.org/gmapprovaldatabase/approvedeventsin/default.asp?CountryID=EG&Country=Egypt
#
getlevel2 <- function(url, id=NULL) {
  singlecrop <- F
  if(is.null(id)) {
    id <- gsub('\\+',' ',strsplit(tolower(url), 'country=')[[1]][2])
    if(is.na(id)) {
      id <- gsub('\\+',' ',strsplit(tolower(url), 'crop=')[[1]][2])
      singlecrop <- T
    }
  }
  time <- as.character(Sys.time())
  src <- getURL(url)
  doc <- htmlParse(src)

  baseurl <- paste(parseURI(url)[c('scheme','server')], collapse='://')

  event.text <- readHTMLTable(doc, which=1, stringsAsFactors=F, header=T)
  event.url <- readHTMLTable(doc, which=1, stringsAsFactors=F, header=T,
    elFun=function(x) {
      # getHTMLLinks(x, baseURL=url, relative=T)[1]
      # paste(baseurl, xmlAttrs(x[['strong']][['a']])['href'], sep='')
      sub('(eventid=\\d+)&.*', '\\1',
        paste(baseurl, xmlAttrs(x[['strong']][['a']])['href'], sep=''),
        ignore.case=T)
    })

  n <- nrow(event.text)
  names(event.text) <- c('Event', 'Trade')
  names(event.url) <- c('Url', 'Crop')
  ev <- data.frame(
    apply(data.frame(event.text, event.url, EventID=rep('',n)), 
      2, as.character),
    stringsAsFactors=F)

  if(singlecrop) {
    ev[,'Crop'] <- id
  } else {
    for(i in 1L:n) 
      ev[i,'Crop'] <- 
        ifelse(is.na(ev[i, 'Crop']), ev[i, 'Event'], ev[i-1, 'Crop'])
  }

  event <- ev[!is.na(ev[,'Trade']),]

  event[,'EventID'] <- 
    sapply(strsplit(tolower(event[,'Url']), 'eventid='),
      function(x) x[2])

  free(doc)

  list(id=id, time=time, url=url, src=src,
    event=event)
}

# scanlevel2(urls, file)
#
scanlevel2 <- function(urls=level2urls.use, file=NULL) {
  bag <- list()
  log <- makelog()
  log <- addlog(log, 'scanlevel2 begins')
  level3urls <- makedatanew()
  level2urls <- list()

  for(url in urls) {
    log <- addlog(log, url)
    level2 <- getlevel2(url)
    log <- addlog(log, level2$id)
    
    if(!is.null(file)) {
      filename <- paste(file, level2$id, '.html', sep='')
      cat(level2$src, file=filename)
      log <- addlog(log, paste('saved', filename))
    }

    name.basic <- paste(level2$id, 'a', sep='.')
    name.event <- paste(level2$id, 'b', sep='.')
    list.basic <- level2[c('id', 'time', 'url')]

    bag[[name.basic]] <- simplify2data.frame(list.basic)
    bag[[name.event]] <- level2$event

    level3urls <- addlog(level3urls, level2$event[,'Url'])
    list.basic$basicinfo <- name.basic
    list.basic$eventinfo <- name.event
    level2urls[[level2$id]] <- list.basic
  }

  bag$level3urls <- 
    data.frame(Url=unique(level3urls[-1,]), stringsAsFactors=F)
  bag$level2urls <- 
    data.frame(t(sapply(level2urls, simplify2array)), stringsAsFactors=F)
  bag$log <- log

  invisible(bag)
}

# scanlevel3(bag, file)
#
scanlevel3 <- function(bag, file=NULL) {
  log <- bag$log
  urls <- bag$level3urls[,1]
  developer <- makedatanew()
  method <- makedatanew()
  gm <- makedatanew()
  commercial <- makedatanew()

  for(url in urls) {
    log <- addlog(log, url)
    level3 <- getlevel3(url)
    log <- addlog(log, level3$id)

    if(!is.null(file)) {
      filename <- paste(file, level3$id, '.html', sep='')
      cat(level3$src, file=filename)
      log <- addlog(log, paste('saved', filename))
    }

    basicinfo <- data.frame(simplify2data.frame(level3[c('id', 'time', 'url')]),
                            developer=paste(level3$developer, collapse=', '),
                            method=paste(level3$method, collapse=', '),
                            gm=paste(level3$gm, collapse=', '),
                            commercial=paste(level3$commercial, collapse=', '),
                            stringsAsFactors=F)
    bag[[paste('E', level3$id, 'a', sep='.')]] <- basicinfo
    bag[[paste('E', level3$id, 'b', sep='.')]] <- level3$genetic

    developer[level3$id, 'id'] <- level3$id
    for(x in level3$developer) developer[level3$id, x] <- x
    method[level3$id, 'id'] <- level3$id
    for(x in level3$method) method[level3$id, x] <- x
    gm[level3$id, 'id'] <- level3$id
    for(x in level3$gm) gm[level3$id, x] <- x
    commercial[level3$id, 'id'] <- level3$id
    for(x in level3$commercial) commercial[level3$id, x] <- x
  }

  bag$developer <- developer[!is.na(developer$id),]
  bag$method <- method[!is.na(method$id),]
  bag$gm <- gm[!is.na(gm$id),]
  bag$commercial <- commercial[!is.na(commercial$id),]

  bag$log <- log

  invisible(bag)
}

# mergelevel3(bag)
#
mergelevel3 <- function(bag) {
  for(eventlist.name in bag$level2urls$eventinfo) {
    eventlist <- bag[[eventlist.name]]
    developer <- bag$developer[eventlist$EventID,]
    names(developer) <- paste('developer', names(developer), sep='_')
    method <- bag$method[eventlist$EventID,]
    names(method) <- paste('method', names(method), sep='_')
    gm <- bag$gm[eventlist$EventID,]
    names(gm) <- paste('gm', names(gm), sep='_')
    commercial <- bag$commercial[eventlist$EventID,]
    names(commercial) <- paste('commercial', names(commercial), sep='_')
    bag[[eventlist.name]] <- cbind(eventlist, developer, method, gm, commercial)
  }

  invisible(bag)
}

