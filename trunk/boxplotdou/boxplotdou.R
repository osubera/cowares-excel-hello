# double boxplot

# boxplotdou(cbind(factor,data1),cbind(factor,data2))


boxplotdou <- function(x,y,col) {

levels.xy <- unique(c(unique(as.character(x[,1])),unique(as.character(y[,1]))))
factor.x <- factor(x[,1],levels=levels.xy)
factor.y <- factor(y[,1],levels=levels.xy)
levels.n <- length(levels.xy)
#print(levels.xy)
#str(factor.x)
print(levels.n)

data.x <- x[,2]
data.y <- y[,2]
name.x <- names(x[2])
name.y <- names(y[2])
name.x <- ifelse(is.null(name.x),"x",name.x)
name.y <- ifelse(is.null(name.y),"y",name.y)
str(data.x)
print(name.x)
str(data.y)
print(name.y)

stats.x <- plot(factor.x,data.x,plot=F)
stats.y <- plot(factor.y,data.y,plot=F)

#factors <- as.factor(unique(c(stats.x$names,stats.y$names)))
print(stats.x$names)
print(stats.y$names)
print(stats.x)
print(stats.y)

min.x <- min(stats.x$stats,na.rm=T)
max.x <- max(stats.x$stats,na.rm=T)
min.y <- min(stats.y$stats,na.rm=T)
max.y <- max(stats.y$stats,na.rm=T)

levels.col <- rainbow(levels.n)
#substr(levels.col,8,9) <- "66"
#print(levels.col)
#str(levels.col)

plot(NULL,xlim=c(min.x,max.x),ylim=c(min.y,max.y),xlab=name.x,ylab=name.y)

for(i in 1:levels.n) mysinglebox(stats.x,stats.y,i,as.character(levels.xy)[i],levels.col[i])

}


mysinglebox <- function(x,y,column.num,column.char,color) {

xlow  <- x$stats[2,column.num]
xhigh <- x$stats[4,column.num]
ylow  <- y$stats[2,column.num]
yhigh <- y$stats[4,column.num]

print(column.num)
print(column.char)
print(color)
print(xlow)
print(xhigh)
print(ylow)
print(yhigh)

color.sheer <- paste(substring(color,1,7),"33",sep="")
print(color.sheer)

rect(xlow,ylow,xhigh,yhigh,col=color.sheer)

xlowend  <- x$stats[1,column.num]
xhighend <- x$stats[5,column.num]
ylowend  <- y$stats[1,column.num]
yhighend <- y$stats[5,column.num]

xcenter  <- x$stats[3,column.num]
ycenter  <- y$stats[3,column.num]

print(xlowend)
print(xcenter)
print(xhighend)
print(ylowend)
print(ycenter)
print(yhighend)

if(F) {
xbarlow  <- xlow
xbarhigh <- xhigh
ybarlow  <- ylow
ybarhigh <- yhigh
} else {
xbarlow  <- xlowend
xbarhigh <- xhighend
ybarlow  <- ylowend
ybarhigh <- yhighend
}

segments(xlowend,ybarlow,xlowend,ybarhigh,col=color)
segments(xhighend,ybarlow,xhighend,ybarhigh,col=color)
segments(xbarlow,ylowend,xbarhigh,ylowend,col=color)
segments(xbarlow,yhighend,xbarhigh,yhighend,col=color)

segments(xlowend,ycenter,xhighend,ycenter,col=color)
segments(xcenter,ylowend,xcenter,yhighend,col=color)

out.x <- x$out
out.x.num <- length(out.x)
out.x.group <- x$group
out.y <- y$out
out.y.num <- length(out.y)
out.y.group <- y$group

print(out.x)
print(out.x.num)
print(out.x.group)
print(out.y)
print(out.y.num)
print(out.y.group)

for(x in out.x[out.x.group==column.num]) points(x,ycenter,col=color,pch=1,cex=2)
for(y in out.y[out.y.group==column.num]) points(xcenter,y,col=color,pch=1,cex=2)

if(T) {
for(x in out.x[out.x.group==column.num]) {
segments(x,ycenter,xcenter,ycenter,col=color.sheer)
segments(x,ylow,x,yhigh,col=color.sheer)
}
for(y in out.y[out.y.group==column.num]) {
segments(xcenter,y,xcenter,ycenter,col=color.sheer)
segments(xlow,y,xhigh,y,col=color.sheer)
}
}

text(xcenter,ycenter,column.char)

if(!is.na(xcenter)) mtext(column.char,side=3,at=xcenter,col=color)
if(!is.na(ycenter)) mtext(column.char,side=4,at=ycenter,col=color)

}

