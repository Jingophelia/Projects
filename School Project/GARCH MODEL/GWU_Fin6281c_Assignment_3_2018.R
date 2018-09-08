# Assignment 3 Code
# You will use this code to calculate summary statistics,
# create plots, and ultimately fit a GARCH(1,1) model and forecast out of sample
# one-month realized volatility to compare to actual realized volatility
# for the S&P 500 Index

# Install packages
install.packages("quantmod")
install.packages("tseries")
install.packages("xts")
install.packages("stats")
install.packages("rugarch")
install.packages("e1071")
install.packages("MASS")
install.packages("zoo")
install.packages("roll")
install.packages("car")

# Require library
library(quantmod)
library(tseries)
library(xts)
library(stats)
library(rugarch)
library(e1071)
library(MASS)
library(zoo)
library(roll)
library(car)

# This is where you read the data in from a CSV file
# You will need to change the path to point towards where you save your data
path <- "C:/Users/Stephen D Young/Documents/Stephen D. Young/Teaching/GWU 2018/Assignments FINA 6281 Summer 2018/Data.csv"
data.file <- read.csv(path,header=TRUE)

# Data structure - tells us what is in the file 
str(data.file)

# View header and footer of data records
head(data.file)
tail(data.file)

# Create data frame object
# MY DATA HAS DATE, HIGH, LOW, OPEN, CLOSE, AND VIX INDEX LEVEL
# YOU ONLY NEED DATE AND CLOSING PRICES BUT DO NOT MODIFY THE CSV FILE
SPX.df.date <- as.Date(data.file$SPX_DATE, format="%m/%d/%Y")
SPX.data.frame <- cbind(SPX.df.date, data.file[,5:6])
SPX.data.frame <- na.omit(SPX.data.frame[,1:3])
head(SPX.data.frame)
tail(SPX.data.frame)

# Create time series object
SPX.time.series <- xts(SPX.data.frame[,2:3], order.by=SPX.data.frame[,1])
head(SPX.time.series)
tail(SPX.time.series)

# Calculate log returns and get rid of first observation as you lose one date
SPX.log.returns <- diff(log(SPX.time.series$SPX_CLOSE))
SPX.log.returns <- SPX.log.returns[-1,]

# Plot time series of log returns to show volatility clustering
# Clustering is the behavior where high volatility is followed by high volatility
# and low volatility is followed by low volatility
class(SPX.log.returns)
SPX.log.returns.df <-cbind(index(SPX.log.returns),data.frame(SPX.log.returns))
names(SPX.log.returns.df) <- paste(c("date","spx"))
head(SPX.log.returns.df)
class(SPX.log.returns.df)

# Calculate mean and standard deviation to use in plot
meanlevel <- mean(SPX.log.returns.df$spx)
sdlevel <- sd(SPX.log.returns.df$spx)

plot(SPX.log.returns.df$date, SPX.log.returns.df$spx, axes=TRUE, ylab = "Log Returns",
     xlab="",ylim=c(-.15,.15), col = ifelse(abs(SPX.log.returns.df$spx) > 2*sdlevel,'red','black'),cex.axis = .7)
title(main="SPX Log Returns", cex=.7)
mtext("Date",side=1,col="black",line=2)
text(x=as.Date('2010-10-15'), y=.135, labels='Volatility Clustering', col="black", cex=.75)
# length slightly shrinks the size of the arrow head; lwd makes the line bolder
arrows(x0=as.Date('2010-10-15'), y0=.13, x1=as.Date('2009-05-01'), y1=.07, col='black', length=0.1, lwd=1)
arrows(x0=as.Date('2010-10-15'), y0=.13, x1=as.Date('2011-07-01'), y1=.07, col='black', length=0.1, lwd=1)
text(x=as.Date('2010-10-15'), y=-.135, labels='Volatility Clustering', col="black", cex=.75)
# length slightly shrinks the size of the arrow head; lwd makes the line bolder
arrows(x0=as.Date('2010-10-15'), y0=-.13, x1=as.Date('2009-05-01'), y1=-.07, col='black', length=0.1, lwd=1)
arrows(x0=as.Date('2010-10-15'), y0=-.13, x1=as.Date('2011-07-01'), y1=-.07, col='black', length=0.1, lwd=1)
abline(h=meanlevel + 2 * sdlevel, col = "red", lty=2, lwd=2)
abline(h=meanlevel - 2 * sdlevel, col = "red", lty=2, lwd=2)
abline(h=meanlevel + 3 * sdlevel, col = "blue", lty=2, lwd=2)
abline(h=meanlevel - 3 * sdlevel, col = "blue", lty=2, lwd=2)
legend("topleft",legend=c("+/-2 Standard Deviations","+/- 3 Standard Deviations"),bty="n",lty=c(2,2), cex=0.80, col=c("red","blue")) 

# Summary statistics for equity index log returns
SPX.obs <- length(SPX.log.returns)
SPX.min.return <- min(SPX.log.returns)
SPX.mean.return <- mean(SPX.log.returns)
SPX.median.return <- median(SPX.log.returns)
SPX.max.return <- max(SPX.log.returns)
SPX.skew.return <- skewness(SPX.log.returns)
SPX.kurt.return <- kurtosis(SPX.log.returns)
SPX.ann.sd <- sd(SPX.log.returns)*sqrt(252)
SPX.JB.stat <- jarque.bera.test(SPX.log.returns)$statistic
SPX.JB.pvalue <- jarque.bera.test(SPX.log.returns)$p.value

# Matrix of outputs
# You create a matrix by using matrix and specifying how many rows and columns
Summary.stats <- as.data.frame(matrix(data=NA,nrow = 10, ncol = 1))

# We will put the summary statistics into the matrix
Summary.stats[1,1] <- SPX.obs
Summary.stats[2,1] <- SPX.min.return
Summary.stats[3,1] <- SPX.mean.return
Summary.stats[4,1] <- SPX.median.return
Summary.stats[5,1] <- SPX.max.return
Summary.stats[6,1] <- SPX.skew.return
Summary.stats[7,1] <- SPX.kurt.return
Summary.stats[8,1] <- SPX.ann.sd
Summary.stats[9,1] <- SPX.JB.stat
Summary.stats[10,1] <- SPX.JB.pvalue

rownames(Summary.stats) <- c("No. Obs.", "Min", "Mean", "Median", "Max", "Skewness",
                                  "Kurtosis", "Ann. Std.", "Jarque-Bera Statistic", "Jarque-Bera p-Value")
colnames(Summary.stats) <- c("SPX")
Summary.stats

# Start and end dates for time series
SPX.start.date <- as.Date(SPX.data.frame[1,1],format="%m/%d/%Y")
SPX.end.date <- as.Date(SPX.data.frame[SPX.obs+1,1],format="%m/%d/%Y")
Summary.dates <- data.frame(c(SPX.start.date, SPX.end.date))
Summary.dates

# Histogram of log returns to index and Normal and Student-t distributions superimposed
hist_returns <- SPX.log.returns
hist_mean <- mean(hist_returns)
hist_std <- sqrt(var(hist_returns))

hist.SPX <- hist(hist_returns, breaks=100)
plot(hist.SPX, xlim=c(-.05,.05), ylim=c(0,60), freq=F, xlab="Log Return",col="white",
     main="Histogram of SPX Log Returns, Normal and Student-t Densities (dof=3)", cex.main=.85, cex.axis=.70, cex.lab=.75)
# I am using 3 degrees of freedom for the Student-t distribution
curve(dt(x*sqrt(3)/hist_std,df=3)*sqrt(3)/hist_std,from=-.05,to=.05,col="red", add=TRUE, lty=2,lwd=2, yaxt="n")
curve(dnorm(x, mean=hist_mean, sd=hist_std), from=-.05,to=.05,col="blue", add=TRUE, lty=2,lwd=2, yaxt="n")
legend("topleft",legend=c("Normal Density","Student-t Density (dof=3)"),bty="n", lty=c(2,2), cex=0.80, col=c("blue","red")) 

# QQ Plots with Normal and Student-t quantiles
qqPlot(coredata(SPX.log.returns),distribution = "norm", main = "QQ Plot for SPX Log Returns - Normal Distribution", 
       xlab="Normal Quantiles", ylab="Log Return Quantiles", cex=.5, envelope = .95, grid=FALSE)
qqPlot(coredata(SPX.log.returns),distribution = "t", df=3,main = "QQ Plot for SPX Log Returns - Student-t Distribution (dof=3)",
       xlab="Student-t Quantiles",ylab="Log Return Quantiles", cex=.5, envelope = .95, grid=FALSE)


# ACFs for log returns and then squared log returns
par(mfrow=c(2,1))
par(mar=c(4,4,3,4))
acf(SPX.log.returns,lag.max=20, xlab = "Lag", main="SPX Log Returns", ci.col="blue",lty=3,lwd=4,col="red")
mtext(side=3,'(Weak to no autocorrelation \n in the ACF of SPX log-returns)', col="black", cex=.75)
acf(SPX.log.returns^2,lag.max=20, xlab = "Lag", main="SPX Log Returns Squared", ci.col="blue",lty=3,lwd=4,col="red")
mtext(side=3,'(Clear pattern of autocorrelation \n in the ACF of SPX log-returns squared)', col="black", cex=.75)

# Function to calculate rolling annualized historical volatility and to get 
# tabular statistics for realized one-month rolling volatility values
# n.roll.real.vol is used for number of trading days to calculate the rolling realized 
# annualized monthly volatility and n is used to subset the time series to drop 
# the first n+1 days as we can only start calculation after some number of observations
n.roll.real.vol <- 22

Roll.real.vol <- function(data,n){
  
  vol <- data[n:length(data)]
  
  for(i in n:length(data)) {
    
    # Here we use simple sum of squared returns with mean assumed
    # equal to zero which is consistent with certain academic and
    # practitioner literature.  We annualize thereby making the realized
    # volatility comparable to the implied volatility index values
    vol[(i-n)+1] <- sqrt(sum(data[((i-n)+1):i]^2)/(n))*sqrt(252)
    
  }
  
  vol
  
}

# Calculate rolling realized volatility for log returns
SPX.rolling.vol <- Roll.real.vol(SPX.log.returns.df$spx,n.roll.real.vol)

# Summary statistics for rolling actual realized one-month volatility
SPX.real.vol.obs <- length(SPX.rolling.vol)
SPX.real.vol.min <- min(SPX.rolling.vol)
SPX.real.vol.mean <- mean(SPX.rolling.vol)
SPX.real.vol.median <- median(SPX.rolling.vol)
SPX.real.vol.max <- max(SPX.rolling.vol)

# Matrix of outputs
Summary.real.vol.stats <- as.data.frame(matrix(data=NA,nrow = 5, ncol = 1))

Summary.real.vol.stats[1,1] <- SPX.real.vol.obs
Summary.real.vol.stats[2,1] <- SPX.real.vol.min
Summary.real.vol.stats[3,1] <- SPX.real.vol.mean
Summary.real.vol.stats[4,1] <- SPX.real.vol.median
Summary.real.vol.stats[5,1] <- SPX.real.vol.max

rownames(Summary.real.vol.stats) <- c("No. Obs.","Min", "Mean", "Median", "Max")
colnames(Summary.real.vol.stats) <- c("SPX")
Summary.real.vol.stats

# GARCH modeling begins here
# Log returns data to use for GARCH modeling
SPX.data <- SPX.log.returns.df$spx*100

# n.GARCH.start is used for number of trading days used to fit the model
# n.GARCH.end is so we end one month before end of data 
n.GARCH.start <- 1260
n.GARCH.end <- 22

# GARCH Modeling
# GARCH spec to use in loop - GARCH(1,1) with normal innovations
GARCH.spec.norm <- ugarchspec(mean.model=list(armaOrder=c(0,0)),
                        variance.model=list(garchOrder=c(1,1)),distribution="norm") 

# Function to calculate rolling historical volatility estimator for GARCH(1,1)
# with normal innovations
Roll.GARCH.norm.est <- function(data,n.GARCH.start,n.GARCH.end){
  
  nlast <- length(data)-n.GARCH.end
  GARCH.est <- data[n.GARCH.start:nlast]
  
  for(i in n.GARCH.start:(length(data)-n.GARCH.end)) {
    
    print(i)
    
    data.set <- data[((i-n.GARCH.start)+1):i]
    # Fit GARCH model
    GARCH.fit <- ugarchfit(data=data.set,spec=GARCH.spec.norm, solver="hybrid") 
    # Forecast 22 days ahead - assumption is 22 days is one-month
    GARCH.forecast <- ugarchforecast(GARCH.fit,data=data.set,n.ahead=22)
    # Calculate forecasted annualized volatilty for the one-month horizon
    # We will compare the forecasts to actual realized volatility over the same
    # time period
    GARCH.est[(i-n.GARCH.start)+1] <- sqrt(mean(sigma(GARCH.forecast)^2))*sqrt(252)
    
    
  }
  
  GARCH.est
  
}

# Rolling estimates for GARCH(1,1) based on normal innovations
# This will take some time as it fits the model using 5 years of returns data
# and then forecasts one month forward, rolls forward one day and fits again
# to forecast one month forward again, etc.  You don't use the last 22 days
# for fitting as you need to forecast for the last month and simply observe versus
# what is realized
SPX.GARCH.norm.vol <- Roll.GARCH.norm.est(SPX.data,n.GARCH.start,n.GARCH.end)

# Create data frame with dates, forecasted volatility values, and realized volatility values
# Recall that there is one more date as it comes from original price data which when
# we compute returns we lose one date
SPX.results.dates <- as.Date(SPX.df.date[1259:5503], format="%m/%d/%Y")
SPX.results.data.frame <- cbind(SPX.rolling.vol[1258:5502],SPX.GARCH.norm.vol/100)

# Mean Squared Error (MSE) calculation
SPX.hist.mse <- mean((SPX.results.data.frame[,2] - SPX.results.data.frame[,1])^2,na.rm=TRUE)

# Plot of forecasted versus realized for SPX
par(mar=c(6, 12, 4, 12) + 0.1)
plot(SPX.results.dates, SPX.results.data.frame[,1], axes=F, ylim=c(-.2,1.0),type="l",col="black", xlab="", ylab="")
axis(2, ylim=c(-.2,1.0),col="black",lwd=2)
mtext(2,text="Realized Volatility Level",line=2.0)

par(new=T)
plot(SPX.results.dates, SPX.results.data.frame[,2], axes=F, ylim=c(-.2,1.0), type="l",lty=2,lwd=2, xlab="", ylab="", col="red", xaxt="n")
axis(2, ylim=c(-.2,1.0),lwd=2,line=3.5)
mtext(2,text="Predicted Volatility Level",line=5.5)

difference <- (SPX.results.data.frame[,2] - SPX.results.data.frame[,1])
mse <- mean((SPX.results.data.frame[,2] - SPX.results.data.frame[,1])^2,na.rm=TRUE)

par(new=T)
plot(SPX.results.dates, difference, axes=F, ylim=c(-.2,1.0), type="l",lty=2,lwd=2, xlab="", ylab="", col="blue", xaxt="n")
abline(h=mse,col='grey50',lty=2,lwd=2)
axis(4, ylim=c(-.2,1.0),lwd=2,line=3.5)
mtext(4,text="SPX Realized - SPX Predicted",line=7.0)

title(main="SPX Realized Volatility and  SPX Predicted Volatility")
axis.Date(side=1,SPX.results.dates, at=seq(from=min(SPX.results.dates),to=max(SPX.results.dates),by=50), format="%m/%d/%Y")
mtext("Date",side=1,col="black",line=2)

legend("topleft",legend=c("SPX Realized Volatility Level","SPX Predicted Volatility", "SPX Realized - SPX Predicted", "MSE"),bty="n",lty=c(1,2,3,2), cex=0.60, col=c("black","red","blue","grey50")) 

