# Markowitz Optimization

# Include library with solve.QP()
install.packages("quadprog")
library(quadprog)

################################################################################################
# If you copy the expected return vector to the clipboard then run this first part of code below 
# you can use this to bring in the vector of expected returns
read.excel.mu <- function(header=FALSE,...) {
  read.table("clipboard",sep="\t")
}

excel.data.mu <- read.excel.mu()
excel.data.mu
################################################################################################

# If you copy the covariance matrix to clipboard then run this second part of code below you can use 
# this to bring in the covariance matrix
read.excel.V <- function(header=FALSE,...) {
  read.table("clipboard",sep="\t")
}

excel.data.V <- read.excel.V()
excel.data.V
################################################################################################

# This assumes that you used the prior two functions to bring in the vector of expected reutrns
# and the covariance matrix
# Expected returns and variance/covariance matrix
mu <- c(excel.data.mu$V1, excel.data.mu$V2, excel.data.mu$V3, excel.data.mu$V4, excel.data.mu$V5,
        excel.data.mu$V6, excel.data.mu$V7, excel.data.mu$V8, excel.data.mu$V9, excel.data.mu$V10,
        excel.data.mu$V11, excel.data.mu$V12)
V <- cbind(excel.data.V$V1, excel.data.V$V2, excel.data.V$V3, excel.data.V$V4, excel.data.V$V5,
       excel.data.V$V6, excel.data.V$V7, excel.data.V$V8, excel.data.V$V9, excel.data.V$V10,
       excel.data.V$V11, excel.data.V$V12)
mu
V
################################################################################################

# Function for the Efficient Frontier
# Need quadprog for this but it is already
# included earlier in file
Frontier <- function(n_assets,mu_vector,V_matrix){
  
# creating a vector of returns to use to sketch out the frontier
mu_frontier <- seq(min(mu_vector)+.001, max(mu_vector)-.001, by=0.001)

# Get length of returns vector used to sketch frontier
m <- length(mu_frontier)
# Set up vector for sigma values with length of m so we have pairs of (mu,sigma) to plot
sigma_vector <- rep(0,m)
# Set up matrix to store weights with m rows and n (input) the number of columns
weights_matrix <- matrix(0,nrow=m,ncol=n_assets)

# d matrix which is zero for this case
d <- rep(0,n_assets)
# Set up matrix which will have diagonal values of one
tmp <- matrix(0,nrow=n_assets,ncol=n_assets)
diag(tmp) <- 1

# Set up A matrix - recall that the problem is: min(-d'b + 1/2*b'Db), s.t. A'b>=b0
# the first two columns of A vector are equality conditions and the rest are short
# selling constraints
A <- cbind(rep(1,n_assets),mu_vector,tmp)

# For loop to solve for weights on frontier for array of mu values
for(i in 1:m){
  
  # b0 vector is so sum of weights is one, portfolio return is mu_frontier and each weight is larger 
  # than zero - no short sales
  b0 <- c(1,mu_frontier[i],rep(0,n_assets))
  # Solve for optimal weights for each portfolio return mu_frontier
  Z <- solve.QP(V_matrix,d,A,b0,meq=2)
  sigma_vector[i] <- sqrt(2*Z$value)
  weights_matrix[i,] <- Z$solution
  
}

# Plot the frontier
plot(sigma_vector, mu_frontier, xlim=c(0,max(sqrt(diag(V_matrix)))),"l",
     xlab="Standard Deviation of the Portfolio", ylab="Mean of the Portfolio",
     main="Efficient Frontier with No-Short Sales Constraint")
list <- paste("weights_matrix", 1:n_assets)
barplot(t(weights_matrix), names.arg=mu_frontier,xlab="Mean of the Portfolio",ylab="Weights for Assets",
        border=0,legend=list, args.legend=list(x="topright"))

return(cbind(weights_matrix, mu_frontier, sigma_vector))

}

# Call to function using prior mu and V values
Frontier(12,mu,V)

