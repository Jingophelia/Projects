# Option Pricing Example

# Include package to connect to Excel
install.packages("openxlsx")
library(openxlsx)

# Path, file, sheet names
Excel_file_path <- "C:/Users/Stephen D Young/Documents/Stephen D. Young/Teaching/GWU 2018/R Code/"
Excel_file_name <- "Black_Scholes.xlsx"
Excel_full_file_name <- paste(Excel_file_path,Excel_file_name,sep="")
Excel_data_sheet_name <- "Output Data"
Excel_graphs_sheet_name <- "Graphs"

# Load the workbook object once to prevent plots showing up as blanks
Excel_workbook <- openxlsx::loadWorkbook(Excel_full_file_name)


# Number of digits for results
options(digits=8)

# Black-Scholes Option 
BlackScholes <- function(S, X, rf, d, T, sigma, flag) {
  
  # d1 and d2 values
  F <- S * exp((rf - d)*T)
  d1 <- (log(F/X) + (.5*(sigma^2) * T))/(sigma * sqrt(T))
  d2 <- (log(F/X) - (.5*(sigma^2) * T))/(sigma * sqrt(T))
  
  # European call and put values and Greeks
  call <- exp(-rf * T) * (F * pnorm(d1) - X * pnorm(d2))
  put <- exp(-rf * T) * (X * pnorm(-d2) - F * pnorm(-d1))
  
  calldelta <- exp(-d * T) * pnorm(d1)
  
  if(flag=='call'){
    
    return(call)
  } else if(flag=='put') {
    
    return(put)
  } else if(flag=='calldelta'){
    
    return(calldelta)
    
  }
  
}

# Sequence of S values to call function followed by inputs and function call
S_sequence <- seq(from = 50, to = 150, by = 1)
X <- 100
rf <- .05
d <- 0
T_one_year <- 1
T_half_year <- 0.5
T_maturity <- .0001
sigma <- .20
flag <- 'call'

Calls_one_year <- BlackScholes(S_sequence, X, rf, d, T_one_year, sigma, flag)
Calls_half_year <- BlackScholes(S_sequence, X, rf, d, T_half_year, sigma, flag)
Calls_maturity <- BlackScholes(S_sequence, X, rf, d, T_maturity, sigma, flag)


# Write results to spreadsheet
openxlsx::writeData(Excel_workbook, Excel_data_sheet_name, Calls_one_year, startCol=2, startRow = 2, colNames=FALSE,
                    rowNames=FALSE, keepNA = TRUE)
openxlsx::writeData(Excel_workbook, Excel_data_sheet_name, Calls_half_year, startCol=3, startRow = 2, colNames=FALSE,
                    rowNames=FALSE, keepNA = TRUE)
openxlsx::writeData(Excel_workbook, Excel_data_sheet_name, Calls_maturity, startCol=4, startRow = 2, colNames=FALSE,
                    rowNames=FALSE, keepNA = TRUE)

# Create plot which we will save to spreadsheet 
Calls <- plot(x = S_sequence, y = Calls_one_year, main = "Black-Scholes Call Values", xlab = "Spot Price",
          ylab = "Option Value", xlim = c(50,150), ylim = c(0, max(Calls_one_year)), type = "l", col = "blue", lwd = 2)
          lines(x = S_sequence, y = Calls_half_year, type = "l", col = "red", lwd = 2, lty = 2)
          lines(x = S_sequence, y = Calls_maturity, type = "l", col = "green", lwd = 2, lty = 2)
          legend(x="topleft", legend=c("1 Yr Call Values","1/2 Yr Call Values", "Maturity Call Values"),
                 bty="n",lty=c(1,2,2), lwd=c(2,2,2), cex=0.75, col=c("blue","red","green"))
          
          
# Print pic to view and insertPlot into Excel
print(Calls)
openxlsx::insertPlot(Excel_workbook,sheet=Excel_graphs_sheet_name,width=7,height=5,startRow=2,startCol=2,fileType="bmp",units="in",dpi=600)

# Save the workbook object exactly once to prevent plots showing up as blanks
openxlsx::saveWorkbook(Excel_workbook,Excel_full_file_name,overwrite=TRUE)