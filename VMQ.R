Analyze <- function() {
  
  #Get SIPro folder location from user, currently titled "12.17.21 Data"
  #The dialoge box is hidden in the back behind open windows
  library(tcltk)

  d <- tk_choose.dir(default = getwd(), caption = "Select Folder")
  
  #List all .csv files to see if they're in there
  f <- GetFiles(d)
  
  #Read data
  Data <- read.csv(f[1], header = F) #SIPro .csv data w/no headers
  Key <- read.csv(f[2], header = F) #has headers
  #stop("Stopped programmatically by Vitaly")
  
  #Rename variables (columns)    
  colnames(Data) <- Key[,2]
  
  Key[,2]
  
  #CALCULATE VARIABLES
  
  L <- length(Data[,1]) # Number of obs. is used a lot herein
  
  #B/P Percentile
  Data[,"BP.Pct"] <- percentile(Data[,"B/P"])
  # E/P Percentile 
  Data[,"EP.Pct"] <- percentile(Data[,"E/P (Continuing & Diluted)"])
  # S/P Percentile
  Data[,"SP.Pct"] <- percentile(Data[,"S/P"])
  # EBITDA/EV
  Data[,"EBITAEV.Pct"] <- percentile(Data[,"EBITDA/EV"])
  # CF/P Percentile
  Data[,"CFP.Pct"] <- percentile(Data[,"CFPS/P"])
  # S.H. Yeild Percentile
  Data[,"SHY.Pct"] <- percentile(Data[,"Shareholder Yield"])
  
  # Calculate Composite Value Factors (CVF2)
  
  for (i in 1:L) { 
    r <- as.numeric(Data[i, c("BP.Pct","EP.Pct","SP.Pct",
                              "EBITAEV.Pct","CFP.Pct","SHY.Pct")])
    if (length(r[is.na(r)]) > 2) {
      #assign .5 to NA's if there are more than 2 NA's
      r[is.na(r)] <- .5
      Data[i, "CVF2"] <- mean(r)
      #print(mean(r))
    } else {
      #average available percentiles
      Data[i, "CVF2"] <- mean(r[!is.na(r)])
      #print(mean(r[!is.na(r)]))
    }
    #i automatically increments i
  }  
  
  # Accruals to Assets Percentile (minus from 1 when lower is better)
  Data[,"ATA.Pct"] <- 1 - percentile(Data[,"Accruals to Assets"])
  # Pct. Chg. in NOA Percentile
  Data[,"PCNOA.Pct"] <- 1 - percentile(Data[,"Pct. Chg. in NOA"])
  # Accruals to Average Assets Percentile
  Data[,"ATAA.Pct"] <- 1 - percentile(Data[,"Accruals to Average Assets"])
  # CapEx to Deprec.Amort. Percentile
  Data[,"PCNOA.Pct"] <- 1 - percentile(Data[,"CapEx to Deprec.Amort."])
  
  # Calculate Earnings Quality Composite (EQC)
  
  for (i in 1:L) { 
    r <- as.numeric(Data[i, c("ATA.Pct","PCNOA.Pct",
                              "ATAA.Pct","PCNOA.Pct")])
    if (length(r[is.na(r)]) > 1) { 
      #assign .5 to NA's if there are more than 1 NA's
      r[is.na(r)] <- .5
      Data[i, "EQC"] <- mean(r)
      #print(mean(r))
    } else {
      #average available percentiles
      Data[i, "EQC"] <- mean(r[!is.na(r)])
      #print(mean(r[!is.na(r)]))
    }
    #i automatically increments i
  }  
  
  #Create value and quality composite for R&D
  Data[, "VMQ_Test"] <- Data$CVF2*.8 + Data$EQC*.2
  
  # ROE Percentile
  Data[,"ROE.Pct"] <- percentile(Data[,"Return on equity 12m"])
  # Asset Turnover Percentile
  Data[,"Turnover.Pct"] <- percentile(Data[,"Asset turnover 12m"])
  
  # Coverage Ratio Percentile
  Data[,"Coverage.Pct"] <- percentile(Data[,"Times interest earned 12m"])
  # Asset to Equity Percentile
  Data[,"AE.Pct"] <- percentile(Data[,"A/E"])
  # External Financing (CF.fin/Avg.A) Percentile
  Data[,"ExFin.Pct"] <- 1 - percentile(Data[,"CF.FINANCING"]) 
  # Debt Change Percentile
  Data[,"DebtChg.Pct"] <- 1 - percentile(Data[,"1Y.DEBT.CHG"])
  # 1-year Op. Income Percent Change Percentile
  Data[,"OpIncChg.Pct"] <- percentile(Data[,"1Y.OP.INCOME.GROWTH"]) 
  
  # Earnings Yeild (EBIT/EV)
  Data[,"EY.Pct"] <- percentile(Data[,"EY"]) 
  # Return on Capital (EBIT/(Net Working Capital + Fixed Assets))
  Data[,"ROC.Pct"] <- percentile(Data[,"ROC"]) 
  # Magic Formula: ROC.Pct + EY.Pct
  Data[,"Magic.Pct"] <- (Data$ROC.Pct + Data$EY.Pct) / 2
  
  # RSI Percentile
  Data[,"RSI.Pct"] <- percentile(Data[,"Relative Strength 26 week"])
  
  # Obtain Top 50 for the following 5 scenarios
  VMQ.VEN.50 <- GetTop(Data, "CVF2", .05, L, 50)
  VMQ.DEC.50 <- GetTop(Data, "CVF2", .10, L, 50)
  MAGIC.VEN.50 <- GetTop(Data, "Magic.Pct", .05, L, 50)
  
  #MAGIC_ONLY.VEN.50 <- GetTop(Data, "Magic.Pct", .05, L, 50)
  
  #VMQ_Test.VEN.50 <- GetTop(Data, "VMQ_Test", .05, L, 50)
  #VMQ_Test.DEC.50 <- GetTop(Data, "VMQ_Test", .10, L, 50)
  
  
  #stop("Stopped before save occurred")
  
  
  #SAVE
  cat("The following has been saved in:", d, "\n")
  date <- format(Sys.time(), "%m.%d.%Y") #used in file names
  obj <- c("VMQ.VEN.50", "VMQ.DEC.50", "MAGIC.VEN.50", "date")
  len <- length(obj) - 1 # Used in for loop for saving .xlsx files
  # make sure this includes only .xlsx objects
  
  # Save as .RData
  r <- paste(d, paste(date, "Analysis.R", sep = " "), sep = "/")
  dump(obj, file = r) #data to 'source()' later
  cat("- Analysis.R\n")
  
  # Save as .xlsx
  suppressMessages(require(XLConnect))
  
  for (i in 1:len) { 
    r <- paste(d, "/", date, " ", obj[i], ".xlsx", sep = "")
    wb <- loadWorkbook(r, create = TRUE)
    createSheet(wb, date)
    writeWorksheet(wb, get(obj[i]), sheet = date, startRow = 1, startCol = 1)
    saveWorkbook(wb)
    cat("- ", obj[i], ".xlsx\n", sep = "") # Print message
  }
  
  # r <- paste(d, paste(date, "VMQ.VEN.50.xlsx", sep = " "), sep = "/")
  # wb <- loadWorkbook(r, create = TRUE)
  # createSheet(wb, date)
  # writeWorksheet(wb, VMQ.VEN.50, sheet = date, startRow = 1, startCol = 1)
  # saveWorkbook(wb)
  # cat("- VMQ.VEN.50.xlsx\n")
  
}

Reflection <- function() {
  #Set working directory, if getwd() is in 'My Documents'
  try(setwd("./R/My Projects/Stock Analysis"), silent = TRUE)
  #user selects folder
  #d <- "./8.6.18 Data"  #placeholder (default path)
  d <- choose.dir(default = getwd(), caption = "Select folder")
  
  #run check if Reflection.R file exists
  r <- "*Reflection.R"
  f <- list.files(path = d, pattern = r, recursive = FALSE)
  
  if (length(f)>0) {
    r <- cat("'", f, "' already exist(s).\nDo you wish",
             " to overwrite or add new files? (y/n): ", 
             sep = "")
    r <- readline(prompt = r)
    
    if (r == "y") {
      #proceed
    } else {
      stop("Find the correct folder to analyze.")
    }
  }
  
  #Check if Analysis.R exists and source()
  r <- "*Analysis.R"
  f <- list.files(path = d, pattern = r, recursive = 
                    FALSE, full.names = TRUE)
  
  if (length(f) > 1) {
    #Get user input 
    cat("Use browser to select your '*.R' file.")
    f <- choose.files(default = d, caption = "Select file",
                      multi = FALSE)
  } else {
    #otherwise load the only available analysis.r
  }
  cat("Loading '", f, "'\n", sep = "")
  source(f)
  
  #Get tickers
  ##use nested for loop to fill "t" to improve
  t <- data.frame(VV50 = VMQ.VEN.50$Ticker,
                  VD50 = VMQ.DEC.50$Ticker,
                  VTV50 = VMQ_Test.VEN.50$Ticker,
                  VTD50 = VMQ_Test.DEC.50$Ticker)
  
  #Update Prices for the data frame of tickers
  old.date <- date
  p <- Prices(t, old.date)  #price function returns data.frame
  ##use nested for loop to fill, it'll be better
  VMQ.VEN.50$`13 Mo Price` <- p[,1]
  VMQ.DEC.50$`13 Mo Price` <- p[,2]
  VMQ_Test.VEN.50$`13 Mo Price` <- p[,3]
  VMQ_Test.DEC.50$`13 Mo Price` <- p[,4]
  
  #Calculate ROI (No Dividends Included)
  VMQ.VEN.50$ROI <- (VMQ.VEN.50$`13 Mo Price`/ VMQ.VEN.50$Price - 1)*100
  VMQ.DEC.50$ROI <- (VMQ.DEC.50$`13 Mo Price`/ VMQ.DEC.50$Price - 1)*100
  VMQ_Test.VEN.50$ROI <- (VMQ_Test.VEN.50$`13 Mo Price`/ VMQ_Test.VEN.50$Price - 1)*100
  VMQ_Test.DEC.50$ROI <- (VMQ_Test.DEC.50$`13 Mo Price`/ VMQ_Test.DEC.50$Price - 1)*100
  
  #cat("CVF2.dec.50 ROI: ", mean(CVF2.dec.50$ROI))
  #cat("VMQ.dec.50 ROI: ", mean(VMQ.dec.50$ROI)) 
  cat("\n")
  
  #Print VMQ.VEN.50 Report
  cat("\nVMQ.VEN.50 ROI: ", round(mean(VMQ.VEN.50$ROI), 2),
      "%", sep = "")
  r <- VMQ.VEN.50[, c("Ticker", "CVF2","Performance Value",
                      "Safety Margin", "ROI", "Keep")]
  r[,2:length(r)]<-apply(r[,2:length(r)], 2, function(x) round(x, 2)) #round each column ("2") in r
  rownames(r) <- 1:length(r[,1])
  cat("\n\n")
  print(r)
  cat("\n")
  
  #Print VMQ.DEC.50 Report
  cat("\nVMQ.DEC.50 ROI: ", round(mean(VMQ.DEC.50$ROI), 2),
      "%", sep = "")
  cat("\n\n")
  r <- VMQ.DEC.50[, c("Ticker", "CVF2","Performance Value",
                      "Safety Margin", "ROI", "Keep")]
  r[,2:length(r)]<-apply(r[,2:length(r)], 2, function(x) round(x, 2)) #round each column ("2") in r
  rownames(r) <- 1:length(r[,1])
  print(r)
  cat("\n")
  
  #Print VMQ_Test.VEN.50 Report
  cat("\nVMQ_Test.VEN.50 ROI: ", round(mean(VMQ_Test.VEN.50$ROI), 2),
      "%", sep = "")
  cat("\n\n")
  r <- VMQ_Test.VEN.50[, c("Ticker", "VMQ_Test","Performance Value",
                           "Safety Margin", "ROI", "Keep")]
  r[,2:length(r)]<-apply(r[,2:length(r)], 2, function(x) round(x, 2)) #round each column ("2") in r
  rownames(r) <- 1:length(r[,1])
  print(r)
  cat("\n")
  
  #Print VMQ_Test.DEC.50 Report
  cat("\nVMQ_Test.DEC.50 ROI: ", round(mean(VMQ_Test.DEC.50$ROI), 2),
      "%", sep = "")
  cat("\n\n")
  r <- VMQ_Test.DEC.50[, c("Ticker", "VMQ_Test","Performance Value",
                           "Safety Margin", "ROI", "Keep")]
  r[,2:length(r)]<-apply(r[,2:length(r)], 2, function(x) round(x, 2)) #round each column ("2") in r
  rownames(r) <- 1:length(r[,1])
  print(r)
  cat("\n")
  
  
  # SAVE
  r <- cat("Do you want to save the .R and .xlsx files in: '", d,"'?\n(y/n): ", 
           sep = "")
  r <- readline(prompt = r)
  if (r == "y") {
    # Save as .RData
    date <- format(Sys.time(), "%m.%d.%Y") #used in file name
    r <- paste(d, paste(date, "Reflection.R", sep = " "), sep = "/")
    dump(c("VMQ.VEN.50", "VMQ.DEC.50","VMQ_Test.VEN.50", 
           "VMQ_Test.DEC.50", "date"), file = r) #data to 'source()' later
    cat("The following files have been saved.\n")
    cat("- Reflection.R\n")
    
    # Save as .xlsx
    suppressMessages(require(XLConnect))
    
    r <- paste(d, paste(old.date, "VMQ.VEN.50.xlsx", sep = " "), sep = "/")
    wb <- loadWorkbook(r, create = FALSE)
    #createSheet(wb, date)
    renameSheet(wb, old.date, date)
    writeWorksheet(wb, VMQ.VEN.50, sheet = date, startRow = 1, startCol = 1)
    saveWorkbook(wb)
    cat("- VMQ.VEN.50.xlsx\n")
    
    r <- paste(d, paste(old.date, "VMQ.DEC.50.xlsx", sep = " "), sep = "/")
    wb <- loadWorkbook(r, create = FALSE)
    #createSheet(wb, date)
    renameSheet(wb, old.date, date)
    writeWorksheet(wb, VMQ.DEC.50, sheet = date, startRow = 1, startCol = 1)
    saveWorkbook(wb)
    cat("- VMQ.DEC.50.xlsx\n")
    
    # r <- paste(d, paste(old.date, "VMQ_Test.VEN.50.xlsx", sep = " "), sep = "/")
    # wb <- loadWorkbook(r, create = FALSE)
    # #createSheet(wb, date)
    # renameSheet(wb, old.date, date)
    # writeWorksheet(wb, VMQ_Test.VEN.50, sheet = date, startRow = 1, startCol = 1)
    # saveWorkbook(wb)
    # cat("- VMQ_Test.VEN.50.xlsx\n")
    
    # r <- paste(d, paste(old.date, "VMQ_Test.DEC.50.xlsx", sep = " "), sep = "/")
    # wb <- loadWorkbook(r, create = FALSE)
    # #createSheet(wb, date)
    # renameSheet(wb, old.date, date)
    # writeWorksheet(wb, VMQ_Test.DEC.50, sheet = date, startRow = 1, startCol = 1)
    # saveWorkbook(wb)
    # cat("- VMQ_Test.DEC.50.xlsx\n\nFiles in Directory:\n")
    # print(dir(d))
    
  } else {
    cat("\nSee you later dude, good luck.\n\n")
  }
  
  #************************************************#
  #     Automated Regression Analysis and Report   #
  #                                                #
  #************************************************#
  
  #Insert code here after you actually run the regression
  #analysis yourself. You'll know what to add here after
  #you do this.
  
  #SAVE
  # cat("The following has been saved in:", d, "\n")
  # date <- format(Sys.time(), "%m.%d.%Y") #used in file names
  # 
  # # Save as .RData
  # r <- paste(d, paste(date, "Reflection.R", sep = " "), sep = "/")
  # dump(c("VMQ.50", "CVF2.50", "date"), file = r) #data to 'source()' later
  # cat("- Reflection.R\n")
  
}

GetFiles <- function(d) {
  #See if analysis was alread run on this folder
  #***place in GetFiles()
  f <- list.files(path = d, pattern = "*Analysis.R", recursive = FALSE)
  
  if (length(f) != 0) {
    r <- cat("Are you sure you wish to overwrite '", f, "'?\n(y/n): ", 
             sep = "")
    r <- readline(prompt = r)
    if (r == "y") {
      #proceed
    } else {
      stop("Find the correct folder to analyze.")
    }
  }
  
  #gather all .xls files 
  f <- list.files(path=d, pattern="*.TXT",
                  full.names=T, recursive=FALSE)
  f
  
  #gather only ones with 'DATA' in name
  f <- f[grep("DATA", f)] #SIPro capitalizes 'Data'
  
  #make sure only 2 usable files
  if (length(f) != 2) {
    stop(paste("Add the two Stock Investor Pro excel data",
               "workbooks to the folder before running",
               "analysis.", sep = " "))
  } else {
    if (1 == grep("Key",f)) {
      x[2] <- f[1]
      x[1] <- f[2]
      f <- x
      rm(x)
    } else {
      f
    }
  }
}

percentile <- function(v) {
  #create rank vector: 'keep's na's in place and assigns
  #'average' percentile to values that are identical/tied.
  v.pctrank <- rank(v, na.last = "keep", ties.method = "average") /
    length(v[!is.na(v)]) # divide rank number by count of all 
  # existing numbers
  v.pctrank
}

Prices <- function(t, old.date = NULL) {
  #check date range
  if (!is.null(old.date)) {
    old.date <- as.Date(old.date, "%m.%d.%Y")
    r <- as.numeric(as.Date(Sys.time()) - old.date)
    cat("It's been", r, "days since the Analysis()",
        "was done.\n")
    
    if (r < 360) {
      r <- cat("Do you still wish to continue? (y/n): ")
      r <- readline(prompt = r)
      if (r == "y") {
        #Proceed
      } else {
        stop("See you later dude.")
      }
    }
  }
  
  #Get prices for tickers in data.frame
  cat("Booting 'quantmod'...")
  suppressMessages(require(quantmod))
  
  print(length(t$VV50))
  print(length(t$VD50))
  print(length(t$VTV50))
  print(length(t$VTD50))
  # VV50 = VMQ.VEN.50$Ticker,
  # VD50 = VMQ.DEC.50$Ticker,
  # VTV50 = VMQ_Test.VEN.50$Ticker,
  # VTD50 = VMQ_Test.DEC.50$Ticker)
  
  for (i in 1:length(t)) {
    r <- getQuote(as.character(t[,i]), what = 
                    yahooQF("Previous Close")) 
    print(head(r))
    t[,i] <- r[,2] #replace tickers ("t") with prices
    #u <- readline(prompt = "Hit enter to download next batch.")
  }
  
  #Return ordinal data.frame w/prices
  t
}

GetTop <- function(Data, Var, Pct, L, Top) {
  #Data = the stock data
  #Var = the variable name to rank the stock data by
  #Pct = the percentage of the highest ranked stocks to select 
  #for further analysis
  #L = length of the data
  #Top = number of the top ranked stocks to output
  
  #Sort Data by the selected variable in desending order
  r <- order(Data[,Var], decreasing = TRUE)
  out <- Data[r,]
  
  #Remove Utilities and Financials if ranking by Magic Formula
  # otherwise remove OTC stocks
  if (Var == "Magic.Pct") {
    out <- out[grepl("financials", out$Sector, TRUE) == FALSE,]
    out <- out[grepl("utilities", out$Sector, TRUE) == FALSE,]
  } else {
    out <- out[grepl("counter", out$Exchange, TRUE) == FALSE,]
  }
  
  # Confine data to highest ranked percentile
  r <- round(L*Pct, digits = 0) # obtain quantity
  out <- out[1:r,]
  
  # Sort percentile data by momentum and select top 50
  r <- order(out[,"Relative Strength 26 week"], 
             decreasing = TRUE)
  out <- out[r,]
  out <- out[1:Top,]
  
  # Mark stocks that fail the following criteria
  
  out[,"Keep"] <- out$ROE.Pct       > 0.2 & # rm. lowest 20% ROE
    out$Coverage.Pct  > 0.2 & # rm. lowest 20% Coverage ratio
    out$EQC           > 0.3 & # rm. lowest 30% EQC
    out$Turnover.Pct  > 0.2 & # rm. lowest 20% Asset Turnover
    out$AE.Pct        > 0.1 & # rm. lowest 10% Asset/Equity
    out$ExFin.Pct     > 0.3 & # rm. lowest 20% (inverse) External Financing
    out$DebtChg.Pct   > 0.2 & # rm. lowest 20% (inverse) Debt Change
    out$RSI.Pct       > 0.3 & # rm. lowest 30% 6 mo. Price Apprec.
    out$OpIncChg.Pct  > 0.2 & # rm. lowest 20% Operating Income Chg. (Instead of EPS)
    out$Magic.Pct     > 0.2 & # rm. lowest 20% of Magic Formula
    out$'F Score TTM' > 2     # rm. below 3 
  
  #out$'Z Score Q1'  > 1.8   # rm. below 1.8
  
  #Quality Rank = Momentum + Quality + Safety + Profitability
  r <- read.csv("Returns.csv", header = TRUE)
  
  # output each historical return for R&D
  out[, "rsi"] <- rsi <- Assign(out$RSI.Pct, r$RSI_Return, Top)
  out[, "eqc"] <- eqc <- Assign(out$EQC, r$EQC_Return, Top)
  out[, "cov"] <- cov <- Assign(out$Coverage.Pct, r$Coverage_Return, Top)
  out[, "ae"] <- ae <- Assign(out$AE.Pct, r$AE_Return, Top)
  out[, "exf"] <- exf <- Assign(out$ExFin.Pct, r$ExFin_Return, Top)
  out[, "dch"] <- dch <- Assign(out$DebtChg.Pct, r$DebtChg_Return, Top)
  out[, "roe"] <- roe <- Assign(out$ROE.Pct, r$ROE_Return, Top)
  out[, "tur"] <- tur <- Assign(out$Turnover.Pct, r$Turnover_Return, Top)
  out[, "mag"] <- mag <- Assign(out$Magic.Pct, r$Magic_Return, Top)
  
  #momentum
  m <- rsi
  
  #quality
  q <- (eqc + mag/4)/(5/4) #incorporate magic formula
  
  #safety
  s <-  (cov + ae + exf + dch)/4
  
  #profitability
  p <-  (roe + tur)/2
  # % Change in Oerating Margins is not included,
  # because there is no backtest performance data yet
  
  out[, "MQSP Rank"] <- (m + q + s + p) / 4
  
  #Calculate the Performance Val and M.O.S.
  BV <- out$'Book value/share Q1'
  PE <- out$'7Y.AVG.PE2'
  div <- out$'7Y.AVG.DIV'
  
  rate <- 0.03  # Discount rate
  
  BV.10 <- BV * (1 + out$G/100)^10  #BV after 10y
  EPS.10 <- out$AVG.ROE/100 * BV.10 #EPS after 10y
  P.10 <- PE * EPS.10               #Price after 10y
  
  #The present value is the performance value
  require("stringi") #"FinCal" package relies on this package
  require("FinCal")
  
  #PV of P.10: pv(r, n, fv = 0, pmt = 0, type = 0)
  # r	= discount rate, or the interest rate at which the amount 
  #     will be compounded each period
  # n	= number of periods
  # fv	= future value
  # pmt	= payment per period
  # type = payments occur at the end of each period (type=0); 
  #        payments occur at the beginning of each period (type=1)
  
  #Add PV, MOS, and Quality Ranking
  for (i in 1:Top) {
    #print("P.10 =", P.10[i], "div[i] =", div[i], sep = " ")
    PV <- round(-pv(rate, 10, fv = P.10[i], pmt = div[i]), 2)
    out$'Performance Value'[i] <- PV
    out$'Safety Margin'[i] <- round((PV / out$Price[i] - 1) * 100, 0)
  }
  
  out
  
}

Assign <- function(m, r, Top) {
  for (i in 1:Top) {
    x <- m[i]
    if (x >= .9) {
      m[i] <- r[10]
    } else if (x >= .8) {
      m[i] <- r[9]
    }else if (x >= .7) {
      m[i] <- r[8]
    }else if (x >= .6) {
      m[i] <- r[7]
    }else if (x >= .5) {
      m[i] <- r[6]
    }else if (x >= .4) {
      m[i] <- r[5]
    }else if (x >= .3) {
      m[i] <- r[4]
    }else if (x >= .2) {
      m[i] <- r[3]
    }else if (x >= .1) {
      m[i] <- r[2]
    } else {
      m[i] <- r[1]
    }
  }
  
  m
  
}