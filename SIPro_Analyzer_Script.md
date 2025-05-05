
# SIPro Analyzer R Script

```r
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
  
  #Rename variables (columns)    
  colnames(Data) <- Key[,2]
  Key[,2]
  
  #CALCULATE VARIABLES
  L <- length(Data[,1]) # Number of obs. is used a lot herein

  # Assign percentiles per value factor to get a "relative factor"
  Data[,"BP.Pct"] <- percentile(Data[,"B/P"])
  Data[,"EP.Pct"] <- percentile(Data[,"E/P (Continuing & Diluted)"])
  Data[,"SP.Pct"] <- percentile(Data[,"S/P"])
  Data[,"EBITAEV.Pct"] <- percentile(Data[,"EBITDA/EV"])
  Data[,"CFP.Pct"] <- percentile(Data[,"CFPS/P"])
  Data[,"SHY.Pct"] <- percentile(Data[,"Shareholder Yield"])
  
  # Composite Value Factors (CVF2)
  for (i in 1:L) { 
    r <- as.numeric(Data[i, c("BP.Pct","EP.Pct","SP.Pct",
                              "EBITAEV.Pct","CFP.Pct","SHY.Pct")])
    if (length(r[is.na(r)]) > 2) {
      r[is.na(r)] <- .5
      Data[i, "CVF2"] <- mean(r)
    } else {
      Data[i, "CVF2"] <- mean(r[!is.na(r)])
    }
  }

  # Earnings Quality Composite (EQC)
  Data[,"ATA.Pct"] <- 1 - percentile(Data[,"Accruals to Assets"])
  Data[,"PCNOA.Pct"] <- 1 - percentile(Data[,"Pct. Chg. in NOA"])
  Data[,"ATAA.Pct"] <- 1 - percentile(Data[,"Accruals to Average Assets"])
  Data[,"PCNOA.Pct"] <- 1 - percentile(Data[,"CapEx to Deprec.Amort."])
  
  for (i in 1:L) { 
    r <- as.numeric(Data[i, c("ATA.Pct","PCNOA.Pct","ATAA.Pct","PCNOA.Pct")])
    if (length(r[is.na(r)]) > 1) { 
      r[is.na(r)] <- .5
      Data[i, "EQC"] <- mean(r)
    } else {
      Data[i, "EQC"] <- mean(r[!is.na(r)])
    }
  }

  Data[, "VMQ_Test"] <- Data$CVF2*.8 + Data$EQC*.2

  # Additional Percentiles and Rankings
  Data[,"Coverage.Pct"] <- percentile(Data[,"Times interest earned 12m"])
  Data[,"AE.Pct"] <- percentile(Data[,"A/E"])
  Data[,"ExFin.Pct"] <- 1 - percentile(Data[,"CF.FINANCING"]) 
  Data[,"DebtChg.Pct"] <- 1 - percentile(Data[,"1Y.DEBT.CHG"])
  Data[,"OpIncChg.Pct"] <- percentile(Data[,"1Y.OP.INCOME.GROWTH"]) 
  Data[,"ROE.Pct"] <- percentile(Data[,"Return on equity 12m"])
  Data[,"Turnover.Pct"] <- percentile(Data[,"Asset turnover 12m"])
  Data[,"EY.Pct"] <- percentile(Data[,"EY"]) 
  Data[,"ROC.Pct"] <- percentile(Data[,"ROC"]) 
  Data[,"Magic.Pct"] <- (Data$ROC.Pct + Data$EY.Pct) / 2
  Data[,"RSI.Pct"] <- percentile(Data[,"Relative Strength 26 week"])

  # Top 50 Scenarios
  VMQ.VEN.50 <- GetTop(Data, "CVF2", .05, L, 50)
  VMQ.DEC.50 <- GetTop(Data, "CVF2", .10, L, 50)
  MAGIC.VEN.50 <- GetTop(Data, "Magic.Pct", .05, L, 50)
  
  cat("The following has been saved in:", d, "\n")
  date <- format(Sys.time(), "%m.%d.%Y")
  obj <- c("VMQ.VEN.50", "VMQ.DEC.50", "MAGIC.VEN.50", "date")
  len <- length(obj) - 1

  r <- paste(d, paste(date, "Analysis.R", sep = " "), sep = "/")
  dump(obj, file = r)
  cat("- Analysis.R\n")
  
  suppressMessages(require(XLConnect))
  for (i in 1:len) { 
    r <- paste(d, "/", date, " ", obj[i], ".xlsx", sep = "")
    wb <- loadWorkbook(r, create = TRUE)
    createSheet(wb, date)
    writeWorksheet(wb, get(obj[i]), sheet = date, startRow = 1, startCol = 1)
    saveWorkbook(wb)
    cat("- ", obj[i], ".xlsx\n", sep = "")
  }
}
```

**Additional Functions** are also included in the full script, such as `GetFiles`, `percentile`, `GetTop`, and `Assign`.

---

## Helper Functions

### `GetFiles`

```r
GetFiles <- function(d) {
  f <- list.files(path = d, pattern = "*Analysis.R", recursive = FALSE)
  if (length(f) != 0) {
    r <- cat("Are you sure you wish to overwrite '", f, "'?\n(y/n): ", sep = "")
    r <- readline(prompt = r)
    if (r != "y") stop("Find the correct folder to analyze.")
  }
  f <- list.files(path = d, pattern = "*.TXT", full.names = TRUE, recursive = FALSE)
  f <- f[grep("DATA", f)]
  if (length(f) != 2) {
    stop("Add the two Stock Investor Pro excel data workbooks to the folder before running analysis.")
  } else {
    if (1 == grep("Key", f)) {
      x[2] <- f[1]
      x[1] <- f[2]
      f <- x
    }
    f
  }
}
```

### `percentile`

```r
percentile <- function(v) {
  v.pctrank <- rank(v, na.last = "keep", ties.method = "average") / length(v[!is.na(v)])
  v.pctrank
}
```

### `GetTop`

```r
GetTop <- function(Data, Var, Pct, L, Top) {
  r <- order(Data[,Var], decreasing = TRUE)
  out <- Data[r,]
  if (Var == "Magic.Pct") {
    out <- out[!grepl("financials", out$Sector, TRUE),]
    out <- out[!grepl("utilities", out$Sector, TRUE),]
  } else {
    out <- out[!grepl("counter", out$Exchange, TRUE),]
  }
  r <- round(L * Pct)
  out <- out[1:r,]
  out <- out[order(out[,"Relative Strength 26 week"], decreasing = TRUE),]
  out <- out[1:Top,]
  out[,"Keep"] <- out$ROE.Pct > 0.2 & out$Coverage.Pct > 0.2 & out$EQC > 0.3 &
                  out$Turnover.Pct > 0.2 & out$AE.Pct > 0.1 & out$ExFin.Pct > 0.3 &
                  out$DebtChg.Pct > 0.2 & out$RSI.Pct > 0.3 & out$OpIncChg.Pct > 0.2 &
                  out$Magic.Pct > 0.2 & out$'F Score TTM' > 2
  r <- read.csv("Returns.csv", header = TRUE)
  out[, "rsi"] <- Assign(out$RSI.Pct, r$RSI_Return, Top)
  out[, "eqc"] <- Assign(out$EQC, r$EQC_Return, Top)
  out[, "cov"] <- Assign(out$Coverage.Pct, r$Coverage_Return, Top)
  out[, "ae"] <- Assign(out$AE.Pct, r$AE_Return, Top)
  out[, "exf"] <- Assign(out$ExFin.Pct, r$ExFin_Return, Top)
  out[, "dch"] <- Assign(out$DebtChg.Pct, r$DebtChg_Return, Top)
  out[, "roe"] <- Assign(out$ROE.Pct, r$ROE_Return, Top)
  out[, "tur"] <- Assign(out$Turnover.Pct, r$Turnover_Return, Top)
  out[, "mag"] <- Assign(out$Magic.Pct, r$Magic_Return, Top)
  m <- out[, "rsi"]
  q <- (out[, "eqc"] + out[, "mag"]/4) / (5/4)
  s <- (out[, "cov"] + out[, "ae"] + out[, "exf"] + out[, "dch"]) / 4
  p <- (out[, "roe"] + out[, "tur"]) / 2
  out[, "MQSP Rank"] <- (m + q + s + p) / 4
  require("stringi")
  require("FinCal")
  BV <- out$'Book value/share Q1'
  PE <- out$'7Y.AVG.PE2'
  div <- out$'7Y.AVG.DIV'
  rate <- 0.03
  BV.10 <- BV * (1 + out$G/100)^10
  EPS.10 <- out$AVG.ROE/100 * BV.10
  P.10 <- PE * EPS.10
  for (i in 1:Top) {
    PV <- round(-pv(rate, 10, fv = P.10[i], pmt = div[i]), 2)
    out$'Performance Value'[i] <- PV
    out$'Safety Margin'[i] <- round((PV / out$Price[i] - 1) * 100, 0)
  }
  out
}
```

### `Assign`

```r
Assign <- function(m, r, Top) {
  for (i in 1:Top) {
    x <- m[i]
    m[i] <- r[ceiling(x * 10)]
  }
  m
}
```
