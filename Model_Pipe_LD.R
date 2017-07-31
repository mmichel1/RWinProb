#   Data Sciene Learning Project Team GER (2017)
#   Win Probability Prediction at Rainbow Pursuit (internal scope)
#
#   Feature Selection with Random Forest
#
#   Author: Hans Karlshoefer
#
#   Version 1.1   HK  initial version
#

# ====================== Parameters ==============
# reads from file in rel_fpath 
rel_fpath   <- "./data/clean_4_modeling.xlsx"
# writes plots and clean data frame to outputfile
outputfile <- "./output/data_exploration.pdata"

Read_Fresh_from_File <- T # Read raw data file with 3.000 lines, needs to be 
# set to TRUE for first run to fill var rawframe
# clean dat frame, set to FALSE for initial data understanding
options(scipen=999)       # suppress exponential form of print output format
printpdata <- FALSE         # if printout in Pdata desired
scatterplot <- FALSE      # it T do scatterplot
lost_dropped <- TRUE      # combine lost and dropped status as lost/dropped

# ===================== Function Lift Chart ======
library(dplyr)
library(ggplot2)
library(stringr)
library(colorRamps)
library(plotly)
library(ROCR)      # plot ROC curve
library(caret)     # machine learning suites

library(readxl)    # moved from load data section 

Lift_Chart <- function (yhat_proba, ytrue, headline, color) {
  # Creating the data frame
  data_lift <- data.frame(yhat_proba, ytrue)
  
  # Ordering the dataset
  data_lift <- data_lift[order(data_lift$yhat_proba, decreasing = TRUE),]
  
  # Creating the cumulative density
  data_lift$cumden <- cumsum(data_lift$ytrue/sum(data_lift$ytrue))
  
  # Creating the % of population
  data_lift$perpop <- (seq(nrow(data_lift))/nrow(data_lift))*100
  
  # Ploting
  plot(data_lift$perpop,data_lift$cumden,type="l",
       xlab="% of Cases",ylab="% of Default's", 
       col=color, main=headline, cex.main=0.75)
  
  y = function(yhat_proba) { return ( yhat_proba/100 ) }
  plot(y, 0, 100, add = TRUE) 
}

# ===================== Function ROC Chart =======
ROC_Chart <- function (yhat_proba, ytrue, headline) {
  pred_obj <- prediction(yhat_proba, ytrue)
  perf     <- performance(pred_obj, 'tpr', 'fpr')
  
  #calcualte AUC (Area under Curve-ROC)
  auc_value <- performance(pred_obj, 'auc')@y.values[[1]]
  print (auc_value)
  plot(perf, colorize=T,
       xlab="1 - Specificity", ylab="Sensitivity", cex.main=0.75,
       main = toString (c(headline, " AUC", round(auc_value, digits = 4))))
}

#-------------  Shiny Functions ----------------------
loadData <- function()
{
  # R e a d   d a t a   as  numeric
  if (Read_Fresh_from_File)  {
    data <<- as.data.frame (read_excel(path = rel_fpath, sheet = "dummies",
                                    col_types = c("numeric", "date", 
                                                  "date", "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric", 
                                                  "numeric", "numeric", "numeric")))
  }
}

plotDensity <- function(colname)
{
#  browser() 
# ===========================
# 4.2 Bi-variate analysis / reduction
# ===========================
# check category relevancy (0-1) w.r.t. output var. "FinStatus" via barplots
fn_norm_max  = function(y) { return (y/max(y)) }
y <- array(1:6, dim=c(2,3))

y[1,3] <- sum(data$FinWon_Status == 1) # Won
y[2,3] <- sum(data$FinWon_Status == 0) # Lost/Dropped

#for (i in 13:ncol(data)) {
  y[1,1] <- sum(data[colname] == 0 & data$FinWon_Status == 1)
  y[2,1] <- sum(data[colname] == 0 & data$FinWon_Status == 0)
  y[1,2] <- sum(data[colname] == 1 & data$FinWon_Status == 1)
  y[2,2] <- sum(data[colname] == 1 & data$FinWon_Status == 0)
  
  yy <- as.data.frame(y)
  yy <- as.data.frame(lapply(yy, fn_norm_max))
#  browser()
  summary (yy)
  
  par(mfrow=c(1,1))
  bplt <- barplot(as.matrix(yy), main=names(data[colname]), names.arg=c("0","1","W/LD"),
                  legend.text=c("Won", "Lost/Dropped"), beside=TRUE,
                  #legend.text=c("Won", "Lost", "Dropped"), beside=TRUE,
                  col=c("white", "darkgray"))
                  #col=c("white", "darkgray", "gray"))
  text(x= bplt, as.matrix(yy)+0.02, labels=as.character(y), xpd=TRUE)
#}

#data_value <- subset(copy(data),select=c('FinWon_Status',colname));
#names(data_value) <- c('Won_LR.Status','VALUE');
#data_value <- na.omit(data_value);

#par(mfrow=c(2,1));
#hist(data_value$VALUE,xlab=paste("Value:",colname),main = 'Histogram',xlim=c(min(data_value$VALUE),max(data_value$VALUE)));
#sm.density.compare(as.numeric(data_value$VALUE),group = data_value$Won_LR.Status, xlab=paste("Value:",colname),xlim=c(min(data_value$VALUE),max(data_value$VALUE)));
#title("Density compare");
#legend("topright",levels(as.factor(data_value$Won_LR.Status)), fill=c(2:(2+length(levels(as.factor(data_value$Won_LR.Status))))),cex=1);
}

plotPCA <- function()
{
  #browser()
  drops <- c( 
             "Time_bef_Purs",    # correlation with ChangeRec 0.75
             "FinDelStart",           # output variable
             "FinOE",                 # output variable
             "DelStart",              # variable relevant with Delivery Start as output var.
             "FinLost_Status",         # colinear to FinWon_Status
             "FinWon_Status")         # output variable
  data_value  <- data[ , !(names(data) %in% drops)]    

# =====================================
# PCA in R
# =====================================
data_pca <- prcomp (data_value, center=TRUE, scale = TRUE)
#print(summary (data_pca))
#print(data_pca$rotation)
plot (cumsum((data_pca$sdev)^2/sum((data_pca$sdev)^2)),
      xlab="PCA components", ylab="Cumulative Proportion of variance")
return (data_pca)
}




#---------------- pseudo main function -----------------
Main <- function () {
  
browser()
# R e a d   d a t a   as factors
data_fac <- as.data.frame (read_excel(path = rel_fpath, sheet = "factors",
            col_types = c("text", 
            "numeric", "date", "text", "text", 
            "date", "text", "text", "text", "text", 
            "text", "text", "text", "text", "numeric", 
            "numeric", "numeric", "text", "text", 
            "text", "text", "numeric", "numeric", 
            "numeric", "text", "numeric")))

# combine lost and dropped status
if (lost_dropped == TRUE) data_fac$FinStatus[data_fac$FinStatus != "Won"] <- "LD"

# factorize non numeric variables
for (i in 1:ncol(data_fac)) {
  if (is.character(data_fac[,i])) 
    data_fac[,i] <- as.factor(unlist(data_fac[,i]))
   }


# ========================================================================= 
# 4. Dimension Reduction and Feature Selection
# =========================================================================
drops <- NULL      # contains varialbes which will be dropped for modelling

# ===========================
# 4.1 Univariate analysis / reduction
# ===========================
# eliminate records with OE < 50kEUR
#data     <-data[data$OE >= log(50),] 
#data_fac <-data_fac[data_fac$OE >= log(50),] 


# ===========================
# 4.3	Multi variate analysis / reduction
# ===========================
# reduce higly correlated, double and constant varialves
drops <- rbind(drops, "Time_bef_Purs")    # correlation with ChangeRec 0.75,
                                          # see chapter 3.3.2 in Data Exploration

# ===========================
# 4.4 Dimension Reduction
# ===========================
# Inclusion/Exclusion of variales

# Additional drops w.r.t. "one_line_file" by exploration phase
# "FixPrice.BillType"         nearly opposite of "Time & Material"
#                                                -> other BillType
# "CloudificLegacy.Offer"     less "1" frequency -> other Offerings
# "MultSuppInt.Offer"         less "1" frequency -> other Offerings
# "DigWorkpl.Offer"           less "1" frequency -> other Offerings
# "IDM.Offer"                 less "1" frequency -> other Offerings

drops <- c(drops, 
           "FinDelStart",           # output variable
           #"FinStatus",            # output variable
           "FinOE",                 # output variable
           "DelStart",              # variable relevant with Delivery Start as output var.
           "FinLost_Status")         # colinear to FinWon_Status

drops_data <-c(
           "FinOE",           # output variable
           "FinDelStart",     # output variable
           "DelStart",        # relevant if Delivery Start is output var
          #"Probability",    
           "ProjMargin",      ##
          #"OE",
           "ChangeRec",
           "Teamfill",        ##
           "Time_bef_Purs",   ## correlation with ChangeRec 0.75,
           "TCV_SD_bef_Purs",
           "FinLost_Status",
          #"FinWon_Status",   # output variable
          #"AM_ServType",
           "Proj_ServType",
           "Prod_ServType",   ##
           "TopAccont",
           "CloudAppBuildDeply_Offer",
           "TelcoSol_Offer",
           "SecCompl_Offer",
           "Codex_Offer",
           "AppStratTransf_Offer",
           "AppMgmtServRUN_Offer",
           "ERPSol_Offer",
           "I4_0_IOT_Offer",
           "SAP_HANA_Offer",
           "WorldGridSol_Offer",
           "CX_Mgmt_Offer",
          #"TM_BillType",
           "Campaign",
          #"NewExit_OppType",
          #"FertExist_OppType",
          #"RenewExist_OppType",
          #"Siemens_Market",
          #"TMU_Market",
           "PH_Market",
           "MRT_Market",
           "BestCanDo_Prio",
           "NotAssign_Prio",
           "Medium_Prio",
           "High_Prio",         ##
          #"FrameAgreement",
           "SW_Platform",
           "BidMgrInPurs",      ##
           "Frequ_Sales",
           "SolMgrInPurs",      ##
           "DelMgrInPurs",
           "less2perc_BidBudPerc",
           "more2perc_BidBudPerc"
           )
drops_data_fac <-c(
          #"FinStatus",         # output variable
           "FinOE",             # output variable
           "FinDelStart",       # output variable
           "ServType",
           "TopAccont",
           "DelStart",          # relevant if Delivery Start is output var
           "Offer",
          #"BillType",
           "Campaign",
          #"OppType",
          #"Market",
           "Prio",
          #"FrameAgreement",
           "SW_Platform",
          #"Probability",
           "ProjMargin",
          #"OE",
           "BidMgrInPurs",
           "Frequ_Sales",
           "SolMgrInPurs",
           "DelMgrInPurs",
           "ChangeRec",
           "Teamfill",
           "Time_bef_Purs",    # correlation with ChangeRec 0.75,
           "BidBudPerc",
           "TCV_SD_bef_Purs"
           )

data     <- data    [ , !(names(data) %in% drops)]              # use all var.
data_fac <- data_fac[ , !(names(data_fac) %in% drops)]

#data     <- data    [ , !(names(data) %in% drops_data)]            # use relevenat var.
#data_fac <- data_fac[ , !(names(data_fac) %in% drops_data_fac)]

# ===========================
# 4.5 Binning
# ===========================
# 
if ("OE" %in% names(data_fac)) {
   data_fac$OE <- cut(data_fac$OE, breaks=c(-Inf,log(50),5,6,7,8,10,Inf)) # first bin up to 50kEUR
   plot (data_fac$OE, main="OE") }
if ("Probability" %in% names(data_fac)) {
   data_fac$Probability <- cut(data_fac$Probability, breaks=c(-Inf,0,31,61,100))
   plot (data_fac$Probability, main="Probability") }
if ("ProjMargin" %in% names(data_fac)) {
   data_fac$ProjMargin <- cut(data_fac$ProjMargin, breaks=c(-Inf,20,25,30,100))
   plot (data_fac$ProjMargin, main="ProjMargin") }
if ("PChangeRec" %in% names(data_fac)) {
   data_fac$ChangeRec <- cut(data_fac$ChangeRec, breaks=c(-Inf,5,Inf))
   plot (data_fac$ChangeRec, main="ChangeRec") }
if ("TCV_SD_bef_Purs" %in% names(data_fac)) {
   data_fac$TCV_SD_bef_Purs <- cut(data_fac$TCV_SD_bef_Purs, breaks=c(-Inf,0.1,3))
   plot (data_fac$TCV_SD_bef_Purs, main="CV_SD_bef_Purs") }

# =====================================
# 5.  M o d e l l i n g   and   E v a l u a t i o n
# =====================================

# Training set contains 60% of total rows.
# Select them randomly using sample function
# train_index will contain 60% of indexes
#set.seed(123)
train_index <- sample(1:nrow(data), nrow(data) * 0.6)

# ===========================
# Rescale predictors to ]0;1[ intervall required vor PCA, Neural Net, ...
# ===========================
# if you need to rescale / normalize all predictor so that all are in same range.
# fn_norm function will return rescalled values such that min=0 and max=1
fn_norm  = function(x) { return ( (x-min(x)) / (max(x) - min(x)) ) }

data_fn = as.data.frame(lapply(data[, 1:ncol(data)], fn_norm))
summary(data_fn)


# =====================================
# Correspondence analysis (CA)
# =====================================
#CA(data_fac, ncp = 5, graph = TRUE)
#res.ca <- CA(data, graph = FALSE)
#print(res.ca)
#summary(res.ca, nb.dec = 2, ncp = 5)


# =====================================
# 5. Variable Selction by Random Forest
# =====================================
library(VSURF)

data_train <- data[train_index,] 
data_test  <- data[-train_index,]

# Step I              # function VSURF is a wrapper for the three
                      # intermediate functions VSURF_thres, VSURF_interp
                      # and VSURF_pred
data_thres <-VSURF_thres(x = data_train[,which(names(data_train)!="FinWon_Status")],
                       y = as.factor(data_train$FinWon_Status), parallel=TRUE)
summary (data_thres)
plot(data_thres)
print(names(data[,data_thres$varselect.thres]))
plot(data_thres, step = "thres", imp.mean = FALSE, ylim = c(0, 0.003))

# Step II
data_interp <- VSURF_interp(x = data_train[,which(names(data)!="FinWon_Status")],
                          y=as.factor(data_train$FinWon_Status),
             vars = data_thres$varselect.thres,       
             parallel=TRUE)
print(names(data[,data_interp$varselect.interp]))

#Step III
data_pred <- VSURF_pred(x = data_train[,which(names(data)!="FinWon_Status")],
                      y = as.factor(data_train$FinWon_Status),
             err.interp = data_interp$err.interp,
             varselect.interp = data_interp$varselect.interp)
print(names(data[,data_pred$varselect.pred]))

# Step I - III wrapped
#data_vsurf <- VSURF((x = data_train[,which(names(data)!="FinWon_Status")],
#                   y = as.factor(data_train$FinWon_Status), parallel=TRUE)

# =====================================
# fit random forest with auto selected variables
# =====================================
library(randomForest)

par(mfrow=c(1,1))

y <- as.factor(data$FinWon_Status)
data_rf <- data[,data_pred$varselect.pred]

data_train <- data_rf[train_index,]
y_train    <- y[train_index]
data_test  <- data_rf[-train_index,]
y_test     <- y[-train_index]

#data_train$FinWon_Status <- as.factor(data_train$FinWon_Status)

fit.rf = randomForest(x = data_train, y = y_train,
                    #  data=data_train,
                      importance=TRUE,
                      proximity=TRUE)
print(fit.rf)
importance(fit.rf)    # MeanDecreaseAccuracy: shows how worse the model
#                   performs without each variable
# MeanDecreaseGini: measures how pure the nodes are
#                   at the end of the tree if
#                   variable is taken out 
varImpPlot(fit.rf)
plot(fit.rf)

summary(fit.rf)
# Evaluation with train data set
#data_test$FinWon_Status <- as.factor(data_test$FinWon_Status)
yhat <- predict(fit.rf, newdata = data_train) 
caret::confusionMatrix(yhat, y_train)
                       
# Evaluation with test data set
#data_test$FinWon_Status <- as.factor(data_test$FinWon_Status)
yhat <- predict(fit.rf, newdata = data_test) 
caret::confusionMatrix(yhat, y_test)

}