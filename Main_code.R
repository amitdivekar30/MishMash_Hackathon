
#importing dataset
library(readxl)
Train_original <- read_excel("Training-Data-Sets.xlsx")
Test_original <- read_excel("Test dataset v1.xlsx")
View(Train_original)

summary(Train_original)
boxplot(Test_original$EQ)
columns<-c(colnames(Train_original))
mean(Train_original$EQ)
mean(Test_original$EQ)

#checking missing data
is.na(Train_original)
sum(is.na(Train_original))

##########################################################################
#################Converting Traing set into period form####################
##########################################################################

#since Training set is daywise and Test set is periodwise, so converting Traing set by period
#1 perod= 4 weeks=28days

a_1<-sapply(Train_original[1:28,c(2:39)], mean)

for (i in seq(1,11984,28)){
  a<-sapply(Train_original[i:(i+27),c(2:39)], mean)
  a<-as.data.frame(rbind(a_1[],a))
  a_1<-a
}

Train<-a[-1,]

library(writexl)
write_xlsx(Train,"Train.xlsx")


##########################################################################
##########################Selecting Drivers###########################
##########################################################################
Train<- read_excel("Train.xlsx") #Train is Training data after time series transformation
View(Train)
summary(Train)
attach(Train)
boxplot(EQ)
normalize <- function(x) {
  return ((x - min(x)) / (max(x) - min(x)))
}

#Normalized all the columns except EQ
Train1<-normalize(Train[,2:38])
Train1 <- cbind(Train1,Train$EQ)
str(Train1)
colnames(Train1)[38] <- "EQ"

#Using Boruta for Driver selection
#install.packages("Boruta")
library(Boruta)
boruta_output <- Boruta(EQ ~ ., data=Train1, doTrace=0) 
names(boruta_output)

# Get significant variables including tentatives
boruta_signif <- getSelectedAttributes(boruta_output, withTentative = TRUE)
print(boruta_signif)  
# Get significant variables without tentatives
boruta_signif_wo <- getSelectedAttributes(boruta_output, withTentative = FALSE)
print(boruta_signif_wo)

# Plot variable importance
plot(boruta_output, cex.axis=.7, las=2, xlab="", main="Variable Importance")  

# Using random forest for variable selection
#install.packages('randomForest')
library('randomForest')
rfModel <-randomForest(EQ ~ ., data = Train1)

# Getting the list of important variables
d<-importance(rfModel)

Train_driver<-Train[c("EQ","Social_Search_Impressions", "Median_Rainfall","Inflation",
                      "pct_PromoMarketDollars_Category","EQ_Category","EQ_Subcategory",
                      "pct_PromoMarketDollars_Subcategory")]

Test_driver<-Test_original[c("EQ","Social_Search_Impressions", "Median_Rainfall","Inflation",
                             "pct_PromoMarketDollars_Category","EQ_Category","EQ_Subcategory",
                             "pct_PromoMarketDollars_Subcategory")]

Train_driver["Period"]<-1:428
Test_driver["Period"]<-429:467
Dataset<-rbind(Train_driver,Test_driver)

#writing dataset with selected driver
library(writexl)
getwd()
write_xlsx(Train_driver,"Train_driver.xlsx")
write_xlsx(Test_driver,"Test_driver.xlsx")
write_xlsx(Dataset,"Dataset.xlsx")



#########################################################################################################
#################Forecasting using Multivariate forecasting####################
########################################################################################################
Train_driver<- read_excel("Train_driver.xlsx")
Test_driver<- read_excel("Test_driver.xlsx")
Dataset<-read_excel("Dataset.xlsx")

#install.packages("vars")
library(forecast)
library(fpp)
library(smooth)
library(readxl)
library(vars)
library(tseries)

windows()
plot(Train_driver$EQ,type="o")
plot.ts(Train_driver)
plot(ts(Dataset$EQ,frequency = 13))

# Converting data into time series object
plot(ts(Train_driver$EQ,frequency = 13))
windows()
plot(ts(Train_driver[-9],frequency = 13))

# converting time series object
data<-ts(Dataset[-9],frequency = 13)

#lag selection
info.bv <- VARselect(data, lag.max = 14, type = "both")
info.bv$selection

model<- VAR(data[1:454,],p=5, type="both", ic="AIC") 
model
k<-predict(model, n.ahead=13)
model_mape<-MAPE(k$fcst$EQ[1:13],data[455:467, 1])*100     
model_mape

#In order to consider seasonality in Test Data creating model on Test Dataset
data1<-ts(Test_driver[-9],frequency = 13)

info.bv <- VARselect(data1[1:26,], lag.max = 13, type = "both")
info.bv$selection

model1<- VAR(data1[1:26, ], p=2, type="both", ic="AIC") 
model1
k1<-predict(model1, n.ahead=13)
model_mape<-MAPE(k1$fcst$EQ[1:13],data1[27:39, 1])*100     
model_mape #43.64

#Final Model considering complete Test dataset
info.bv <- VARselect(data1, lag.max = 13, type = "both")
info.bv$selection
final_model<- VAR(data1, p=3, type="both", ic="AIC") 
final_model
k2<-predict(final_model, n.ahead=6)

z<-k2$fcst$EQ
f<-Test_driver[1]
g<-z[1]
colnames(g)[1]<-"EQ"
a3<-rbind(f,g)
plot(ts(a3,frequency = 13, start = c(2016)))
plot(ts(a3[1:39, ],frequency = 13, start = c(2016)))
plot(ts(a3[40:45, ],frequency = 13, start = c(2019)))

#install.packages('coefplot')
library(coefplot)
coefplot(final_model$varresult$EQ)
names(final_model$varresult)

#forecast for next 6 periods
sales_forecast<-predict(final_model, n.ahead=6)$fcst$EQ
sales_forecast
