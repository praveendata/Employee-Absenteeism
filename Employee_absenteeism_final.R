#===================================================================================
#                    Employee_Absenteeism_Project
#===================================================================================


# Clearing R environment
rm(list = ls())

# Setting working directory 
setwd("C:/Users/Garima/Downloads/Edwisor")

# Checking, whether path is set or not 
getwd()

# Loading required libraries
required_lib = c("rJava","xlsx","ggplot2","gridExtra","corrplot","tree","caret","randomForest","usdm")
install.packages(required_lib)
lapply(required_lib,require,character.only = TRUE)

# Loading dataset 
absenteeism_data = read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex = 1,header = T)
View(absenteeism_data)

# Checking the structure of the data 
str(absenteeism_data)
## All the variables are of numerical datatypes

## Checking summary of the data
summary(absenteeism_data)

## Let us separate categorical data and numerical data 
## on analysing the data and as mentioned in the problem statement we observe following feature set as categorical : reason for absence ,month of absence,day of the week
## season ,disciplinary failure ,education,social drinker,social smoker.

# Separating categorical variable and numerical variable
numerical_set = c("ID","Transportation.expense","Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.","Hit.target","Son","Pet","Height","Weight","Body.mass.index","Absenteeism.time.in.hours")
categorical_set = c("Reason.for.absence","Month.of.absence","Day.of.the.week","Seasons","Disciplinary.failure","Education","Social.drinker","Social.smoker")

# Converting datatype of categorical data into factor levels 
for(i in categorical_set){
  absenteeism_data[,i] = as.factor(absenteeism_data[,i])
}

#==================== Missing Value Analysis ====================
#================================================================

## We evaluate the percentage of missing value in each column .
## On viewing the dataset ,we see certain entries have 0 value which is not an acceptable entry logically .Hence we replace them as NA 
for(i in c("Reason.for.absence","Month.of.absence","Day.of.the.week","Seasons","Education","ID","Age","Weight","Height","Body.mass.index")){
  absenteeism_data[i][(absenteeism_data[i] == 0)] = NA
}

# Calculating total sum of missing values column wise
missing_val_column_wise = data.frame(apply(absenteeism_data,2,function(x){sum(is.na(x))}))

names(missing_val_column_wise)[1] = "NA_Sum"

missing_val_column_wise$NA_Percent = (missing_val_column_wise$NA_Sum/nrow(absenteeism_data))*100

# If any column has more than 30% of missing value ,we will drop that column and update numerical and categorical set accordingly
for(i in 1:length(absenteeism_data)){
  if(missing_val_column_wise$NA_Percent[i]>=30){
    absenteeism_data = absenteeism_data[,-i]
    if(any(numerical_set %in% colnames(absenteeism_data[i]))){
      numerical_set = numerical_set[-(grep(colnames(absenteeism_data[i],numerical_set)))]
    }
    else if(any(categorical_set %in% colnames(absenteeism_data[i]))){
      categorical_set = categorical_set[-(grep(colnames(absenteeism_data[i],categorical_set)))]
    }
  }
}

## none of the columns has more than 30% missing value.Thus all the columns are retained for now.


## We will check which imputation method suits best in our data set model
## To check accuracy of our method ,let us impute NA values in row 10 and check with our imputated values with actual values
##data_check = absenteeism_data[11,]
#absenteeism_data[11,] = NA

## Let us impute missing values 
## Mean method for numerical data 
## Mode method for categorical data 

getmode = function(x){
  unique_x = unique(x)
  mode_val = which.max(tabulate(match(x,unique(x))))
}

## Mean method 
#Function : impute_mean_mode
#Parameter : dataset 

impute_mean_mode = function(data_set){
  for(i in colnames(data_set)){
    if(sum(is.na(data_set[,i]))!=0){
      if(is.numeric(data_set[,i])){
        
        data_set[is.na(data_set[,i]),i] = round(mean(data_set[,i],na.rm = TRUE))
      }
      else if(is.factor(data_set[,i])){
        
        data_set[is.na(data_set[,i]),i] = getmode(data_set[,i])
      }
      
    }
  }
  return(data_set)
}

## Median Method for numerical data
#Function : impute_mode_mode
#Parameter : dataset 

impute_median_mode = function(data_set){
  for(i in colnames(data_set)){
    if(sum(is.na(data_set[,i]))!=0){
      if(is.numeric(data_set[,i])){
        
        data_set[is.na(data_set[,i]),i] = median(data_set[,i],na.rm = TRUE)
      }
      
      
      else if(is.factor(data_set[,i])){
        
        data_set[is.na(data_set[,i]),i] = getmode(data_set[,i])
      }
      
    }
  }
  #print(data_set)
  return(data_set)
}

## KNN method
#Function : impute_knn
#Parameter : dataset

impute_knn = function(data_set){
  install.packages("DMwR")
  library(DMwR)
  if(sum(is.na(data_set))!=0){
    print("Imputing missing value in the data")
    ##Knn works on numeric data set 
    for(i in categorical_set){
      data_set[,i] = as.numeric(data_set[,i])
    }
    data_set = knnImputation(data_set,k=5)
    data_set = round(data_set)
  }
  return(data_set)
}

# Selecting the best method

#absenteeism_data = impute_mean_mode(absenteeism_data)
absenteeism_data = impute_median_mode(absenteeism_data)
##absenteeism_data = impute_knn(absenteeism_data)

################################ Data visualisation #########################################
#============================================================================================

# Numerical data

hist(absenteeism_data$ID,col = "blue",main = "ID frequency",xlab = "ID",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Transportation.expense,col = "blue",main = "Transportation expense frequency",xlab = "Transportation expense",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Distance.from.Residence.to.Work,col = "blue",main = "Dist from Residence to Work frequency",xlab = "Distance from Residence to Work",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Service.time,col = "blue",main = "Service time frequency",xlab = "Service time",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Age,col = "blue",main = "Age frequency",xlab = "Age",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Work.load.Average.day.,col = "blue",main = "Workload Average day frequency",xlab = "Workload Average day",breaks = 50,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Hit.target,col = "blue",main = "Hit target frequency",xlab = "Hit target",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Son,col = "blue",main = "Son frequency",xlab = "Son",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Pet,col = "blue",main = "Pet frequency",xlab = "Pet",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Height,col = "blue",main = "Height frequency",xlab = "Height",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Weight,col = "blue",main = "Weight frequency",xlab = "Weight",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Body.mass.index,col = "blue",main = "Body Mass Index frequency",xlab = "Body Mass Index",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(absenteeism_data$Absenteeism.time.in.hours,col = "blue",main = "Absent hours frequency",xlab = "Absent hours",breaks = 30,cex.lab = 1,cex.axis = 1)
dev.off()

# Categoral data

barplot(table(absenteeism_data$Reason.for.absence),col = "blue",main = "Reason for absence frequency",xlab = "Reason for absence",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Month.of.absence),col = "blue",main = "Month of absence frequency",xlab = "Month of absence expense",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Day.of.the.week),col = "blue",main = "Day.of.the.week frequency",xlab = "Day of the week",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Seasons),col = "blue",main = "Seasons frequency",xlab = "Seasons time",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Disciplinary.failure),col = "blue",main = "Disciplinary failure frequency",xlab = "Disciplinary failure",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Education),col = "blue",main = "Education frequency",xlab = "Education",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Social.drinker),col = "blue",main = "Social.drinker frequency",xlab = "Social drinker",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(absenteeism_data$Social.smoker),col = "blue",main = "Social.smoker frequency",xlab = "Social smoker",ylab = "Frequency",cex.lab = 1,cex.axis = 1)


# Boxplot of categorical data set vs target variable
dev.off()
for(i in 1:length(categorical_set)){
  assign(paste0("gg",i),ggplot(aes_string(y=absenteeism_data$Absenteeism.time.in.hours,x=absenteeism_data[,categorical_set[i]]),data = subset(absenteeism_data))
         + stat_boxplot(geom = "errorbar",width = 0.3) +
           geom_boxplot(outlier.colour = "red",fill = "blue",outlier.shape = 18,outlier.size = 1) +
           labs(y = "Absenteeism.time.in.hours",x=names(absenteeism_data[categorical_set[i]])) + 
           ggtitle(names(absenteeism_data[categorical_set[i]])))
  
}
gridExtra::grid.arrange(gg1,gg2,nrow = 2,ncol=1)
gridExtra::grid.arrange(gg3,gg4,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg5,gg6,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg7,gg8,nrow = 2,ncol = 1)


# Scatter plot of numerical data set vs the target variable

plot(absenteeism_data$ID,absenteeism_data$Absenteeism.time.in.hours,xlab = "ID",ylab = "Absenteeism time in hours",main = "Absenteeism time vs ID",col="blue")
plot(absenteeism_data$Transportation.expense,absenteeism_data$Absenteeism.time.in.hours,xlab = "Transportation expense",ylab = "Absenteeism time in hours",main = "Absent time vs Transp expense",col="blue")
plot(absenteeism_data$Distance.from.Residence.to.Work,absenteeism_data$Absenteeism.time.in.hours,xlab = "Distance.from.Residence.to.Work",ylab = "Absenteeism time in hours",main = "Absent time vs Travel distance",col="blue")
plot(absenteeism_data$Service.time,absenteeism_data$Absenteeism.time.in.hours,xlab = "Service Time",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Service time",col="blue")
plot(absenteeism_data$Age,absenteeism_data$Absenteeism.time.in.hours,xlab = "Age",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Age",col="blue")
plot(absenteeism_data$Work.load.Average.day.,absenteeism_data$Absenteeism.time.in.hours,xlab = "Work.load.Average.day",ylab = "Absenteeism time in hours",main = "Absent time vs Work.load.Avg.day",col="blue")
plot(absenteeism_data$Hit.target,absenteeism_data$Absenteeism.time.in.hours,xlab = "Hit target",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Hit target",col="blue")
plot(absenteeism_data$Son,absenteeism_data$Absenteeism.time.in.hours,xlab = "Son",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Son",col="blue")
plot(absenteeism_data$Pet,absenteeism_data$Absenteeism.time.in.hours,xlab = "Pet",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Pet",col="blue")
plot(absenteeism_data$Height,absenteeism_data$Absenteeism.time.in.hours,xlab = "Height",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Height",col="blue")
plot(absenteeism_data$Weight,absenteeism_data$Absenteeism.time.in.hours,xlab = "Weight",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Weight",col="blue")
plot(absenteeism_data$Body.mass.index,absenteeism_data$Absenteeism.time.in.hours,xlab = "BMI",ylab = "Absenteeism time in hours",main = "Absenteeism time vs BMI",col="blue")


############################Outlier Analysis ########################################
#====================================================================================

# Replacing outliers in numerical dataset with NAs using boxplot method

for(i in numerical_set){
  outlier_value = boxplot.stats(absenteeism_data[,i])$out
  print(names(absenteeism_data[i]))
  print(outlier_value)
  absenteeism_data[which(absenteeism_data[,i] %in% outlier_value),i] = NA
}

# Imputing missing values with the best method 
absenteeism_data = impute_median_mode(absenteeism_data)

##############################Feature Selection #############################################################
#============================================================================================================

## Correlation analysis 

dev.off()
pairs(absenteeism_data[numerical_set[c(1:6,13)]])
pairs(absenteeism_data[numerical_set[7:13]])

correlation_matrix = cor(absenteeism_data[,numerical_set])
dev.off()
corrplot(correlation_matrix,method = "number",type = 'lower')

## "Weight" and "BMI" has highcorrelation value therefore dropping the weight varibale
## Using VIF for multicollinearity check 


# Checking multi-collinearity
vif(absenteeism_data[,numerical_set])

##If VIF>5 thus we can remove those features
# view(absenteeism_data)

# ANOVA test for categorical data 

for(i in categorical_set){
  print(i)
  aov_summary = summary(aov(Absenteeism.time.in.hours~absenteeism_data[,i],data = absenteeism_data))
  print(aov_summary)
}


# Dimensionality reduction
absenteeism_data = subset(absenteeism_data, select = -(which(names(absenteeism_data) %in% c("Weight","Day.of.the.week","Seasons","Education","Social.smoker","Social.drinker"))))

## observing p value of the table ,we conclude that we should drop features whose p value is greater than 0.05


##################### Sampling ######################################################################
#====================================================================================================

# Separating the data set into test and train 
# Separate dataset into test and train sets

set.seed(2810)
#Stratified sampling 

# Using createdataPartition for sampling because we want to include sample of all the levels in the feature reason for absence in our train model
#We chose 80% of the data in the train set and 20% in the test set

train = createDataPartition(absenteeism_data$Reason.for.absence,times = 1,p = 0.80,list = F)
test = -(train)


############################### Model Development  ####################################################
#======================================================================================================

## Linear Regression
#=====================
modelLR = lm(Absenteeism.time.in.hours~.,data = absenteeism_data[train,])
summary(modelLR)

par(mar = c(2,2,2,2))

plot(modelLR)

par(mfrow = c(3,1))

predictLR = predict(modelLR,absenteeism_data[test,])

plot(absenteeism_data$Absenteeism.time.in.hours[test])
lines(predictLR,col='red')

sqrt(mean((predictLR-absenteeism_data$Absenteeism.time.in.hours[test])^2))

## RMSE : 2.94


## Decision Trees
#=================
modelDT = tree(Absenteeism.time.in.hours~.,absenteeism_data,subset = train)
summary(modelDT)
plot(modelDT)
text(modelDT,pretty = 0)
predictDT = predict(modelDT,newdata = absenteeism_data[test,])
plot(absenteeism_data$Absenteeism.time.in.hours[test])
lines(predictDT,col='green')
sqrt(mean((predictDT-absenteeism_data$Absenteeism.time.in.hours[test])^2))

##RMSE 3.06


## Random Forest
#=================
modelRF = randomForest(Absenteeism.time.in.hours~.,data = absenteeism_data,subset = test,mtry = 10,ntree=10,importance = TRUE)
varImpPlot(modelRF)
importance(modelRF)
predictRF = predict(modelRF,newdata = absenteeism_data[test,])
plot(absenteeism_data$Absenteeism.time.in.hours[test])
lines(predictRF,col='blue')
sqrt(mean((predictRF-absenteeism_data$Absenteeism.time.in.hours[test])^2))

#mtry ntree rmse 
#5      5  1.57
#5     10  1.65
#10     5  1.59
#10    10  1.45

#######################################Model Inference #########################################
#===============================================================================================

## We observe that Reason for absence is the most important predictor in absenteeism hours
## Let us see the sum and mean of the absenteeism hours Reason wise

reason_sum_hrs = aggregate(absenteeism_data$Absenteeism.time.in.hours,by = list(Category = absenteeism_data$Reason.for.absence),FUN = sum)
names(reason_sum_hrs)=c("Reason no.","Sum of absent hours")
reason_mean_hrs = aggregate(absenteeism_data$Absenteeism.time.in.hours,by = list(Category = absenteeism_data$Reason.for.absence),FUN = mean)
names(reason_mean_hrs)=c("Reason no.","Mean of absent hours")
table(absenteeism_data$Reason.for.absence)

################################### Monthly loss for the Company#################################################
#===============================================================================================================

loss_data = absenteeism_data[,c("Month.of.absence","Work.load.Average.day.","Service.time","Absenteeism.time.in.hours")]
str(loss_data)

loss_data$WorkLoss = round((loss_data$Work.load.Average.day./loss_data$Service.time)*loss_data$Absenteeism.time.in.hours)
View(loss_data)

monthly_loss = aggregate(loss_data$WorkLoss,by = list(Category = loss_data$Month.of.absence),FUN = sum)
names(monthly_loss) = c("Month","WorkLoss")

ggplot(monthly_loss,aes(monthly_loss$Month,monthly_loss$WorkLoss))+geom_bar(stat = "identity",fill = "blue")+labs(y="WorkLoss",x="Months")



# End Of The Project.