### Innovation Challenge: PepsiCo Data Science
### Blake Bullwinkel
### M.S. in Data Science Student, Harvard University
### jbullwinkel@fas.harvard.edu


## 0. INTRODUCTION
# I predicted the Assessment Score based on a variety of predictors by implementing a 
# multiple linear regression (MLR) model in R. To start, I conducted data wrangling and 
# combined the features across all three Excel sheets by matching Site ID's. This allowed 
# me to consider the full range of possible predictors, including the weather variables, 
# which I incorporated by averaging each over time from the site Sowing Date to the particular 
# Assessment Date. Next, I performed variable selection using simple statistics, boxplots, 
# and hypothesis testing (t-tests, partial F-test) to decide which predictors are significant 
# enough to be included in the model. Ultimately, I built a model with 8 predictors (1 
#categorical and 7 numerical), and a single response. My final model has an R-squared value 
# of 0.848, meaning that it explains around 85% of the variation in Assessment Score. The 
# model performs relatively well, and my predictions match the actual values quite closely. 
# The main downside of my model is that it sometimes predicts negative Assessment Scores, 
# which are not possible. This is difficult to avoid with MLR, but perhaps I could attempt 
# to bound the response variable above zero by using a generalized linear model in a future 
# study.


## 1. SETUP

# Libraries
library(readxl)
library(openxlsx)
library(ggplot2)


## 2. DATA WRANGLING

# Read in the data from each Excel sheet
crop.grain.data <- read_excel("nyas-challenge-2020-data.xlsx", sheet=1)
weather.data <- read_excel("nyas-challenge-2020-data.xlsx", sheet=2)
site.data <- read_excel("nyas-challenge-2020-data.xlsx", sheet=3)

# Initialize a dataframe "wrangled.data" that will hold data from all three Excel sheets
wrangled.data <- crop.grain.data[,]

# Remove "Year" from the Site ID string so that we can match Site ID across all three dataframes
wrangled.data$`Site ID` <- gsub("\\sYear","", wrangled.data$`Site ID`)

# Append Sowing Date, Latitude, Elevation, Soil Param A, Soil Param B, and Amount Fertilizer Applied to wrangled.data by matching Site ID
wrangled.data$`Sowing Date` <- as.Date(site.data$`Sowing Date (mm/dd/year)`[match(wrangled.data$`Site ID`, site.data$`Site ID`)], format="%m/%d/%y")
wrangled.data$Latitude <- site.data$Latitude[match(wrangled.data$`Site ID`, site.data$`Site ID`)]
wrangled.data$Elevation <- site.data$`Elevation (m)`[match(wrangled.data$`Site ID`, site.data$`Site ID`)]
wrangled.data$`Soil Param A` <- site.data$`Soil Parameter A`[match(wrangled.data$`Site ID`, site.data$`Site ID`)]
wrangled.data$`Soil Param B` <- site.data$`Soil Parameter B`[match(wrangled.data$`Site ID`, site.data$`Site ID`)]
wrangled.data$`Amount Fertilizer Applied` <- site.data$`Amount Fertilizer Applied`[match(wrangled.data$`Site ID`, site.data$`Site ID`)]

# Convert date columns from String to Date data types
wrangled.data$`Assessment Date (mm/dd/year)` <- as.Date(wrangled.data$`Assessment Date (mm/dd/year)`, format="%m/%d/%y")
weather.data$`Date (mm/dd/year)` <- as.Date(weather.data$`Date (mm/dd/year)`, format="%m/%d/%y")

# Incorporate weather data by averaging each weather variable from the site Sowing Date to the particular Assessment Date
weather_letters <- c("A", "B", "C", "D", "E", "F")
for(letter in weather_letters){
  wrangled.data[[paste("Avg Weather",letter,sep=" ")]] <- apply(wrangled.data, 1, function(z) {
    mean(weather.data[[paste("Weather Variable",letter,sep=" ")]][which(weather.data$`Date (mm/dd/year)` >= z[8] & weather.data$`Date (mm/dd/year)` <= z[5])])
  })
}

# Convert Assessment Score from character to numeric data type
wrangled.data$`Assessment Score` <- as.numeric(as.character(wrangled.data$`Assessment Score`))

# Numerical variables
numer_var = c("Growth Stage", "Latitude", "Elevation", "Soil Param A", "Soil Param B", "Amount Fertilizer Applied", "Assessment Score")
numer_var_idx = c()
for(i in 1:length(numer_var)){
  numer_var_idx[i] <- which(names(wrangled.data) == numer_var[i])
}

# Average weather variables
weather_var <- c("Avg Weather A", "Avg Weather B", "Avg Weather C", "Avg Weather D", "Avg Weather E", "Avg Weather F", "Assessment Score")
weather_var_idx <- c()
for(i in 1:length(weather_var)){
  weather_var_idx[i] <- which(names(wrangled.data) == weather_var[i])
}


## 3. VARIABLE SELECTION

# Running the following tapply command, we find that the average Assessment Score for the two varieties are very similar
# (~57.5 for Variety B and ~63.8 for Variety M). A boxplot confirms that this categorical variable does not have a large
# effect on Assessment Score, so we exclude it from our model.
tapply(wrangled.data$`Assessment Score`, wrangled.data$Variety, mean)
ggplot(wrangled.data, aes(x = `Variety`, y = `Assessment Score`)) + geom_boxplot()

# Running the following tapply command, we find that there is significant variation in the average Assessment Score across
# different Assessment Types.
tapply(wrangled.data$`Assessment Score`, wrangled.data$`Assessment Type`, mean)
ggplot(wrangled.data, aes(x = `Assessment Type`, y = `Assessment Score`)) + geom_boxplot()

# Since this is a categorical variable, we incorporate it into our model by encoding it as a factor.
wrangled.data$`Assessment Type` <- factor(wrangled.data$`Assessment Type`)

# We perform t-tests to decide whether individual predictors can be excluded from the model. We first test whether the model
# coefficient for Amount Fertilizer Applied is equal to zero (null hypothesis). We obtain a p-value of 0.0189 and fail to reject
# the null hypothesis at a confidence level alpha = 0.01. Therefore, we exclude it from the model.
# Next, we test whether the coefficient for Growth Stage is equal to zero. We obtain a p-value of 0.0248 and fail to reject
# the null hypothesis at a confidence level alpha = 0.01. Therefore, we exclude it from the model.

# We perform a partial F-test using the ANOVA function to decide which Avg Weather variables to include. We find that Avg Weather 
# A, B, and D are not significant and therefore exclude them from our model.


## 4. FINAL MODEL

# Building our final MLR model using the selected predictors
final.model <- lm(`Assessment Score` ~ `Latitude`+`Elevation`+`Assessment Type`+`Soil Param A`+`Soil Param B`
                                        +`Avg Weather C`+`Avg Weather E`+`Avg Weather F`, data = wrangled.data)
summary(final.model)


## 5. MODEL PREDICTION

# Make predictions on the full dataset using the final MLR model
predicted.values <- predict.lm(final.model, wrangled.data)

# Save predictions to an Excel file (to be transferred into the 'P' column of the original Excel file)
write.xlsx(predicted.values, "assessment-score-predictions.xlsx")
