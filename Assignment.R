# Load necessary libraries
library(readxl)   # For reading Excel files
library(writexl)  # For writing data to Excel files
library(dplyr)    # For data manipulation
library(tidyr)    # For tidying data
library(ggplot2)  # For creating visualizations

#Part0:Data preprocessing
# Read data from Excel files
biomarkers <- read_excel("C:/Users/dell/OneDrive - University of Edinburgh/桌面/SEM1/Introductory Probability and Statistics/Final Assignment/biomarkers.xlsx")
covariates <- read_excel("C:/Users/dell/OneDrive - University of Edinburgh/桌面/SEM1/Introductory Probability and Statistics/Final Assignment/covariates.xlsx")

# Extract Patient ID from the Biomarker column, retaining only the numeric part
biomarkers$PatientID <- sub("-.*", "", biomarkers$`Biomarker`)

# Ensure that PatientID in covariates is of character type
covariates$PatientID <- as.character(covariates$PatientID)

# Merge biomarkers and covariates data by PatientID, while dropping the original PatientID column from the merged dataset
merged_data <- merge(biomarkers, covariates, by = "PatientID", all = TRUE) %>%
  select(-PatientID)  # Remove the PatientID column from the merged dataset

# Extract a new PatientID from the Biomarker column
if ("Biomarker" %in% colnames(merged_data)) {
  merged_data$PatientID <- sub("-.*", "", merged_data$`Biomarker`)  # Update to the new PatientID
} else {
  stop("Biomarker column does not exist. Please check the data.")
}

# Create a TimePoint column and retain only the necessary time points
merged_data <- merged_data %>%
  mutate(TimePoint = case_when(  # Create TimePoint based on the Biomarker string
    grepl("0weeks", Biomarker) ~ "0weeks",
    grepl("6weeks", Biomarker) ~ "6weeks",
    grepl("12months", Biomarker) ~ "12months",
    TRUE ~ NA_character_  # Set to NA for any other cases
  )) %>%
  filter(!is.na(TimePoint) )  # Filter out rows where TimePoint is NA

# Create a complete dataset combining PatientID and TimePoint
data <- merged_data %>%
  select(PatientID, TimePoint) %>%
  distinct() %>%  # Remove duplicate rows
  complete(PatientID, TimePoint = c("0weeks", "6weeks", "12months")) %>%  # Ensure all combinations are present
  left_join(merged_data, by = c("PatientID", "TimePoint")) %>%  # Join back to get original data
  arrange(as.numeric(PatientID), TimePoint)  # Sort by numeric PatientID and then by TimePoint

# Define the factor levels for TimePoint to maintain order
data$TimePoint <- factor(data$TimePoint, levels = c("0weeks", "6weeks", "12months"))

# Resort the data to ensure PatientID is in ascending order
data <- data %>%
  arrange(as.numeric(PatientID), TimePoint)

# Remove the Biomarker column from the final dataset
data <- data %>%
  select(-Biomarker)

# Detect and remove outliers using IQR method
if ("SomeNumericColumn" %in% colnames(data)) {
  Q1 <- quantile(data$SomeNumericColumn, 0.25, na.rm = TRUE)
  Q3 <- quantile(data$SomeNumericColumn, 0.75, na.rm = TRUE)
  IQR <- Q3 - Q1
  
  # Define the limits for outliers
  lower_limit <- Q1 - 1.5 * IQR
  upper_limit <- Q3 + 1.5 * IQR
  
  # Filter out outliers
  data <- data %>%
    filter(data$SomeNumericColumn >= lower_limit & data$SomeNumericColumn <= upper_limit)
}
# Display the first few rows of the final merged dataset
print("Final Complete Data:")
print(head(data))

# Export the final dataset to an Excel file
write_xlsx(data, "C:/Users/dell/OneDrive - University of Edinburgh/桌面/SEM1/Introductory Probability and Statistics/Final Assignment/Final Complete Data.xlsx")



#Part1：Statistical hypothesis testing

# Step1:Filter data for 0 weeks and remove NA values
data1 <- data %>%
  filter(TimePoint == "0weeks") %>%
  na.omit()

# Step2:Define the list of biomarkers to test
biomarkers <- c("IL-8", "VEGF-A", "OPG", "TGF-beta-1", "IL-6", "CXCL9", "CXCL1", "IL-18", "CSF-1")

# Step3:Create an empty data frame to store normality test results 
normality_results <- data.frame(
  Biomarker = character(),
  Male_pValue = numeric(),
  Female_pValue = numeric(),
  Male_Distribution = character(),
  Female_Distribution = character(),
  stringsAsFactors = FALSE
)

# Create an empty data frame to store test results 
test_results <- data.frame(
  Biomarker = character(),
  Test_Type = character(),
  t_statistic = numeric(),
  p_value = numeric(),
  Conclusion = character(),
  stringsAsFactors = FALSE
)

# Create an empty data frame to store Bonferroni correction results
Bonferroni_results <- data.frame(
  Biomarker = character(),
  Test_Type = character(),
  t_statistic = numeric(),
  p_value = numeric(),
  Adjusted_Conclusion = character(),
  stringsAsFactors = FALSE
)

# Set options to display numbers in decimal format instead of scientific notation
options(scipen = 999)

# Step4:Conduct normality tests for each biomarker and test hypotheses
for (biomarker in biomarkers) {
  # Filter data by gender
  male_data <- data1[data1$`Sex (1=male, 2=female)` == 1, biomarker, drop = FALSE]
  female_data <- data1[data1$`Sex (1=male, 2=female)` == 2, biomarker, drop = FALSE]
  
  # Conduct normality test for males
  normality_test_male <- shapiro.test(male_data[[1]])
  distribution_type_male <- ifelse(normality_test_male$p.value > 0.05, "Normal", "Non-normal")
  
  # Conduct normality test for females
  normality_test_female <- shapiro.test(female_data[[1]])
  distribution_type_female <- ifelse(normality_test_female$p.value > 0.05, "Normal", "Non-normal")
  
  # Save normality test results to Table(normality_results)
  normality_results <- rbind(normality_results, data.frame(
    Biomarker = biomarker,
    Male_pValue = normality_test_male$p.value,
    Female_pValue = normality_test_female$p.value,
    Male_Distribution = distribution_type_male,
    Female_Distribution = distribution_type_female,
    stringsAsFactors = FALSE
  ))
  
  # Prepare for group comparison
  test_result <- NULL
  test_type <- NULL
  
  if (distribution_type_male == "Normal" && distribution_type_female == "Normal") {
    # Conduct t-test if both groups are normally distributed
    test_result <- t.test(male_data[[1]], female_data[[1]], var.equal = TRUE)
    test_type <- "t-test"
  } else {
    # Conduct Mann-Whitney U test if either group is not normally distributed
    test_result <- wilcox.test(male_data[[1]], female_data[[1]])
    test_type <- "Mann-Whitney U test"
  }
  
  # Save comparison results to Table(test_results)
  test_results <- rbind(test_results, data.frame(
    Biomarker = biomarker,
    Test_Type = test_type,
    t_statistic = ifelse(test_type == "t-test", test_result$statistic, NA),
    p_value = test_result$p.value,
    Conclusion = ifelse(test_result$p.value < 0.05, "Reject H0", "Fail to reject H0"),
    stringsAsFactors = FALSE
  ))
  
# Step5:Output initial test results
  # Output the normality test results
  print("Normality test results:")
  print(normality_results)
  
  # Output the group test results 
  print("Group test results :")
  print(test_results)
  
# Step6:Bonferroni correction
  
# Store Bonferroni correction results
Bonferroni_results <- rbind(Bonferroni_results, data.frame(
  Biomarker = biomarker,
  Test_Type = test_type,
  t_statistic = ifelse(test_type == "t-test", test_result$statistic, NA),
  p_value = test_result$p.value,
  Adjusted_Conclusion = NA,  # Placeholder for now
  stringsAsFactors = FALSE
))
}

# Calculate Bonferroni-adjusted alpha level
num_tests <- length(biomarkers)
alpha <- 0.05
alpha_adjusted <- alpha / num_tests

# Add adjusted conclusions based on Bonferroni correction
Bonferroni_results$Adjusted_Conclusion <- ifelse(Bonferroni_results$p_value < alpha_adjusted, "Reject H0", "Fail to reject H0")

# Print Bonferroni-adjusted significance level
cat("Bonferroni-adjusted alpha level:", alpha_adjusted, "\n")

# Output Bonferroni correction results 
print("Bonferroni correction results:")
print(Bonferroni_results)



# Part2: Regression modelling

# Step1:Filter data for the 0weeks(at inclusion) TimePoint 
data2 <- data %>%
  filter(TimePoint == "0weeks") %>%
  na.omit()  # Remove rows that contain any missing values

# Step2:Set the random seed for reproducibility of the results
set.seed(123)

# Step3:Split the filtered dataset into training (80%) and testing (20%) datasets
train_indices <- sample(1:nrow(data2), size = 0.8 * nrow(data2))  # Randomly sample 80% of the data for training
train_data <- data2[train_indices, ]  # Assign the selected 80% of the data to train_data
test_data <- data2[-train_indices, ]  # Assign the remaining 20% of the data to test_data

# Step4:Fit a linear regression model using the training dataset
# The dependent variable is 'Vas-12months', and the independent variables are biomarkers and covariates
model <- lm(`Vas-12months` ~ `IL-8` + `VEGF-A` + `OPG` + `TGF-beta-1` + `IL-6` + `CXCL9` + 
              `CXCL1` + `IL-18` + `CSF-1` + `Age` + `Sex (1=male, 2=female)` + `Smoker (1=yes, 2=no)` + 
              `VAS-at-inclusion`, data = train_data)

# Step5:View the summary of the model, which includes coefficient estimates, R-squared value, etc.
summary(model)

# Step6: Extract the coefficient table from the model summary
coef_table <- summary(model)$coefficients  # Extract the coefficients (betas) of the model
# Create a dataframe with beta labels starting from β0 and preserve column names
coef_table <- data.frame(
  βi = paste0("β", seq(0, nrow(coef_table)-1)),  # Create labels from β0 to βn
  coef_table,  # Add the original coefficient table to the dataframe
  check.names = FALSE  # Preserve original column names
)
# View the table with labels
print(coef_table)  # Print the coefficient table with beta labels

# Step7:Plot diagnostic plots for the linear model
par(mfrow = c(2, 2))  # Set up a 2x2 layout for the plots
plot(model)  # Generate the four default diagnostic plots: residuals vs fitted, normal Q-Q, scale-location, and residuals vs leverage

# Step8:Retrieve and print the R-squared value of the model (coefficient of determination)
r_squared <- summary(model)$r.squared  # Extract R-squared from the model summary
print(paste("R-squared:", r_squared))  # Print R-squared value to assess model fit

# Step9:Use the trained model to make predictions on the test data
predictions <- predict(model, newdata = test_data)  # Generate predicted values based on test data

# Step10:Create a data frame with actual values and predicted values
results1 <- data.frame(Actual_Value = test_data$`Vas-12months`, Predicted_Value = predictions)  

# Step11:Compute the absolute error between actual and predicted values
ae <- abs(results1$Actual_Value - results1$Predicted_Value)  
print(paste("Absolute Error:", ae))  
# Step12:Compute the squared error 
se <- (results1$Actual_Value - results1$Predicted_Value)^2  
print(paste("Squared Error:", se))  
# Step13:Compute the Mean Absolute Error (MAE)
mae <- mean(abs(results1$Actual_Value - results1$Predicted_Value))  
print(paste("Mean Absolute Error:", mae))  
# Step14:Compute the Mean Squared Error (MSE)
mse <- mean((results1$Actual_Value - results1$Predicted_Value)^2)  
print(paste("Mean Squared Error:", mse))  

# Step15:Create a final table that includes all values above
predict_results <- data.frame(PatientID = test_data$PatientID, 
                     Actual_Value = test_data$`Vas-12months`, 
                     Predicted_Value = predictions, 
                     Absolute_Error = ae, 
                     Squared_Error = se)

# Step16:Show the table
print(predict_results)  





