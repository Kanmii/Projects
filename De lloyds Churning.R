library(tidyverse)
library(dplyr)
library(readxl)
library(lubridate)
library(DataExplorer)
library(explore)
library(SmartEDA)

# Get all sheet names
sheet_names <- excel_sheets("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx")
sheet_names

demographics <- read_excel("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx", sheet = "Customer_Demographics")
transaction_history <- read_excel("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx", sheet = "Transaction_History")
customer_service <- read_excel("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx", sheet = "Customer_Service")
online_activity <- read_excel("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx", sheet = "Online_Activity")
churn <- read_excel("C:/Users/fatai/Desktop/Tyy/Personal Project/Customer_Churn_Data_Large.xlsx", sheet = "Churn_Status")

glimpse(demographics)
glimpse(transaction_history)
glimpse(customer_service)
glimpse(online_activity)
glimpse(churn)
view(transaction_history)
view(customer_service)


# Group transaction_history by CustomerID and summarize
trans_history <- transaction_history %>%
  group_by(CustomerID) %>%
  summarise(
    TotalTransactions = n(),
    TotalAmountSpent = sum(AmountSpent, na.rm = TRUE),
    AvgAmountSpent = mean(AmountSpent, na.rm = TRUE),
    MedianAmountSpent = median(AmountSpent, na.rm = TRUE),
    MaxAmountSpent = max(AmountSpent, na.rm = TRUE),
    MinAmountSpent = min(AmountSpent, na.rm = TRUE),
    StdAmountSpent = sd(AmountSpent, na.rm = TRUE),
    
    FirstPurchase = min(TransactionDate, na.rm = TRUE),
    LastPurchase = max(TransactionDate, na.rm = TRUE),
    
    Recency = as.numeric(Sys.Date() - max(TransactionDate)), # Days since last purchase
    Tenure = as.numeric(max(TransactionDate) - min(TransactionDate)), # How long they've been buying
    
    PurchaseFrequency = ifelse(n() > 1, as.numeric((max(TransactionDate) - min(TransactionDate)) / (n() - 1)), NA),
    #PurchaseFrequency = as.numeric(difftime(max(TransactionDate), min(TransactionDate), units = "days")) / n(),
                                   
    UniqueProductCategories = n_distinct(ProductCategory),
    .groups = 'drop'
    )

str(trans_history)
view(trans_history)


# Group customer_service by CustomerID
#customer_s <- customer_service %>%
#  group_by(CustomerID) %>%
#  summarise(Total_Interactions = n(),
# .groups = 'drop')
#view(customer_s)  
  

# Group customer_service by CustomerID and summarize
customer <- customer_service %>%
  group_by(CustomerID) %>%
  summarise(
    Total_Interactions = n(),
    Resolved_Count = sum(ResolutionStatus == "Resolved"),
    Unresolved_Count = sum(ResolutionStatus == "Unresolved"),
    First_Interaction = min(InteractionDate),
    Last_Interaction = max(InteractionDate),
    .groups = 'drop'
  )

view(customer)


# Combining sheets with leftjoin()
df = left_join(demographics, trans_history, by = "CustomerID")
df1 = left_join(df, customer, by = "CustomerID")
df2 = left_join(df1, online_activity, by = "CustomerID")
churn_data = left_join(df2, churn, by = "CustomerID")
view(churn_data)


# Removing NA's
# Method 1: Using Base R na.omit()
#churn_data_base <- na.omit(churn_data)
#view(churn_data_base)


# Method 2: Using dplyr drop_na
library(dplyr)
churn_data <- churn_data %>% drop_na()
view(churn_data)


# Removing NA in only SPecific Columns
#clean_data <- purchase_summary %>% drop_na(TotalAmountSpent, PurchaseFrequency)


# Encoding Churn column into numeric if churn is (Yes, No)
#churn_data$Churn <- ifelse(churn_data$ChurnStatus == "Yes", 1, 0)
#glimpse(churn_data$Churn)


# Explore feature Correlation (corrplot)
library("corrplot")

# Keep only numeric column for correlation
numeric_data <- churn_data %>% select(where(is.numeric))
glimpse(numeric_data)


# Correlation Matrix method 2
# Correlation matrix
corr_matrix <- cor(numeric_data, use = "complete.obs")
corr_matrix

# Extract correlations with the churn variable
churn_corr <- corr_matrix[, "ChurnStatus"]

# Create a table with correlation values
churn_corr_table <- data.frame(
  variable = names(churn_corr),
  correlation = churn_corr
)
churn_corr_table

# Filter the table
filtered_table <- churn_corr_table[abs(churn_corr_table$correlation) >= 0.1 & churn_corr_table$variable != "Churn", ]
filtered_table


names(churn_data)
churn_data %>% introduce()
churn_data %>% describe()
summary(churn_data)
churn_data %>% plot_missing()
churn_data %>% explore()
