library(tidyverse)
library(forecast)
library(tseries)
library(DataExplorer)
library(explore)
library(SmartEDA)
library(ggplot2)
library(plotly)
library(lubridate)
library(scales)  # For better axis formatting

# === Load Dataset === 
data <- read.csv("C:/Users/fatai/Desktop/GROUP/Dangote_Mtn_Stock_6yrs.csv")
glimpse(data)


# === Change Date Format === 
#data$Date <- as.Date(data$Date)
data$Date <- parse_date_time(data$Date, orders = c( "mdy", "mdY"))
data <- data %>% arrange(Date)
glimpse(data)


# Rename for easier reference
data <- data %>%
  rename(
    Dangote_Close = Dangote_cement,
    MTN_Close = Mtn
  )

# === Exploratory Statistics === 
data %>% introduce()
data %>% describe_all()
data %>% plot_intro()
data %>% plot_density()
data %>% plot_histogram()
data %>% explore_all()


# === Calculate Daily Returns ===
data <- data %>%
  mutate(
    MTN_Return = (MTN_Close - lag(MTN_Close)) / lag(MTN_Close) * 100,
    Dangote_Return = (Dangote_Close - lag(Dangote_Close)) / lag(Dangote_Close) * 100
  ) %>%
  filter(is.finite(MTN_Return), is.finite(Dangote_Return))  # Remove NA and Inf
glimpse(data)


# === Summary Statistics ===
summary_stats <- data %>% 
  select(MTN_Return, Dangote_Return) %>% 
  summary()
print(summary_stats)


# === Histograms ===
ggplot(data, aes(x = MTN_Return)) +
  geom_histogram(bins = 20, fill = "steelblue", alpha = 0.7) +
  labs(title = "Histogram of MTN Daily Returns", x = "Return (%)", y = "Frequency") +
  theme_minimal()

ggplot(data, aes(x = Dangote_Return)) +
  geom_histogram(bins = 20, fill = "darkorange", alpha = 0.7) +
  labs(title = "Histogram of Dangote Daily Returns", x = "Return (%)", y = "Frequency") +
  theme_minimal()


# === Boxplots ===
ggplot(data, aes(y = MTN_Return)) +
  geom_boxplot(fill = "lightblue") +
  labs(title = "Boxplot of MTN Daily Returns", y = "Return (%)") +
  theme_minimal()

ggplot(data, aes(y = Dangote_Return)) +
  geom_boxplot(fill = "orange") +
  labs(title = "Boxplot of Dangote Daily Returns", y = "Return (%)") +
  theme_minimal()


# === Line Plot of Daily Returns ===
daily_lineplot <- ggplot(data, aes(x = Date)) +
  geom_line(aes(y = MTN_Return, color = "MTN")) +
  geom_line(aes(y = Dangote_Return, color = "Dangote")) +
  labs(title = "Daily Returns Over Time", x = "Date", y = "Return (%)") +
  scale_color_manual(values = c("MTN" = "blue", "Dangote" = "orange")) +
  theme_minimal()
daily_lineplot
ggplotly(daily_lineplot)


# === QQ Plots ===
par(mfrow = c(1, 2), mar = c(4, 4, 2, 1))  # Two plots side-by-side
qqnorm(data$MTN_Return, main = "QQ Plot - MTN Returns")
qqline(data$MTN_Return, col = "blue")

qqnorm(data$Dangote_Return, main = "QQ Plot - Dangote Returns")
qqline(data$Dangote_Return, col = "orange")
par(mfrow = c(1, 1))  # Reset layout


# === Convert data into time series format === 
mtn_ts <- ts(data$MTN_Return, frequency = 252) # Assuming daily data
dangote_ts <- ts(data$Dangote_Return, frequency = 252) # Assuming daily data


# === Perform Augmented Dickey-Fuller (ADF) Test for Stationarity === 
adf_test_mtn <- adf.test(mtn_ts)
print(adf_test_mtn)

adf_test_dangote <- adf.test(dangote_ts)
print(adf_test_dangote)


# Differencing if Series is Non-Stationary (Based on ADF Test)
# If p-value is > 0.05, the series is non-stationary, and we will difference the series
#if(adf_test_mtn$p.value > 0.05) {
 # mtn_ts <- diff(mtn_ts)
#}

#if(adf_test_dangote$p.value > 0.05) {
 # dangote_ts <- diff(dangote_ts)
#}


# Re-check ADF after differencing if necessary
#adf_test_mtn <- adf.test(mtn_ts)
#print(adf_test_mtn)
#adf_test_dangote <- adf.test(dangote_ts)
#print(adf_test_dangote)


# === Forecasting using ARIMA (Auto ARIMA) === 
# Fit ARIMA model
arima_mtn <- auto.arima(mtn_ts, trace = TRUE)
arima_dangote <- auto.arima(dangote_ts, trace = TRUE)


# === Forecast the next 30 days === 
forecast_mtn <- forecast(arima_mtn, h = 30)
forecast_mtn
forecast_dangote <- forecast(arima_dangote, h = 30)
forecast_dangote


# === Plot Forecasts === 
# MTN Forecast
autoplot(forecast_mtn) + 
  labs(title = "30-Day Forecast for MTN Returns", x = "Day", y = "Predicted Return (%)") +
  theme_minimal()

# Dangote Forecast
autoplot(forecast_dangote) + 
  labs(title = "30-Day Forecast for Dangote Returns", x = "Day", y = "Predicted Return (%)") +
  theme_minimal()


# === Evaluate Forecast Model - Residual Diagnostics === 
# Residual plot and Ljung-Box test for autocorrelation
checkresiduals(arima_mtn)
checkresiduals(arima_dangote)


# === Summary of models === 
summary(arima_mtn)
summary(arima_dangote)


