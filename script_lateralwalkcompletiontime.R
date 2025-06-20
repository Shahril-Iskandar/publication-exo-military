# Load required libraries
library(readxl)
library(tidyr)
library(dplyr)
library(lme4)
library(ggplot2)

# Load the data from the Excel file
file_path <- "LateralWalkCompletionTime.xlsx"
data <- read_excel(file_path)

# Calculate differences for normality
data$differences <- data$WithExo - data$WithoutExo # Positive value = Longer time on Exo = No benefit
normality_shapiro <- shapiro.test(data$differences)
print(normality_shapiro)

# Perform paired t-test
ttest_result <- t.test(data$WithoutExo, data$WithExo, paired = TRUE)
print(ttest_result)

# Transpose to make long
data_long <- pivot_longer(data, 
                          cols=c("WithoutExo","WithExo"), 
                          names_to = c("Condition"), 
                          values_to = "CompletionTime"
)

withExo <- subset(data_long, Condition =="WithExo", CompletionTime, drop=TRUE)
withoutExo <- subset(data_long, Condition =="WithoutExo", CompletionTime, drop=TRUE)

data_long<- data_long %>%
  group_by(Subject) %>%
  mutate(trend = ifelse(diff(CompletionTime)>0,"Increase","Decrease"))

data_long$Condition <- factor(data_long$Condition, levels = c("WithoutExo", "WithExo"))

# Figure 8
# Plot box plot figure
ggplot(data_long, aes(x = as.factor(Condition), y = CompletionTime)) +
  geom_boxplot(fill = "darkgrey", alpha = 0.2) + 
  xlab("Condition") + ylab("Completion Time (s)") + theme_classic() +
  geom_point() + 
  geom_line(aes(group = Subject, color = trend), size = 0.5) +  # Color by trend
  scale_color_manual(values = c("Increase" = "red", "Decrease" = "green")) + # Custom colors for increase/decrease
  labs(color = "Trend") +  # Add a legend for the trend
  # geom_text(aes(label = Subject), vjust = -0.5, hjust=3) + # Add subject labels to points
  theme(legend.position = "right",
        axis.text.x = element_text(size = 12, color="black"),
        axis.text.y = element_text(size = 12, color="black"),
        axis.title.x = element_text(size = 13, color = "black"),
        axis.title.y = element_text(size = 13, color = "black"),
        legend.text = element_text(size = 11)    # Legend text (labels) font size
        ) +
  scale_x_discrete(labels = c("WithoutExo" = "Non-exo", "WithExo" = "Exo"))
# 600 x 600

# Get descriptive stats
mean_sd_results <- data_long %>%
  group_by(Condition) %>%
  get_summary_stats(CompletionTime, type = "mean_sd") %>%
  ungroup()
print(mean_sd_results)
