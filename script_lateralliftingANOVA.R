# Load required libraries
library(readxl)
library(tidyr)
library(dplyr)
library(lme4)
library(multcomp)
library(ggplot2)
library(stringr)
library(purrr)
library(rstatix)
library(car)

# Load the data from the Excel file
file_path <- "Data_LateralLift.xlsx"
data <- read_excel(file_path)

data <- data %>%
  mutate(MUSCLE = ifelse(MUSCLE == "L-Lliocostalis", "L-Iliocostalis", MUSCLE)) %>%
  mutate(MUSCLE = ifelse(MUSCLE == "R-Lliocostalis", "R-Iliocostalis", MUSCLE)) %>%
  ungroup()

data <- data %>%
  mutate(SUBJECT = as.character(SUBJECT),            # Convert to character
         SUBJECT = ifelse(str_detect(SUBJECT, "S[0-9]$"), # Check if it's S1 to S9
                          str_replace(SUBJECT, "S([0-9])", "S0\\1"),  # Add leading zero
                          SUBJECT))  # Keep S10 unchanged

# Calculate the average of 3 trials
data <- data %>%
  group_by(SUBJECT, MUSCLE, CONDITIONS) %>%
  summarise(Average_MVC = mean(`Normalised result(%MVC)`, na.rm = TRUE)) %>%
  ungroup()

data$Side <- sub("-.*", "", data$MUSCLE)
data$Muscle <- sub("^[LR]-", "", data$MUSCLE)
data$SUBJECT <- factor(data$SUBJECT)
# data$TRIAL <- factor(data$TRIAL)

# Convert CONDITIONS to factor
data$CONDITIONS <- factor(data$CONDITIONS, labels = c("Non-exo", "Exo"))

data$Side <- factor(data$Side, labels = c("Left", "Right"))

# List of unique muscles
muscles <- unique(data$Muscle)

# Log transform data
data$logMVC <- log(data$Average_MVC)

muscles <- c("Iliocostalis", "Multifidus", "Longissimus") # or unique(data$MUSCLE)
conditions <- c("C1", "C2")
heights <- c("M1", "M2")
loads <- c("L1", "L2")
trials <- c("T1", "T2", "T3")
phases <- c("P1", "P2", "P3")
sides <- c("Left", "Right")

# Create all possible combinations
anova_combinations <- expand.grid(Muscle = muscles)

# Define function to process each combinations
perform_anova <- function(muscle) {
  # Filter data
  data_subset <- filter(data, Muscle == muscle)
  
  if(nrow(data_subset) > 0) {
    # Descriptive analysis
    mean_sd_results <- data_subset %>%
      group_by(CONDITIONS, Side) %>%
      get_summary_stats(Average_MVC, type = "mean_sd") %>%
      ungroup()
    
    shapiro_results <- data_subset %>%
      group_by(CONDITIONS, Side) %>%
      shapiro_test(Average_MVC) %>%
      ungroup()
    
    levene_results <- leveneTest(Average_MVC ~ CONDITIONS * Side, data = data_subset)
    
    anova_results <- anova_test(
      data = data_subset, dv = Average_MVC, wid = SUBJECT,
      within = c(CONDITIONS, Side), effect.size = "pes"
    )
  }
  
  return(list(
    left_nonexo_mean = filter(mean_sd_results, CONDITIONS=="Non-exo", Side=="Left")$mean,
    left_nonexo_sd = filter(mean_sd_results, CONDITIONS=="Non-exo", Side=="Left")$sd,
    right_nonexo_mean = filter(mean_sd_results, CONDITIONS=="Non-exo", Side=="Right")$mean,
    right_nonexo_sd = filter(mean_sd_results, CONDITIONS=="Non-exo", Side=="Right")$sd,
    left_exo_mean = filter(mean_sd_results, CONDITIONS=="Exo", Side=="Left")$mean,
    left_exo_sd = filter(mean_sd_results, CONDITIONS=="Exo", Side=="Left")$sd,
    right_exo_mean = filter(mean_sd_results, CONDITIONS=="Exo", Side=="Right")$mean,
    right_exo_sd = filter(mean_sd_results, CONDITIONS=="Exo", Side=="Right")$sd,
    conditions_pvalue = anova_results$p[1], conditions_eta = anova_results$pes[1],
    sides_pvalue = anova_results$p[2], sides_eta = anova_results$pes[2],
    interaction_pvalue = anova_results$p[3], interaction_eta = anova_results$pes[3],
    levene_pvalue = levene_results$`Pr(>F)`[1],
    normality_left_nonexo = shapiro_results$p[1],
    normality_right_nonexo = shapiro_results$p[2],
    normality_left_exo = shapiro_results$p[3],
    normality_right_exo = shapiro_results$p[4]
  ))
}

anova_results_list <- anova_combinations %>%
  mutate(results = pmap(list(muscles),
                        ~ perform_anova(..1)))

anova_results_list <- anova_results_list %>%
  mutate(left_nonexo_mean = map_dbl(results, "left_nonexo_mean"), left_nonexo_sd = map_dbl(results, "left_nonexo_sd"),
         right_nonexo_mean = map_dbl(results, "right_nonexo_mean"), right_nonexo_sd = map_dbl(results, "right_nonexo_sd"),
         left_exo_mean = map_dbl(results, "left_exo_mean"), left_exo_sd = map_dbl(results, "left_exo_sd"),
         right_exo_mean = map_dbl(results, "right_exo_mean"), right_exo_sd = map_dbl(results, "right_exo_sd"),
         conditions_pvalue = map_dbl(results, "conditions_pvalue"), conditions_eta = map_dbl(results, "conditions_eta"),
         sides_pvalue = map_dbl(results, "sides_pvalue"), sides_eta = map_dbl(results, "sides_eta"),
         interaction_pvalue = map_dbl(results, "interaction_pvalue"), interaction_eta = map_dbl(results, "interaction_eta"),
         levene_pvalue = map_dbl(results, "levene_pvalue"), 
         normality_left_nonexo = map_dbl(results, "normality_left_nonexo"), normality_right_nonexo = map_dbl(results, "normality_right_nonexo"),
         normality_left_exo = map_dbl(results, "normality_left_exo"), normality_right_exo = map_dbl(results, "normality_right_exo"),
         results=NULL
         )

# Filter
filter(anova_results_list, conditions_pvalue < 0.05)

anova_summary_data <- data %>%
  group_by(Muscle, CONDITIONS, Side) %>%
  summarise(
    count= n(),
    mean = mean(Average_MVC, na.rm=TRUE),
    sd = sd(Average_MVC, na.rm=TRUE)
  )
print(anova_summary_data)

anova_summary_data <- anova_summary_data %>%
  mutate(CONDITIONS = factor(CONDITIONS, levels = c("Exo", "Non-exo")))

# Figure 9
# Plot each muscle with conditions and sides
ggplot(anova_summary_data) +
  geom_bar(aes(x = Side, y = mean, fill = CONDITIONS), stat = "identity", position = "dodge") +
  geom_errorbar(aes(x = Side, ymin = mean - sd, ymax = mean + sd, group = CONDITIONS), 
                position = position_dodge(width = 0.9), width = 0.3) +
  facet_grid(~Muscle) +
  labs(x = "Side",
       y = "Normalised %MVC") +
  theme_classic() +
  scale_fill_manual(name="Condition",
                   labels=c("Exo", "Non-exo"),
                   values=c("blue", "red")) +
  theme(legend.position = "right",
        strip.text = element_text(size = 12, color="black"), # Change the muscle name fontsize
        axis.text.x = element_text(size = 12, color="black"),
        axis.text.y = element_text(size = 12, color="black"),
        axis.title.x = element_text(size = 13, color = "black"),
        axis.title.y = element_text(size = 13, color = "black"),
        legend.text = element_text(size = 11)    # Legend text (labels) font size
  )
# 800 x 500

# Figure 10
# Plot per subject for each CONDITION and Muscle
ggplot(data, aes(x = CONDITIONS, y = Average_MVC, color = Muscle, group = interaction(Side, Muscle))) +
  geom_point(aes(shape = Side), size = 2.2) +   # Use different shapes for Left/Right sides
  geom_line(aes(linetype = Side), size = 0.5) + # Connect points for Left/Right sides
  facet_wrap(~ SUBJECT, scale="free", nrow=2, ncol=5) +                     # Facet by subject
  labs(title = "Normalised MVC by Conditions and Sides for Each Subject and Muscle",
       x = "Condition",
       y = "Normalised %MVC",
       color = "Muscle", shape = "Side", linetype = "Side") +
  theme_classic() +
  theme(legend.position = "right",
        strip.text = element_text(size = 12, color="black"), # Change the muscle name fontsize
        axis.text.x = element_text(size = 12, color="black"),
        axis.text.y = element_text(size = 12, color="black"),
        axis.title.x = element_text(size = 13, color = "black"),
        axis.title.y = element_text(size = 13, color = "black"),
        legend.text = element_text(size = 11)    # Legend text (labels) font size
  ) +
  scale_color_manual(values = c("Iliocostalis" = "red", "Multifidus" = "blue", "Longissimus" = "darkgreen"))
# 1200 x 600