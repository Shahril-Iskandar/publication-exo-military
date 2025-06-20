# Load required libraries
library(readxl)
library(dplyr)
library(purrr)
library(writexl)
library(ggplot2)
library(car)
library(tidyverse)
library(ggpubr)
library(rstatix)
library(cowplot)

# Load the data from the Excel file
file_path <- "Data_VerticalLift.xlsx"
data <- read_excel(file_path)

# Rename the MUSCLE column
data$MUSCLE <- gsub("LT ", "L-", data$MUSCLE) # Replace LT with L-
data$MUSCLE <- gsub("RT ", "R-", data$MUSCLE) # Replace RT with R-
data$MUSCLE <- gsub("LUMBAR ES", "Iliocostalis", data$MUSCLE)  # Replace LUMBAR ES with Iliocostalis
data$MUSCLE <- gsub("THORACIC ES", "Longissimus", data$MUSCLE) # Replace THORACIC ES with Longissimus
data$MUSCLE <- gsub("MULTIFIDII", "Multifidus", data$MUSCLE)   # Replace MULTIFIDII with Multifidus

names(data)[names(data) == "Normalized Results (%MVC)"] <- "NORMALIZED_MVC"

# Separate the side and muscle name
data <- data %>%
  mutate(SIDE = ifelse(grepl("^L-", MUSCLE), "Left", "Right"),
         MUSCLE = sub("^[LR]-", "", MUSCLE))

# Define all combinations of muscles, heights, loads, trials, and phases
muscles <- c("Iliocostalis", "Multifidus", "Longissimus") # or unique(data$MUSCLE)
conditions <- c("C1", "C2")
heights <- c("M1", "M2")
loads <- c("L1", "L2")
trials <- c("T1", "T2", "T3")
phases <- c("P1", "P2", "P3")

################## Paired t-test: Use the average of 3 trials ##################
# Filter the data based on muscle, condition, load, and height
data_filtered <- data %>%
  filter(MUSCLE %in% muscles,
         CONDITION %in% conditions,
         HEIGHT %in% heights,
         LOAD %in% loads,
         PHASE %in% phases)

# Group by MUSCLE, CONDITION, LOAD, HEIGHT, PHASE, SIDE, and SUBJECT, then compute the average of the 3 trials
data_avg_trials <- data_filtered %>%
  group_by(MUSCLE, CONDITION, LOAD, HEIGHT, PHASE, SIDE, SUBJECT) %>%
  summarise(avg_value = mean(NORMALIZED_MVC, na.rm = TRUE)) %>%
  ungroup()
print(data_avg_trials)

# Do log transform on avg_trials
data_avg_trials$avg_value_log <- log(data_avg_trials$avg_value)

# Use expand.grid to get all combinations of parameters
param_combinations_avg <- expand.grid(MUSCLE = muscles, CONDITION = conditions, HEIGHT = heights, LOAD = loads, 
                                      PHASE = phases, stringsAsFactors = FALSE)

# Create function to perform paired t-test for each combination, using avg trial
perform_paired_ttest_avg <- function(muscle_name, condition, height, load, phase) {
  # Filter data for Left side
  data_left <- data_avg_trials %>%
    filter(MUSCLE == muscle_name, CONDITION == condition, HEIGHT == height, LOAD == load, 
           PHASE == phase, SIDE == "Left")
  
  # Filter data for Right side2
  data_right <- data_avg_trials %>%
    filter(MUSCLE == muscle_name, CONDITION == condition, HEIGHT == height, LOAD == load, 
           PHASE == phase, SIDE == "Right")
  # print(nrow(data_left))
  # print(nrow(data_right))
  
  # Check normality
  normality_shapiro <- shapiro.test(data_left$avg_value - data_right$avg_value)
  
  # Ensure both left and right data exist and perform paired t-test
  if(nrow(data_left) > 0 & nrow(data_right) > 0) {
    # print(nrow(data_left))
    # print(nrow(data_right))
    ttest_result <- t.test(data_left$avg_value, data_right$avg_value, paired = TRUE) # Paired t-test
    wilcoxon_result <- wilcox.test(data_left$avg_value, data_right$avg_value, paired = TRUE) # Wilcoxon test
    return(list(p_value = ttest_result$p.value, normality = normality_shapiro$p.value, wilcoxon = wilcoxon_result$p.value))  # list p-value of paired ttest and normality
  } else {
    return(list(p_value = NA, normality = NA, wilcoxon = NA))  # Return NA if not enough data
  }
}

# Apply the paired t-test function to each combination
param_combinations_avg <- param_combinations_avg %>%
  mutate(results = pmap(list(MUSCLE, CONDITION, HEIGHT, LOAD, PHASE), 
                        ~ perform_paired_ttest_avg(..1, ..2, ..3, ..4, ..5)))

# Extract p-values and normality status from results
param_combinations_avg <- param_combinations_avg %>%
  mutate(p_value = map_dbl(results, "p_value"),
         normality = map_dbl(results, "normality"),
         wilcoxon = map_dbl(results, "wilcoxon"),
         results = NULL)

print(param_combinations_avg)

# Export to Excel
# write_xlsx(param_combinations_avg, "VerticalLifting_PairedTtest.xlsx")

# Find where normality < 0.05 and wilcoxon p-value < 0.05
filter(param_combinations_avg, normality<0.05, wilcoxon <0.05)
# Result: None. All no significant differences

# Find where paired ttest p_value < 0.05
filter(param_combinations_avg, p_value<0.05)
# Result: Longissumis C2M1L1P2 and C2M2L1P3, Iliocostalis C1M2L1P3

# Conclusion: No significant differences in most = can combine left and right value for ANOVA
data_avg_both_side <- data_avg_trials %>%
  group_by(MUSCLE, CONDITION, LOAD, HEIGHT, PHASE, SUBJECT) %>%
  summarise(both_value = mean(avg_value, na.rm = TRUE))
print(data_avg_both_side)

# Plot paired t-test where results are significant
plot_data <- filter(data_avg_trials, MUSCLE == "Iliocostalis", HEIGHT=="M2", CONDITION == "C2", LOAD == "L1", PHASE=="P3")
plot_data<- plot_data %>%
  group_by(SUBJECT) %>%
  mutate(trend = ifelse(diff(avg_value)>0,"Increase","Decrease"))

################## Perform 2x3 ANOVA using average of 3 trials ##################
# Log transform data
data_avg_both_side$both_value_log <- log(data_avg_both_side$both_value)

data_avg_both_side <- data_avg_both_side %>%
  convert_as_factor(CONDITION, PHASE, SUBJECT)

# Create all possible combinations
anova_combinations <- expand.grid(MUSCLE = muscles, HEIGHT = heights, LOAD = loads)

# Define the function to process each combination
perform_anova <- function(muscle, height, load) {
  # Filter data based on the current combination
  data_subset <- data_avg_both_side %>%
    filter(MUSCLE == muscle, HEIGHT == height, LOAD == load)
  
  if(nrow(data_subset) > 0) {
    # Perform the analyses
    mean_sd_results <- data_subset %>%
      group_by(CONDITION, PHASE) %>%
      get_summary_stats(both_value, type = "mean_sd") %>%
      ungroup()
    
    outlier_results <- data_subset %>%
      group_by(CONDITION, PHASE) %>%
      identify_outliers(both_value) %>%
      ungroup()
    
    shapiro_results <- data_subset %>%
      group_by(CONDITION, PHASE) %>%
      shapiro_test(both_value) %>%
      ungroup()
    
    qqplot_result <- ggqqplot(data_subset, "both_value", ggtheme = theme_bw()) +
      facet_grid(PHASE ~ CONDITION, labeller = "label_both")
    
    levene_results <- leveneTest(both_value ~ CONDITION * PHASE, data = data_subset)
    
    anova_results <- anova_test(
      data = data_subset, dv = both_value, wid = SUBJECT,
      within = c(CONDITION, PHASE), effect.size = "pes"
    )
    
    bxp <- ggboxplot(
      data_subset, x = "PHASE", y = "both_value",
      color = "CONDITION", palette = "jco"
    )
    
    bxp_stats <- bxp + labs(
      subtitle = get_test_label(anova_results, detailed = TRUE)
    )
    
    # Option 1: Return in an object format
    # Compile all results into a list
    # return(list(
    #   summary_stats = mean_sd_results,
    #   outlier = outlier_results,
    #   shapiro = shapiro_results,
    #   levene = levene_results,
    #   box_plot = bxp_stats,
    #   qq_plot = qqplot_result,
    #   anova_result = anova_results
    # )
    # )
    
    # Option 2: Return individual results, without normality
    return(list(
      p1_nonexo_mean = filter(mean_sd_results, CONDITION=="C1", PHASE=="P1")$mean ,
      p1_nonexo_sd = filter(mean_sd_results, CONDITION=="C1", PHASE=="P1")$sd,
      p1_exo_mean = filter(mean_sd_results, CONDITION=="C2", PHASE=="P1")$mean,
      p1_exo_sd = filter(mean_sd_results, CONDITION=="C2", PHASE=="P1")$sd,
      p2_nonexo_mean = filter(mean_sd_results, CONDITION=="C1", PHASE=="P2")$mean,
      p2_nonexo_sd = filter(mean_sd_results, CONDITION=="C1", PHASE=="P2")$sd,
      p2_exo_mean = filter(mean_sd_results, CONDITION=="C2", PHASE=="P2")$mean,
      p2_exo_sd = filter(mean_sd_results, CONDITION=="C2", PHASE=="P2")$sd,
      p3_nonexo_mean = filter(mean_sd_results, CONDITION=="C1", PHASE=="P3")$mean,
      p3_nonexo_sd = filter(mean_sd_results, CONDITION=="C1", PHASE=="P3")$sd,
      p3_exo_mean = filter(mean_sd_results, CONDITION=="C2", PHASE=="P3")$mean,
      p3_exo_sd = filter(mean_sd_results, CONDITION=="C2", PHASE=="P3")$sd,
      conditions_pvalue = anova_results$ANOVA$p[1], conditions_eta = anova_results$ANOVA$pes[1],
      phases_pvalue = anova_results$ANOVA$p[2], phases_eta = anova_results$ANOVA$pes[2],
      interaction_pvalue = anova_results$ANOVA$p[3], interaction_eta = anova_results$ANOVA$pes[3],
      levene_pvalue = levene_results$`Pr(>F)`[1],
      normality_pvalue_p1_nonexo = shapiro_results$p[1],
      normality_pvalue_p2_nonexo = shapiro_results$p[2],
      normality_pvalue_p3_nonexo = shapiro_results$p[3],
      normality_pvalue_p1_exo = shapiro_results$p[4],
      normality_pvalue_p2_exo = shapiro_results$p[5],
      normality_pvalue_p3_exo = shapiro_results$p[6]
    ))
  } else {
    return(list(      p1_nonexo_mean = NA ,
                      p1_nonexo_sd = NA,
                      p1_exo_mean = NA,
                      p1_exo_sd = NA,
                      p2_nonexo_mean = NA,
                      p2_nonexo_sd = NA,
                      p2_exo_mean = NA,
                      p2_exo_sd = NA,
                      p3_nonexo_mean = NA,
                      p3_nonexo_sd = NA,
                      p3_exo_mean = NA,
                      p3_exo_sd = NA,
                      conditions_pvalue = NA, conditions_eta = NA,
                      phases_pvalue = NA, phases_eta = NA,
                      interaction_pvalue = NA, interaction_eta = NA,
                      levene_pvalue = NA, 
                      normality_pvalue_p1_nonexo = NA, 
                      normality_pvalue_p2_nonexo = NA, 
                      normality_pvalue_p3_nonexo = NA,
                      normality_pvalue_p1_exo = NA, 
                      normality_pvalue_p2_exo = NA, 
                      normality_pvalue_p3_exo = NA))
  }
}

data_avg_both_side <- data_avg_both_side %>% ungroup()

# Apply the function to each combination and store the results
# If Option 1 is selected
# results_list <- pmap(
#   anova_combinations,
#   ~ perform_anova(..1, ..2, ..3)
# )

# If Option 2 is selected
anova_results_list <- anova_combinations %>%
  mutate(results = pmap(list(MUSCLE, HEIGHT, LOAD),
                        ~ perform_anova(..1, ..2, ..3)))

# Extract p-values and normality status from results
anova_results_list <- anova_results_list %>%
  mutate(p1_nonexo_mean = map_dbl(results, "p1_nonexo_mean"), p1_nonexo_sd = map_dbl(results, "p1_nonexo_sd"),
         p1_exo_mean = map_dbl(results, "p1_exo_mean"), p1_exo_sd = map_dbl(results, "p1_exo_sd"),
         p2_nonexo_mean = map_dbl(results, "p2_nonexo_mean"), p2_nonexo_sd = map_dbl(results, "p2_nonexo_sd"),
         p2_exo_mean = map_dbl(results, "p2_exo_mean"), p2_exo_sd = map_dbl(results, "p2_exo_sd"),
         p3_nonexo_mean = map_dbl(results, "p3_nonexo_mean"), p3_nonexo_sd = map_dbl(results, "p3_nonexo_sd"),
         p3_exo_mean = map_dbl(results, "p3_exo_mean"), p3_exo_sd = map_dbl(results, "p3_exo_sd"),
         conditions_pvalue = map_dbl(results, "conditions_pvalue"), conditions_eta = map_dbl(results, "conditions_eta"),
         phases_pvalue = map_dbl(results, "phases_pvalue"), phases_eta = map_dbl(results, "phases_eta"),
         interaction_pvalue = map_dbl(results, "interaction_pvalue"), interaction_eta = map_dbl(results, "interaction_eta"),
         levene_pvalue = map_dbl(results, "levene_pvalue"), 
         normality_pvalue_p1_nonexo = map_dbl(results, "normality_pvalue_p1_nonexo"),
         normality_pvalue_p2_nonexo = map_dbl(results, "normality_pvalue_p2_nonexo"),
         normality_pvalue_p3_nonexo = map_dbl(results, "normality_pvalue_p3_nonexo"),
         normality_pvalue_p1_exo = map_dbl(results, "normality_pvalue_p1_exo"),
         normality_pvalue_p2_exo = map_dbl(results, "normality_pvalue_p2_exo"),
         normality_pvalue_p3_exo = map_dbl(results, "normality_pvalue_p3_exo"),
         results = NULL)

anova_results_list <- anova_results_list %>% arrange(MUSCLE, HEIGHT)
anova_results_list

# Export to excel
# write_xlsx(anova_results_list, "VerticalLifting_ANOVA_table.xlsx")

# filter(anova_results_list, normality_pvalue_p1_nonexo < 0.05, normality_pvalue_p2_nonexo < 0.05, normality_pvalue_p3_nonexo < 0.05)
filter(anova_results_list, conditions_pvalue < 0.05) # Results: 0
filter(anova_results_list, phases_pvalue < 0.05) # Results: All significant
# Conclusion: No significant in Condition, only in Phase for all

# Get mean and sd of each permutation
anova_summary_data <- data_avg_both_side %>%
  group_by(MUSCLE, CONDITION, PHASE, HEIGHT, LOAD) %>%
  summarise(
    count= n(),
    mean = mean(both_value, na.rm=TRUE),
    sd = sd(both_value, na.rm=TRUE)
  )
print(anova_summary_data)

# Prevent scientific notation
# options(scipen=999)

################## Plot ANOVA ##################
# Replace values in the HEIGHT column

#### Figure 6
anova_summary_data$HEIGHT[anova_summary_data$HEIGHT == "M1"] <- "0.5m"
anova_summary_data$HEIGHT[anova_summary_data$HEIGHT == "M2"] <- "1.2m"
anova_summary_data$LOAD[anova_summary_data$LOAD == "L1"] <- "15kg"
anova_summary_data$LOAD[anova_summary_data$LOAD == "L2"] <- "25kg"

fig_data <- filter(anova_summary_data, MUSCLE=="Iliocostalis")
plot1 <- ggplot(fig_data, aes(x=PHASE, y=mean, group=CONDITION)) +
  geom_errorbar(aes(ymin=mean-sd, ymax=mean+sd, color=CONDITION), width=.1) +
  geom_line(aes(color=CONDITION)) +
  geom_point(aes(color=CONDITION)) +
  facet_grid(HEIGHT ~ LOAD) +
  labs(title = "Iliocostalis",
       x = "Phases",
       y = "Normalised %MVC") +
  theme_classic() + 
  scale_color_manual(name="Condition",
                     labels=c("Exo", "Non-exo"),
                     values=c("blue", "red"))

fig_data <- filter(anova_summary_data, MUSCLE=="Multifidus")
plot2 <- ggplot(fig_data, aes(x=PHASE, y=mean, group=CONDITION)) +
  geom_errorbar(aes(ymin=mean-sd, ymax=mean+sd, color=CONDITION), width=.1) +
  geom_line(aes(color=CONDITION)) +
  geom_point(aes(color=CONDITION)) +
  facet_grid(HEIGHT ~ LOAD) +
  labs(title = "Multifidus",
       x = "Phases",
       y = "Normalised %MVC") +
  theme_classic() + 
  scale_color_manual(name="Condition",
                     labels=c("Exo", "Non-exo"),
                     values=c("blue", "red"))

fig_data <- filter(anova_summary_data, MUSCLE=="Longissimus")
plot3 <- ggplot(fig_data, aes(x=PHASE, y=mean, group=CONDITION)) +
  geom_errorbar(aes(ymin=mean-sd, ymax=mean+sd, color=CONDITION), width=.1) +
  geom_line(aes(color=CONDITION)) +
  geom_point(aes(color=CONDITION)) +
  facet_grid(HEIGHT ~ LOAD) +
  labs(title = "Longissimus",
       x = "Phases",
       y = "Normalised %MVC") +
  theme_classic() + 
  scale_color_manual(name="Condition",
                     labels=c("Exo", "Non-exo"),
                     values=c("blue", "red"))
plot_grid(plot1, plot2, plot3, labels = c("A", "B", "C"), nrow = 3)
# 700 x 1000 height

################## Plotting individual (Phase 2 only) ##################
#### Figure 7
data_avg_both_side_phase2 <- filter(data_avg_both_side, PHASE=="P2")
ggplot(data_avg_both_side_phase2, aes(x = CONDITION, y = both_value, color = MUSCLE, group = interaction(HEIGHT, LOAD, MUSCLE))) +
  geom_point(aes(shape = HEIGHT), size = 2.2) +   # Use different shapes for Left/Right sides
  scale_shape_manual(values = c(1, 17), name="Height", labels=c("0.5 m", "1.2 m")) +
  geom_line(aes(linetype = LOAD), size = 0.5) + # Connect points for Left/Right sides
  scale_linetype_manual(values = c("solid", "dashed"), name="Load", labels=c("15 kg", "25 kg")) +
  facet_wrap(~SUBJECT, scale="free", nrow=2, ncol=5) +                     # Facet by subject
  labs(title = "Normalised %MVC For Each Subject",
       x = "Condition",
       y = "Normalised %MVC",
       color = "Muscle") +
  theme_classic() +
  theme(legend.position = "right",
        strip.text = element_text(size = 12, color="black"), # Change the subject fontsize
        axis.text.x = element_text(size = 12, color="black"),
        axis.text.y = element_text(size = 12, color="black"),
        axis.title.x = element_text(size = 13, color = "black"),
        axis.title.y = element_text(size = 13, color = "black"),
        legend.text = element_text(size = 11)    # Legend text (labels) font size
  ) +
  scale_color_manual(values = c("Iliocostalis" = "red", "Multifidus" = "blue", "Longissimus" = "darkgreen"))
# 1200 x 600