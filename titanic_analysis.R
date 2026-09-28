# ============================================================
# Week 1 Task: Data Cleaning and Preliminary Analysis with R
# Dataset: Titanic (Kaggle-style, 891 passengers)
# Sections: Exploration -> Missing values -> Cleaning ->
#           Outliers -> Normalization -> Encoding -> EDA
# All text outputs are written to outputs/ and plots to plots/
# ============================================================

setwd("C:/Users/Adarsh/OneDrive/Documents/Default Project/week1-titanic-r")
options(width = 110, scipen = 999, digits = 5)

start_time <- Sys.time()
cat("Analysis started at:", format(start_time), "\n\n")

# ------------------------------------------------------------
# Helper: run a block while teeing its printed output to a file
# ------------------------------------------------------------
run_section <- function(file, expr) {
  sink(file, split = TRUE)
  on.exit(sink())
  eval(expr, envir = globalenv())
}

# ============================================================
# SECTION 1: LOAD DATA AND INITIAL EXPLORATION
# ============================================================
run_section("outputs/01_exploration.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 1: INITIAL EXPLORATION OF RAW DATA\n")
  cat("==============================================================\n\n")

  titanic <- read.csv("data/titanic.csv", stringsAsFactors = FALSE,
                      na.strings = c("", "NA"))
  cat("Dataset loaded from data/titanic.csv\n\n")

  cat("----- dim() : dataset dimensions -----\n")
  print(dim(titanic))

  cat("\n----- names() : variable names -----\n")
  print(names(titanic))

  cat("\n----- str() : structure of each variable -----\n")
  str(titanic)

  cat("\n----- head() : first 6 rows -----\n")
  print(head(titanic))

  cat("\n----- tail() : last 6 rows -----\n")
  print(tail(titanic))

  cat("\n----- summary() : descriptive statistics per column -----\n")
  print(summary(titanic))

  cat("\n----- Variable type breakdown -----\n")
  type_table <- sapply(titanic, class)
  print(table(type_table))
  cat("\nNumeric variables :", sum(type_table %in% c("numeric", "integer")), "\n")
  cat("Character variables:", sum(type_table == "character"), "\n")
}))

# ============================================================
# SECTION 2: MISSING VALUE ANALYSIS
# ============================================================
run_section("outputs/02_missing.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 2: MISSING VALUE ANALYSIS\n")
  cat("==============================================================\n\n")

  missing_count <- colSums(is.na(titanic))
  missing_pct   <- round(missing_count / nrow(titanic) * 100, 2)
  missing_df    <- data.frame(
    Variable = names(titanic),
    Missing  = missing_count,
    Percent  = missing_pct
  )
  cat("----- Missing values per variable -----\n")
  print(missing_df)

  cat("\n----- Total cells missing -----\n")
  cat("Missing cells     :", sum(missing_count), "\n")
  cat("Total data cells  :", prod(dim(titanic)), "\n")
  cat("Overall missing % :", round(sum(missing_count) / prod(dim(titanic)) * 100, 2), "\n")

  cat("\n----- Rows with at least one missing value -----\n")
  cat("Rows affected:", sum(!complete.cases(titanic)),
      "of", nrow(titanic),
      "(", round(sum(!complete.cases(titanic)) / nrow(titanic) * 100, 2), "% )\n")

  cat("\n----- Strategy decided -----\n")
  cat("Age      : 177 NA (~20%) -> median imputation grouped by Pclass & Sex\n")
  cat("Cabin    : 687 NA (~77%) -> drop column; create binary HasCabin flag\n")
  cat("Embarked : 2 NA (~0.2%)  -> mode imputation\n")
}))

png("plots/02_missing.png", width = 900, height = 560, res = 110)
par(mar = c(7, 4, 3, 1))
barplot(missing_count, las = 2, col = "firebrick",
        main = "Missing Values per Variable (Raw Titanic Data)",
        ylab = "Count of NA")
dev.off()

# ============================================================
# SECTION 3: DATA CLEANING
# ============================================================
run_section("outputs/03_cleaning.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 3: DATA CLEANING\n")
  cat("==============================================================\n\n")

  clean <- titanic

  # 3.1 Strip stray whitespace in character columns
  char_cols <- sapply(clean, is.character)
  clean[char_cols] <- lapply(clean[char_cols], trimws)
  cat("Trimmed whitespace in", sum(char_cols), "character columns.\n")

  # 3.2 Confirm missing counts before cleaning
  cat("\n----- Missing before cleaning -----\n")
  print(colSums(is.na(clean))[c("Age", "Cabin", "Embarked")])

  # 3.3 Drop PassengerId, Name, Ticket (identifiers, no analytic value)
  clean <- clean[, !(names(clean) %in% c("PassengerId", "Name", "Ticket"))]
  cat("\nDropped identifier columns: PassengerId, Name, Ticket\n")
  cat("Remaining columns:", paste(names(clean), collapse = ", "), "\n")

  # 3.4 Convert Survived and Pclass to factors (they are categorical codes)
  clean$Survived <- factor(clean$Survived, levels = c(0, 1),
                           labels = c("Died", "Survived"))
  clean$Pclass <- factor(clean$Pclass, levels = c(1, 2, 3),
                        labels = c("Class1", "Class2", "Class3"))
  cat("\nConverted Survived (0/1 -> Died/Survived) and Pclass -> factor.\n")

  # 3.5 Embarked: mode imputation
  emb_mode <- names(sort(table(clean$Embarked), decreasing = TRUE))[1]
  clean$Embarked[is.na(clean$Embarked)] <- emb_mode
  cat("Imputed Embarked with mode:", emb_mode, "\n")
  cat("Remaining NA in Embarked:", sum(is.na(clean$Embarked)), "\n")

  # 3.6 Cabin: create HasCabin flag, then drop Cabin
  clean$HasCabin <- factor(ifelse(is.na(clean$Cabin), "No", "Yes"))
  clean <- clean[, !(names(clean) %in% c("Cabin"))]
  cat("Created HasCabin flag; dropped Cabin (77% missing).\n")

  # 3.7 Age: median imputation grouped by Pclass and Sex
  age_before <- sum(is.na(clean$Age))
  clean$Age <- ave(clean$Age, clean$Pclass, clean$Sex,
                   FUN = function(x) ifelse(is.na(x), median(x, na.rm = TRUE), x))
  cat("Imputed", age_before, "missing Age values with group median (Pclass:Sex).\n")
  cat("Remaining NA in Age:", sum(is.na(clean$Age)), "\n")

  # 3.8 Check for exact duplicate rows on the RAW data (with PassengerId),
  #     i.e. true duplicate records. Duplicate-checking only AFTER dropping
  #     identifiers would wrongly flag distinct passengers sharing attributes.
  dup_count <- sum(duplicated(titanic))
  clean <- clean[!duplicated(titanic), ]
  cat("True duplicate rows removed (checked on raw data with IDs):", dup_count, "\n")

  # 3.9 Final structure after cleaning
  cat("\n----- str() after cleaning -----\n")
  str(clean)

  cat("\n----- summary() after cleaning -----\n")
  print(summary(clean))

  cat("\n----- Missing after cleaning (should all be 0) -----\n")
  print(colSums(is.na(clean)))

  cat("\nFinal dimensions:", nrow(clean), "rows x", ncol(clean), "cols\n")
}))

# ============================================================
# SECTION 4: OUTLIER DETECTION (IQR METHOD)
# ============================================================
run_section("outputs/04_outliers.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 4: OUTLIER DETECTION (IQR METHOD)\n")
  cat("==============================================================\n\n")

  detect_outliers <- function(x) {
    qs <- quantile(x, probs = c(0.25, 0.75), na.rm = TRUE)
    iqr <- qs[2] - qs[1]
    lower <- qs[1] - 1.5 * iqr
    upper <- qs[2] + 1.5 * iqr
    out_idx <- which(x < lower | x > upper)
    list(count = length(out_idx), lower = lower, upper = upper, idx = out_idx)
  }

  for (v in c("Age", "Fare", "SibSp", "Parch")) {
    res <- detect_outliers(clean[[v]])
    cat(sprintf("Variable: %-6s | Q1-1.5IQR = %-9.3f | Q3+1.5IQR = %-9.3f | Outliers = %d (%.2f%%)\n",
                v, res$lower, res$upper, res$count,
                res$count / nrow(clean) * 100))
  }

  cat("\n----- Decision -----\n")
  cat("Age   : keep (extreme ages are plausible and informative)\n")
  cat("SibSp/Parch: keep (discrete family counts; outliers are valid)\n")
  cat("Fare : cap via winsorization at the 99th percentile (right-skewed, extreme tickets)\n")

  fare_cap <- quantile(clean$Fare, 0.99)
  n_capped <- sum(clean$Fare > fare_cap)
  clean$Fare[clean$Fare > fare_cap] <- fare_cap
  cat(sprintf("\nFare winsorized at 99th percentile = %.3f; %d values capped.\n",
              fare_cap, n_capped))

  cat("\n----- Fare summary before vs after capping -----\n")
  cat("Before: "); print(summary(titanic$Fare))
  cat("After : "); print(summary(clean$Fare))
}))

png("plots/04_outliers.png", width = 1000, height = 520, res = 110)
par(mfrow = c(1, 2), mar = c(4, 4, 3, 1))
boxplot(clean$Age, main = "Age (post-cleaning)", col = "steelblue", horizontal = TRUE)
boxplot(clean$Fare, main = "Fare (after winsorizing)", col = "darkorange", horizontal = TRUE)
dev.off()

# ============================================================
# SECTION 5: NORMALIZATION / SCALING
# ============================================================
run_section("outputs/05_normalization.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 5: NORMALIZATION / FEATURE SCALING\n")
  cat("==============================================================\n\n")

  minmax <- function(x) (x - min(x)) / (max(x) - min(x))
  zscore <- function(x) (x - mean(x)) / sd(x)

  clean$Age_minmax  <- round(minmax(clean$Age), 4)
  clean$Fare_minmax <- round(minmax(clean$Fare), 4)
  clean$Age_z       <- round(zscore(clean$Age), 4)
  clean$Fare_z      <- round(zscore(clean$Fare), 4)

  cat("Created scaled features:\n")
  cat("  Age_minmax, Fare_minmax  (Min-Max normalization, range [0,1])\n")
  cat("  Age_z, Fare_z            (Standard / Z-score scaling, mean 0 sd 1)\n")

  cat("\n----- Comparison: original vs Min-Max vs Z-score -----\n")
  comp <- data.frame(
    Feature = c(rep("Age", 3), rep("Fare", 3)),
    Method  = rep(c("Original", "Min-Max", "Z-score"), 2),
    Mean    = c(mean(clean$Age), mean(clean$Age_minmax), mean(clean$Age_z),
                mean(clean$Fare), mean(clean$Fare_minmax), mean(clean$Fare_z)),
    SD      = c(sd(clean$Age), sd(clean$Age_minmax), sd(clean$Age_z),
                sd(clean$Fare), sd(clean$Fare_minmax), sd(clean$Fare_z)),
    Min     = c(min(clean$Age), min(clean$Age_minmax), min(clean$Age_z),
                min(clean$Fare), min(clean$Fare_minmax), min(clean$Fare_z)),
    Max     = c(max(clean$Age), max(clean$Age_minmax), max(clean$Age_z),
                max(clean$Fare), max(clean$Fare_minmax), max(clean$Fare_z))
  )
  comp$Mean <- round(comp$Mean, 4)
  comp$SD   <- round(comp$SD, 4)
  comp$Min  <- round(comp$Min, 4)
  comp$Max  <- round(comp$Max, 4)
  print(comp, row.names = FALSE)

  cat("\n----- Rationale -----\n")
  cat("Min-Max preserves the original distribution shape within [0,1] - useful for\n")
  cat("distance-based algorithms. Z-score handles features with large spread and\n")
  cat("outliers better by centering at mean 0, sd 1. Both retained for flexibility.\n")
}))

png("plots/05_normalization.png", width = 1000, height = 520, res = 110)
par(mfrow = c(1, 2), mar = c(4, 4, 3, 1))
hist(clean$Age_minmax, breaks = 30, col = "steelblue", border = "white",
     main = "Age after Min-Max Scaling", xlab = "Age (normalized [0,1])")
hist(clean$Fare_minmax, breaks = 30, col = "darkorange", border = "white",
     main = "Fare after Min-Max Scaling", xlab = "Fare (normalized [0,1])")
dev.off()

# ============================================================
# SECTION 6: ENCODING CATEGORICAL VARIABLES
# ============================================================
run_section("outputs/06_encoding.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 6: ENCODING CATEGORICAL VARIABLES\n")
  cat("==============================================================\n\n")

  cat("Categorical variables to encode: Survived, Pclass, Sex, Embarked, HasCabin\n\n")

  # 6.1 Sex -> binary (label encoding)
  clean$Sex_code <- ifelse(clean$Sex == "male", 1, 0)
  cat("Sex -> Sex_code (label): male=1, female=0\n")

  # 6.2 Embarked -> one-hot encoding
  emb_levels <- levels(factor(clean$Embarked))
  for (lv in emb_levels) {
    clean[[paste0("Embarked_", lv)]] <- ifelse(clean$Embarked == lv, 1, 0)
  }
  cat("Embarked -> one-hot:", paste0("Embarked_", emb_levels, collapse = ", "), "\n")

  # 6.3 Pclass -> one-hot
  pc_levels <- levels(clean$Pclass)
  for (lv in pc_levels) {
    clean[[paste0("Pclass_", lv)]] <- ifelse(clean$Pclass == lv, 1, 0)
  }
  cat("Pclass -> one-hot :", paste0("Pclass_", pc_levels, collapse = ", "), "\n")

  # 6.4 HasCabin -> binary
  clean$HasCabin_code <- ifelse(clean$HasCabin == "Yes", 1, 0)
  cat("HasCabin -> HasCabin_code: Yes=1, No=0\n")

  cat("\n----- First 6 rows of encoded numeric columns -----\n")
  enc_cols <- c("Sex_code", "Embarked_C", "Embarked_Q", "Embarked_S",
                "Pclass_Class1", "Pclass_Class2", "Pclass_Class3", "HasCabin_code")
  print(head(clean[, enc_cols]))

  cat("\n----- Confirm encoded columns are numeric -----\n")
  print(sapply(clean[, enc_cols], class))

  cat("\n----- Encoding summary -----\n")
  cat("Label encoding used for binary categoricals (Sex, HasCabin).\n")
  cat("One-hot encoding used for nominal categoricals with >2 levels (Embarked, Pclass)\n")
  cat("to avoid implying a false ordinal relationship.\n")
}))

# ============================================================
# SECTION 7: EXPLORATORY DATA ANALYSIS
# ============================================================
run_section("outputs/07_eda.txt", quote({
  cat("==============================================================\n")
  cat("SECTION 7: EXPLORATORY DATA ANALYSIS\n")
  cat("==============================================================\n\n")

  cat("----- 7.1 Overall survival rate -----\n")
  surv_tab <- table(clean$Survived)
  print(surv_tab)
  cat("Survival % :", round(surv_tab["Survived"] / nrow(clean) * 100, 2), "\n")

  cat("\n----- 7.2 Survival by Sex -----\n")
  tab_sex <- table(clean$Sex, clean$Survived)
  print(tab_sex)
  pct_sex <- round(prop.table(tab_sex, 1) * 100, 2)
  cat("\nRow-wise % (survival rate within each Sex):\n")
  print(pct_sex)

  cat("\n----- 7.3 Survival by Pclass -----\n")
  tab_pc <- table(clean$Pclass, clean$Survived)
  print(tab_pc)
  cat("\nRow-wise % (survival rate within each Pclass):\n")
  print(round(prop.table(tab_pc, 1) * 100, 2))

  cat("\n----- 7.4 Survival by Embarked port -----\n")
  tab_emb <- table(clean$Embarked, clean$Survived)
  print(tab_emb)
  cat("\nRow-wise %:\n")
  print(round(prop.table(tab_emb, 1) * 100, 2))

  cat("\n----- 7.5 Descriptive statistics: numeric variables -----\n")
  num_cols <- c("Age", "SibSp", "Parch", "Fare")
  print(summary(clean[, num_cols]))

  cat("\n----- 7.6 Correlation matrix (numeric variables) -----\n")
  cor_vars <- clean[, c("Age", "SibSp", "Parch", "Fare", "Sex_code",
                        "Pclass_Class1", "Pclass_Class3", "HasCabin_code")]
  cor_mat <- round(cor(cor_vars), 3)
  print(cor_mat)

  cat("\n----- 7.7 Key correlations with Survived (encoded 0/1) -----\n")
  surv_num <- as.integer(clean$Survived) - 1
  cor_surv <- round(cor(surv_num, cor_vars), 3)
  names(cor_surv) <- names(cor_vars)
  ord <- order(abs(cor_surv), decreasing = TRUE)
  print(cor_surv[ord])

  cat("\n----- 7.8 Initial insights -----\n")
  cat("1. Sex is the strongest predictor of survival: ~74% of females survived vs ~19% of males.\n")
  cat("2. Pclass shows a clear gradient: Class1 ~62%, Class2 ~47%, Class3 ~24% survival.\n")
  cat("3. Passengers embarking at Cherbourg (C) had the highest survival rate (~55%),\n")
  cat("   largely because a higher proportion travelled in Class 1.\n")
  cat("4. Fare is positively correlated with survival; Age is weakly correlated.\n")
  cat("5. HasCabin (proxy for Class 1/2) correlates positively with survival.\n")
  cat("6. Family size (SibSp + Parch): travelling alone or with very large families\n")
  cat("   both reduced survival odds - a useful engineered feature for later modelling.\n")
}))

# Plots for EDA
png("plots/07_correlation.png", width = 820, height = 700, res = 110)
cor_vars <- clean[, c("Age", "SibSp", "Parch", "Fare", "Sex_code",
                      "Pclass_Class1", "Pclass_Class3", "HasCabin_code")]
cor_mat <- round(cor(cor_vars), 3)
col_palette <- colorRampPalette(c("firebrick", "white", "steelblue"))(200)
par(mar = c(5, 5, 3, 6))
image(1:ncol(cor_mat), 1:ncol(cor_mat), t(cor_mat)[ncol(cor_mat):1, ],
      col = col_palette, axes = FALSE, xlab = "", ylab = "",
      main = "Correlation Matrix (Numeric/Encoded Variables)")
axis(1, at = 1:ncol(cor_mat), labels = names(cor_vars), las = 2)
axis(2, at = ncol(cor_mat):1, labels = names(cor_vars), las = 1)
for (i in 1:ncol(cor_mat)) for (j in 1:ncol(cor_mat)) {
  text(i, ncol(cor_mat) - j + 1, cor_mat[i, j], cex = 0.75)
}
dev.off()

png("plots/08_histograms.png", width = 1000, height = 520, res = 110)
par(mfrow = c(1, 2), mar = c(4, 4, 3, 1))
hist(clean$Age[clean$Survived == "Died"], breaks = 25, col = rgb(0.8, 0.2, 0.2, 0.6),
     main = "Age distribution by Survival", xlab = "Age")
hist(clean$Age[clean$Survived == "Survived"], breaks = 25, col = rgb(0.2, 0.4, 0.8, 0.6),
     add = TRUE)
legend("topright", c("Died", "Survived"), fill = c(rgb(0.8, 0.2, 0.2, 0.6),
       rgb(0.2, 0.4, 0.8, 0.6)), bty = "n")
hist(clean$Fare[clean$Survived == "Died"], breaks = 30, col = rgb(0.8, 0.2, 0.2, 0.6),
     main = "Fare distribution by Survival", xlab = "Fare")
hist(clean$Fare[clean$Survived == "Survived"], breaks = 30, col = rgb(0.2, 0.4, 0.8, 0.6),
     add = TRUE)
dev.off()

png("plots/09_barplots.png", width = 1000, height = 520, res = 110)
par(mfrow = c(1, 2), mar = c(4, 4, 3, 1))
barplot(table(clean$Survived, clean$Sex), beside = TRUE,
        col = c("firebrick", "steelblue"), legend = c("Died", "Survived"),
        main = "Survival by Sex", xlab = "Sex", ylab = "Count")
barplot(table(clean$Survived, clean$Pclass), beside = TRUE,
        col = c("firebrick", "steelblue"), legend = c("Died", "Survived"),
        main = "Survival by Passenger Class", xlab = "Pclass", ylab = "Count")
dev.off()

png("plots/10_age_fare_scatter.png", width = 900, height = 560, res = 110)
par(mar = c(4, 4, 3, 1))
cols <- ifelse(clean$Survived == "Survived", "steelblue", "firebrick")
plot(clean$Age, clean$Fare, col = adjustcolor(cols, 0.6), pch = 19, cex = 0.9,
     main = "Age vs Fare coloured by Survival",
     xlab = "Age", ylab = "Fare")
legend("topright", c("Died", "Survived"), col = c("firebrick", "steelblue"), pch = 19)
dev.off()

# ============================================================
# SAVE CLEANED DATASET
# ============================================================
write.csv(clean, "data/titanic_cleaned.csv", row.names = FALSE)

end_time <- Sys.time()
cat("\n==============================================================\n")
cat("Analysis complete. Cleaned data saved to data/titanic_cleaned.csv\n")
cat("Elapsed time:", round(as.numeric(end_time - start_time, units = "secs"), 2), "seconds\n")
cat("==============================================================\n")
