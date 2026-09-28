# =====================================================================
# generate_docx.ps1
# Builds Week1_Titanic_R_DataCleaning_Analysis.docx via Word COM automation
# Embeds curated R code, captured R output text, and PNG plots.
# =====================================================================

$ErrorActionPreference = "Stop"
$base   = "C:\Users\Adarsh\OneDrive\Documents\Default Project\week1-titanic-r"
$plots  = Join-Path $base "plots"
$outs   = Join-Path $base "outputs"
$outDoc = "C:\Users\Adarsh\OneDrive\Documents\Default Project\Week1_Titanic_R_DataCleaning_Analysis.docx"

# ---- WdBuiltinStyle constants (locale-independent) ----
$wdStyleTitle    = -63
$wdStyleSubtitle = -75
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdStyleNormal   = -1
$wdAlignCenter   = 1
$wdStory         = 6
$wdAlignLeft     = 0

# ---- Colors (RGB Long) ----
$colCode   = 8388608    # RGB(0,0,128)  dark blue
$colOutput = 5921280    # RGB(0,90,90)  teal
$colGray   = 15921906   # RGB(242,242,242) light gray (table header)
$colBlack  = 0

# =====================================================================
# START WORD
# =====================================================================
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$sel = $word.Selection

# Default body font
$normalStyle = $doc.Styles.Item($wdStyleNormal)
$normalStyle.Font.Name = "Calibri"
$normalStyle.Font.Size = 11

# Page margins (1 inch)
$doc.PageSetup.TopMargin    = $word.InchesToPoints(1)
$doc.PageSetup.BottomMargin = $word.InchesToPoints(1)
$doc.PageSetup.LeftMargin   = $word.InchesToPoints(1)
$doc.PageSetup.RightMargin  = $word.InchesToPoints(1)

# =====================================================================
# HELPER FUNCTIONS
# =====================================================================
function Set-Normal {
    $sel.Style = $doc.Styles.Item($wdStyleNormal)
    $sel.Font.Name = "Calibri"; $sel.Font.Size = 11; $sel.Font.Color = $colBlack
    $sel.ParagraphFormat.LeftIndent = 0
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.ParagraphFormat.SpaceBefore = 0
    $sel.ParagraphFormat.SpaceAfter  = 6
}

function Add-Title($t) {
    $sel.Style = $doc.Styles.Item($wdStyleTitle)
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-Subtitle($t) {
    $sel.Style = $doc.Styles.Item($wdStyleSubtitle)
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-CenterLine($t, $size) {
    Set-Normal
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.Font.Size = $size
    $sel.Font.Bold = $false
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-H1($t) {
    $sel.Style = $doc.Styles.Item($wdStyleHeading1)
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-H2($t) {
    $sel.Style = $doc.Styles.Item($wdStyleHeading2)
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-H3($t) {
    $sel.Style = $doc.Styles.Item($wdStyleHeading3)
    $sel.TypeText($t); $sel.TypeParagraph()
    Set-Normal
}

function Add-Body($t) {
    Set-Normal
    $sel.ParagraphFormat.Alignment = $wdAlignLeft
    $sel.ParagraphFormat.SpaceAfter = 8
    $sel.TypeText($t); $sel.TypeParagraph()
}

function Add-Bullet($t) {
    Set-Normal
    $sel.ParagraphFormat.LeftIndent = $word.InchesToPoints(0.3)
    $sel.TypeText([char]0x2022 + "  " + $t)
    $sel.TypeParagraph()
    Set-Normal
}

function Add-Code($text, $size) {
    if (-not $size) { $size = 9 }
    Set-Normal
    $sel.Font.Name = "Consolas"
    $sel.Font.Size = $size
    $sel.Font.Color = $colCode
    $sel.ParagraphFormat.LeftIndent = $word.InchesToPoints(0.2)
    $sel.ParagraphFormat.SpaceBefore = 3
    $sel.ParagraphFormat.SpaceAfter = 3
    foreach ($line in ($text -split "`r?`n")) {
        $sel.TypeText($line)
        $sel.TypeParagraph()
    }
    Set-Normal
}

function Add-OutputLabel {
    Set-Normal
    $sel.Font.Italic = $true
    $sel.Font.Size = 10
    $sel.Font.Color = $colBlack
    $sel.TypeText("R Output:")
    $sel.TypeParagraph()
    $sel.Font.Italic = $false
}

function Add-Output($text) {
    Set-Normal
    $sel.Font.Name = "Consolas"
    $sel.Font.Size = 8.5
    $sel.Font.Color = $colOutput
    $sel.ParagraphFormat.LeftIndent = $word.InchesToPoints(0.2)
    $sel.ParagraphFormat.SpaceBefore = 2
    $sel.ParagraphFormat.SpaceAfter = 6
    foreach ($line in ($text -split "`r?`n")) {
        $sel.TypeText($line)
        $sel.TypeParagraph()
    }
    Set-Normal
}

function Add-Image($path, $caption) {
    Set-Normal
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.ParagraphFormat.SpaceBefore = 6
    $shape = $sel.InlineShapes.AddPicture($path)
    if ($shape.Width -gt 446) { $shape.Width = 446 }
    $sel.TypeParagraph()
    if ($caption) {
        $sel.ParagraphFormat.Alignment = $wdAlignCenter
        $sel.Font.Italic = $true
        $sel.Font.Size = 9.5
        $sel.Font.Color = 5460976   # RGB(83,83,83) gray
        $sel.TypeText($caption)
        $sel.TypeParagraph()
    }
    Set-Normal
}

function Add-PageBreak {
    $sel.InsertBreak(7)   # wdPageBreak = 7
}

function Add-Table($headers, $rows) {
    Set-Normal
    $range = $sel.Range
    $ncols = $headers.Count
    $nrows = $rows.Count + 1
    $tbl = $doc.Tables.Add($range, $nrows, $ncols)
    try { $tbl.Borders.Enable = $true } catch {}
    for ($c = 0; $c -lt $ncols; $c++) {
        $cell = $tbl.Cell(1, $c + 1)
        $cell.Range.Text = $headers[$c]
        $cell.Range.Bold = $true
        try { $cell.Shading.BackgroundPatternColor = $colGray } catch {}
    }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $ncols; $c++) {
            $tbl.Cell($r + 2, $c + 1).Range.Text = $rows[$r][$c]
        }
    }
    $sel.EndKey($wdStory) | Out-Null
    $sel.TypeParagraph()
    Set-Normal
}

function Clean-Output($path) {
    $raw = Get-Content -LiteralPath $path -Raw
    $lines = $raw -split "`r?`n"
    $lines = @($lines | Where-Object { $_ -notmatch '^=+\s*$' -and $_ -notmatch '^SECTION\s' })
    $start = 0
    while ($start -lt $lines.Count -and [string]::IsNullOrWhiteSpace($lines[$start])) { $start++ }
    $end = $lines.Count - 1
    while ($end -gt $start -and [string]::IsNullOrWhiteSpace($lines[$end])) { $end-- }
    if ($start -gt $end) { return "" }
    ($lines[$start..$end]) -join "`n"
}

# =====================================================================
# TITLE PAGE
# =====================================================================
Add-Title "Data Cleaning and Preliminary Analysis with R"
Add-Subtitle "Week 1 Task Report"
$sel.TypeParagraph(); $sel.TypeParagraph()
Add-CenterLine "Dataset: Titanic Passenger Survival (Kaggle)" 12
Add-CenterLine "Tool: R 4.6.1 (base R) on Windows" 12
Add-CenterLine "Date: 28 September 2026" 12
Add-PageBreak

# =====================================================================
# TABLE OF CONTENTS
# =====================================================================
Add-H1 "Table of Contents"
$tocRange = $sel.Range
$toc = $doc.TablesOfContents.Add($tocRange)
Add-PageBreak

# =====================================================================
# 1. INTRODUCTION
# =====================================================================
Add-H1 "1. Introduction"
Add-Body "This report documents the Week 1 deliverable: data cleaning, preprocessing, and preliminary analysis of a publicly available dataset using R. The work follows a structured, reproducible pipeline - dataset selection, initial exploration, missing-value handling, outlier detection, normalization, categorical encoding, and exploratory data analysis (EDA). All analysis is executed in R 4.6.1 using base R only (no external packages), so the script runs in any standard R installation. The complete, reproducible script (titanic_analysis.R) is provided in Appendix A."

Add-H2 "1.1 Dataset Selection and Overview"
Add-Body "The Titanic passenger dataset was selected because it satisfies the task's complexity requirements: it contains a mixture of categorical variables (Sex, Embarked, Pclass, Survived) and numerical variables (Age, Fare, SibSp, Parch), and it exhibits meaningful missingness in Age (about 20 percent), Cabin (about 77 percent), and Embarked (about 0.2 percent). The dataset comprises 891 passenger records across 12 variables, with Survived (0 = died, 1 = survived) as the target variable. This combination of mixed types and missing values makes it ideal for demonstrating a complete data-cleaning workflow."

Add-H3 "Variable Dictionary"
$varHeaders = @("Variable","Type","Role","Description")
$varRows = @(
    @("PassengerId","integer","identifier","Unique passenger ID (dropped in cleaning)"),
    @("Survived","integer/factor","target","0 = died, 1 = survived"),
    @("Pclass","integer/factor","feature","Ticket class: 1 = 1st, 2 = 2nd, 3 = 3rd"),
    @("Name","character","identifier","Passenger name (dropped; title derivable later)"),
    @("Sex","character","feature","male / female"),
    @("Age","numeric","feature","Age in years (177 missing)"),
    @("SibSp","integer","feature","# siblings / spouses aboard"),
    @("Parch","integer","feature","# parents / children aboard"),
    @("Ticket","character","identifier","Ticket number (dropped)"),
    @("Fare","numeric","feature","Passenger fare (right-skewed)"),
    @("Cabin","character","feature","Cabin number (687 missing -> HasCabin flag)"),
    @("Embarked","character","feature","Port: C = Cherbourg, Q = Queenstown, S = Southampton")
)
Add-Table $varHeaders $varRows

Add-H2 "1.2 Environment and Reproducibility"
Add-Body "R version 4.6.1 was used on Windows. Only base R functions are used (read.csv, str, summary, ave, quantile, cor, table, prop.table, and base graphics), guaranteeing the script runs without installing additional packages. Console output for each analysis stage is captured to text files in the outputs/ folder, and every plot is saved as a 110-DPI PNG in the plots/ folder, so all evidence embedded in this report is generated directly by the script."

# =====================================================================
# 2. INITIAL EXPLORATION
# =====================================================================
Add-H1 "2. Initial Data Exploration"
Add-Body "The raw CSV is loaded with blank strings coerced to NA via na.strings, so that genuine missingness is immediately detectable. The functions dim(), names(), str(), head(), tail(), and summary() give the first comprehensive picture of the data's structure, variable types, and distributions."

$code1 = @'
titanic <- read.csv("data/titanic.csv",
                    stringsAsFactors = FALSE,
                    na.strings = c("", "NA"))

dim(titanic)        # 891 rows x 12 columns
names(titanic)
str(titanic)
head(titanic)
tail(titanic)
summary(titanic)

# Variable type breakdown
table(sapply(titanic, class))
'@
Add-Code $code1
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "01_exploration.txt"))

Add-Body "summary() immediately reveals 177 missing values in Age and 687 in Cabin, plus a maximum Fare of 512.33 (a likely outlier). The type breakdown shows 5 character, 5 integer, and 2 numeric columns, confirming the categorical/numerical mix required by the task."

# =====================================================================
# 3. MISSING VALUE ANALYSIS
# =====================================================================
Add-H1 "3. Missing Value Analysis"
Add-Body "Missingness is quantified per variable and as an overall rate. A clear picture of where and how much data is missing drives the imputation strategy."

$code2 = @'
missing_count <- colSums(is.na(titanic))
missing_pct   <- round(missing_count / nrow(titanic) * 100, 2)

data.frame(Variable = names(titanic),
           Missing  = missing_count,
           Percent  = missing_pct)

sum(missing_count)                # total missing cells
sum(!complete.cases(titanic))     # rows with >= 1 NA

barplot(missing_count, las = 2, col = "firebrick",
        main = "Missing Values per Variable",
        ylab = "Count of NA")
'@
Add-Code $code2
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "02_missing.txt"))
Add-Image (Join-Path $plots "02_missing.png") "Figure 1. Missing value counts per variable in the raw Titanic dataset."

Add-Body "Cabin has 77 percent missing - far too much for reliable imputation - so it is converted to a binary HasCabin flag and the raw column is dropped. Age (about 20 percent missing) is imputed with the median within each Pclass x Sex group, which preserves demographic structure far better than a single global median. Embarked (only 2 missing cases) is imputed with the mode, 'S' (Southampton)."

# =====================================================================
# 4. DATA CLEANING
# =====================================================================
Add-H1 "4. Data Cleaning Process"
Add-Body "Cleaning is performed in a logical sequence: trim stray whitespace, drop identifier columns that carry no analytic signal, convert coded integers to labelled factors, impute the remaining missing values, engineer a HasCabin flag from the sparse Cabin column, and finally check for true duplicate records on the identifier-inclusive raw data."

$code3 = @'
clean <- titanic

# 4.1 trim whitespace in character columns
clean[] <- lapply(clean, function(x)
  if (is.character(x)) trimws(x) else x)

# 4.2 drop identifier columns
clean <- clean[, !(names(clean) %in%
                   c("PassengerId", "Name", "Ticket"))]

# 4.3 coded integers -> labelled factors
clean$Survived <- factor(clean$Survived, levels = c(0, 1),
                         labels = c("Died", "Survived"))
clean$Pclass   <- factor(clean$Pclass, levels = c(1, 2, 3),
                         labels = c("Class1", "Class2", "Class3"))

# 4.4 Embarked: mode imputation
emb_mode <- names(sort(table(clean$Embarked), decreasing = TRUE))[1]
clean$Embarked[is.na(clean$Embarked)] <- emb_mode

# 4.5 Cabin -> binary HasCabin flag, drop Cabin
clean$HasCabin <- factor(ifelse(is.na(clean$Cabin), "No", "Yes"))
clean <- clean[, names(clean) != "Cabin"]

# 4.6 Age: group-wise (Pclass:Sex) median imputation
clean$Age <- ave(clean$Age, clean$Pclass, clean$Sex,
  FUN = function(x) ifelse(is.na(x), median(x, na.rm = TRUE), x))

# 4.7 true duplicates checked on the RAW data (with PassengerId)
dup_count <- sum(duplicated(titanic))
clean <- clean[!duplicated(titanic), ]

str(clean); summary(clean); colSums(is.na(clean))
'@
Add-Code $code3
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "03_cleaning.txt"))

Add-Body "After cleaning, every column has zero missing values and the dataset remains 891 rows by 9 columns (no true duplicates existed). Survived and Pclass are now labelled factors, ensuring R treats them as categorical in later summaries and models. Note that duplicate detection was deliberately performed on the raw, identifier-inclusive data: had it been run only after dropping PassengerId/Name/Ticket, many distinct passengers sharing the remaining attributes would have been wrongly flagged as duplicates and discarded."

# =====================================================================
# 5. OUTLIER DETECTION
# =====================================================================
Add-H1 "5. Outlier Detection"
Add-Body "The interquartile-range (IQR) rule flags any value below Q1 - 1.5*IQR or above Q3 + 1.5*IQR as an outlier. This is applied to all four numeric variables; the decision to keep, cap, or remove each set of outliers is justified individually."

$code4 = @'
detect_outliers <- function(x) {
  qs   <- quantile(x, c(0.25, 0.75), na.rm = TRUE)
  iqr  <- qs[2] - qs[1]
  low  <- qs[1] - 1.5 * iqr
  high <- qs[2] + 1.5 * iqr
  which(x < low | x > high)
}

for (v in c("Age", "Fare", "SibSp", "Parch"))
  cat(sprintf("%-6s outliers: %d (%.2f%%)\n",
              v, length(detect_outliers(clean[[v]])),
              length(detect_outliers(clean[[v]])) / nrow(clean) * 100))

# Fare is right-skewed: winsorize at the 99th percentile
fare_cap <- quantile(clean$Fare, 0.99)
clean$Fare[clean$Fare > fare_cap] <- fare_cap

boxplot(clean$Age, clean$Fare)
'@
Add-Code $code4
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "04_outliers.txt"))
Add-Image (Join-Path $plots "04_outliers.png") "Figure 2. Boxplots of Age and Fare after cleaning (Fare shown after winsorizing)."

Add-Body "Age, SibSp, and Parch outliers are retained because they represent plausible real values (elderly passengers and large families). Fare is right-skewed with 13 percent outliers and a maximum of 512.33; it is winsorized at the 99th percentile (cap about 249.01), reducing the influence of extreme tickets while preserving the bulk of the distribution. Winsorizing is preferred over deletion here because the highest fares belong disproportionately to Class 1 survivors, and deleting them would bias the data."

# =====================================================================
# 6. NORMALIZATION / SCALING
# =====================================================================
Add-H1 "6. Normalization and Feature Scaling"
Add-Body "Two complementary scaling methods are applied to Age and to the winsorized Fare. Both scaled versions are retained so downstream modelling can choose the appropriate one."

$code5 = @'
minmax <- function(x) (x - min(x)) / (max(x) - min(x))   # range [0,1]
zscore <- function(x) (x - mean(x)) / sd(x)             # mean 0, sd 1

clean$Age_minmax  <- minmax(clean$Age)
clean$Fare_minmax <- minmax(clean$Fare)
clean$Age_z       <- zscore(clean$Age)
clean$Fare_z      <- zscore(clean$Fare)
'@
Add-Code $code5
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "05_normalization.txt"))
Add-Image (Join-Path $plots "05_normalization.png") "Figure 3. Distribution of Age and Fare after Min-Max normalization."

Add-Body "Min-Max normalization preserves the shape of the original distribution within the bounded range [0, 1], which suits distance- and gradient-based algorithms. Z-score (standard) scaling centers each variable at mean 0 with standard deviation 1, handling features with large spread better. Because Fare is winsorized before scaling, its z-score maximum (about 5.12) is far less extreme than it would have been from the raw maximum of 512.33."

# =====================================================================
# 7. ENCODING CATEGORICAL VARIABLES
# =====================================================================
Add-H1 "7. Encoding Categorical Variables"
Add-Body "Binary categoricals are label-encoded; nominal categoricals with more than two levels are one-hot encoded to avoid imposing a false ordinal relationship among their categories."

$code6 = @'
# 7.1 binary label encoding
clean$Sex_code     <- ifelse(clean$Sex == "male", 1, 0)
clean$HasCabin_code <- ifelse(clean$HasCabin == "Yes", 1, 0)

# 7.2 one-hot encoding for Embarked (3 levels)
for (lv in levels(factor(clean$Embarked)))
  clean[[paste0("Embarked_", lv)]] <- as.integer(clean$Embarked == lv)

# 7.3 one-hot encoding for Pclass (3 levels)
for (lv in levels(clean$Pclass))
  clean[[paste0("Pclass_", lv)]] <- as.integer(clean$Pclass == lv)

head(clean[, c("Sex_code", "Embarked_C", "Embarked_Q", "Embarked_S",
              "Pclass_Class1", "Pclass_Class2", "Pclass_Class3",
              "HasCabin_code")])
'@
Add-Code $code6
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "06_encoding.txt"))

Add-Body "Sex becomes Sex_code (male = 1, female = 0); Embarked expands to three 0/1 dummy columns (Embarked_C, Embarked_Q, Embarked_S); Pclass expands to three dummy columns; and HasCabin becomes HasCabin_code. After this step the transformed data frame is fully numeric-ready for distance- and matrix-based modelling algorithms."

# =====================================================================
# 8. EXPLORATORY DATA ANALYSIS
# =====================================================================
Add-H1 "8. Exploratory Data Analysis"
Add-Body "With a clean, fully typed, and encoded dataset, exploratory analysis examines survival across key categorical groupings and the correlations among numeric/encoded features."

$code7 = @'
# 8.1 overall and grouped survival rates
table(clean$Survived)
table(clean$Sex, clean$Survived)
prop.table(table(clean$Pclass, clean$Survived), 1) * 100
prop.table(table(clean$Embarked, clean$Survived), 1) * 100

# 8.2 correlation among numeric / encoded features
cor_vars <- clean[, c("Age", "SibSp", "Parch", "Fare", "Sex_code",
                      "Pclass_Class1", "Pclass_Class3", "HasCabin_code")]
cor_mat <- round(cor(cor_vars), 3)
print(cor_mat)

# 8.3 correlation of each feature with Survived (encoded 0/1)
surv_num <- as.integer(clean$Survived) - 1
sort(cor(surv_num, cor_vars), decreasing = TRUE)
'@
Add-Code $code7
Add-OutputLabel
Add-Output (Clean-Output (Join-Path $outs "07_eda.txt"))

Add-Image (Join-Path $plots "07_correlation.png") "Figure 4. Correlation matrix of numeric and encoded variables."
Add-Image (Join-Path $plots "08_histograms.png") "Figure 5. Age and Fare distributions coloured by survival outcome."
Add-Image (Join-Path $plots "09_barplots.png") "Figure 6. Survival counts by Sex and by Passenger Class."
Add-Image (Join-Path $plots "10_age_fare_scatter.png") "Figure 7. Age vs Fare scatter, coloured by survival outcome."

Add-H2 "8.1 Key Findings"
Add-Bullet "Overall survival rate is 38.4 percent (342 of 891)."
Add-Bullet "Sex is the strongest predictor: 74.2 percent of females survived versus 18.9 percent of males; the correlation of Sex_code with Survived is -0.54."
Add-Bullet "A clear class gradient exists: Class 1 63.0 percent, Class 2 47.3 percent, Class 3 24.2 percent survival."
Add-Bullet "Passengers embarking at Cherbourg (C) had the highest survival rate (55.4 percent), largely because a higher proportion of them travelled in Class 1."
Add-Bullet "Fare is positively correlated with survival (0.27); Age is only weakly correlated (-0.06)."
Add-Bullet "HasCabin (a proxy for Class 1/2) correlates 0.32 with survival."

# =====================================================================
# 9. CONCLUSIONS AND NEXT STEPS
# =====================================================================
Add-H1 "9. Conclusions and Next Steps"
Add-Body "The cleaning pipeline transformed the raw Titanic data into a tidy, fully typed, scaled, and encoded data frame of 891 rows, saved as titanic_cleaned.csv and ready for predictive modelling in subsequent weeks. Sex, Pclass, and Fare emerge as the dominant predictors of survival; Age and family size are weak individually but valuable once engineered. Planned next steps include: deriving a FamilySize feature from SibSp + Parch and an isAlone flag; extracting a Title feature from passenger names; performing a stratified train/test split; and fitting baseline models (logistic regression and decision trees) to quantify the predictive value of the cleaned and engineered features."

# =====================================================================
# APPENDIX A: FULL R SCRIPT
# =====================================================================
Add-H1 "Appendix A: Complete R Script (titanic_analysis.R)"
$rscript = Get-Content -LiteralPath (Join-Path $base "titanic_analysis.R") -Raw
Add-Code $rscript 8

# =====================================================================
# APPENDIX B: FILE MANIFEST
# =====================================================================
Add-H1 "Appendix B: File Manifest"
Add-Bullet "titanic_analysis.R - reproducible R analysis script"
Add-Bullet "data/titanic.csv - raw Titanic dataset (891 x 12)"
Add-Bullet "data/titanic_cleaned.csv - cleaned, scaled and encoded dataset"
Add-Bullet "outputs/01_exploration.txt ... 07_eda.txt - captured R console output"
Add-Bullet "plots/02_missing.png ... 10_age_fare_scatter.png - generated visualizations"

# =====================================================================
# UPDATE TOC AND SAVE
# =====================================================================
try { $doc.TablesOfContents.Item(1).Update() } catch {}
foreach ($f in $doc.Fields) { try { $f.Update() | Out-Null } catch {} }

# Set document core properties
try {
    $doc.BuiltInDocumentProperties.Item("Title").Value  = "Data Cleaning and Preliminary Analysis with R"
    $doc.BuiltInDocumentProperties.Item("Author").Value = "Week 1 Task"
    $doc.BuiltInDocumentProperties.Item("Subject").Value = "Titanic Dataset - R Data Cleaning and EDA"
} catch {}

if (Test-Path $outDoc) { Remove-Item -LiteralPath $outDoc -Force }
$doc.SaveAs([ref]$outDoc, [ref]16)   # wdFormatDocumentDefault = 16 (.docx)
$doc.Close()
$word.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Output "SUCCESS: Document saved to $outDoc"
Write-Output ("Size: " + (Get-Item $outDoc).Length + " bytes")
