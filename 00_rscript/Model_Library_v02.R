

list.of.packages <- c("rstudioapi","tidyr","dplyr","reshape2", "data.table",
                      "broom","car","ggplot2","tibble","xlsx","readxl",
                      "MASS", "forecast", "glmnet","apc","plotly","pscl"
)


#####################################
# Check if library packages are installed or not
#####################################
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
#####################################
# Attach library packages
#####################################
sapply(list.of.packages, library, character.only = TRUE)




# Set working directory to the source file
# Only works in RStudio
rstudioapi::getActiveDocumentContext()$path
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

