library(energy)
library(ggplot2)
library(readxl)
library(dplyr)
library(knitr)   
library(openxlsx)  

# Define file path
#----------------------------------------#--
#----------------------------------------#--
    file_path <- "Book1.xlsx"
#----------------------------------------#--
#----------------------------------------#--
data <- as.data.frame(read_excel(file_path, sheet = 1, col_names = TRUE, .name_repair = "minimal"))
# Read the CSV file
#data <- read.csv(file_path, header = TRUE, sep = ";", stringsAsFactors = FALSE,check.names = FALSE)

# Convert the first column to row names
rownames(data) <- data[[1]]  # Assign first column as row names
data <- data[, -1]  # Remove the first column from the dataset

# Remove empty rows and columns dynamically
data <- data[rowSums(is.na(data) | data == "") < ncol(data), ]  # Remove empty rows
data <- data[, colSums(is.na(data) | data == "") < nrow(data)]  # Remove empty columns

# Convert all remaining data to numeric
data[] <- lapply(data, function(x) as.numeric(as.character(x)))

# Correctly extract the beneficial row
beneficial <- suppressWarnings(as.numeric(data[1, ]))  # Convert first row to numeric
data <- data[-1, ]  # Remove beneficial row from data

# Ensure no NA values in beneficial
if (any(is.na(beneficial))) {
  stop("Error: Beneficial row contains NA values. Please check input data.")
}



#---------------------------------------------------------------------------------------------





normalize_data<-function(data)
{
  # (1) Remove all columns where all values are the same(zero variance) :
  constant_cols <- which(apply(data, 2, function(x) length(unique(na.omit(x))) == 1))
  
  if (length(constant_cols) > 0) 
  {
    cat("\n Warning: Removing constant criteria with zero variance :\n")
    print(colnames(data)[constant_cols])
    cat("\n\n")
    
    # Remove constant columns from data
    data <- data[, -constant_cols, drop = FALSE]
    beneficial<-beneficial[-constant_cols]
  }
  
  # (2) Remove columns that contain any missing values(NA) :
  cols_with_na <- which(colSums(is.na(data)) > 0)  # Find columns with at least 1 NA
  
  if (length(cols_with_na) > 0) 
  {
    cat("\n Warning: Removing columns with missing values :\n")
    print(colnames(data)[cols_with_na])
    cat("\n\n")
    
    # Remove columns with missing values(NA)
    data <- data[, -cols_with_na, drop = FALSE]
    beneficial<-beneficial[-cols_with_na]
  }

  
  # Normalization of data
  norm_data<-data
  r<-nrow(data)
  c<-ncol(data)
  
  for(i in 1:r)
  {
    for(j in 1:c)
    {
      if(beneficial[j]==1)
      {
        norm_data[i,j]<-(data[i,j]-min(data[,j]))/(max(data[,j])-min(data[,j]))
        
      }
      else if(beneficial[j]==0)
      {
        norm_data[i,j]<-(data[i,j]-max(data[,j]))/(min(data[,j])-max(data[,j]))
      }
      else
      {
        stop("Benefitial factor is not binary. Check and try again!")
      }
    }
  }
  
  return(norm_data)
}



#---------------------------------------------------------------------------------------------




criteria_weight<-function(data)
{
  data<-normalize_data(data)
  std_dev<-apply(data,2,sd)
  r<-nrow(data)
  c<-ncol(data)
  cor_matrix<-matrix(NA,nrow = c,ncol = c)
  I<-numeric(c)
  W<-array(NA,dim = c)
  for (j in 1:c) 
  {
    for (k in 1:c) 
    {
      #correlation matrix : 
      cor_matrix[j,k]<-dcor(data[,j],data[,k])
    }
  }
  # Calculating the information content
  for (i in 1:c) 
  {
    I[i]<-std_dev[i]*sum(1-cor_matrix[i,])
  }
  # Calculating the weights of each criteria
  for (i in 1:c) 
  {
    W[i]<-I[i]/sum(I)
  }
  cat("\n\n")
  cat("--------------------------------\n")
  cat("The weight of each criteria is :\n")
  cat("--------------------------------\n")
  names(W)<-colnames(data)
  print(kable(W, format = "markdown", digits = 2))
  
  
  cat("\n\n")
  cat("----------------------------------------------------\n")
  cat("The weight of each criteria in descending order is :\n")
  cat("----------------------------------------------------\n")
  sorted_weights<-sort(W,decreasing = TRUE)
  print(kable(sorted_weights, format = "markdown", digits = 2))
  
  cat("\n\n")
  cat("----------------------------------------------\n")
  cat("The following barplot represents the weights :\n")
  cat("----------------------------------------------\n")
  W_df<-data.frame(Criteria=names(W),Weight=W)
  plot<-ggplot(W_df, aes(x = reorder(Criteria, -Weight), y = Weight, fill = Criteria)) +
    geom_bar(stat = "identity", color = "black") +  # Create the bar for each criteria
    geom_text(aes(label = format(Weight, digits=5,nsmall=4)), vjust = -0.5, size = 5) +  # Show the weight of each criteria
    theme_minimal() +
    labs(title = "Final Weights of Each Criterion",
         x = "Criteria", y = "Weight") +
    theme(legend.position = "none")  # Hide the legend
  
  
  print(plot)
  return(list("Weights" = W, "Sorted Weights" = sorted_weights))
}




#---------------------------------------------------------------------------------------------




analysis <- function(data) 
{
  r <- nrow(data)
  
  # Initialize indices as NA
  npv             <- numeric(r)
  lcoe            <- numeric(r)
  capacity_factor <- numeric(r)
  Cp              <- numeric(r)
  
  # Extract the names of the alternatives(rownames)
  alternative_name<-rownames(data)
  
  
  #  (1) Net Present Value (NPV)
  # Checks  whether all required columns exist in data before proceeding with calculations
  required_cols_npv <- c("Cash Flow(€/yr)", "Discount Rate (%/year)", "CAPEX (€)", "Subsidies ( €/yr)", "Life Exp. (years)")
  
  if (all(required_cols_npv %in% colnames(data))) 
  {
    cf <- data[,"Cash Flow(€/yr)"]
    dr <- data[,"Discount Rate (%/year)"]
    capex <- data[,"CAPEX (€)"]
    subsidies <- data[, "Subsidies ( €/yr)"]
    exp_life <- data[,"Life Exp. (years)"]
    opex <- data[,"OPEX (€/yr)"]
    decom<-data[,"Dec. Cost (€)"]
    
    for (i in 1:r) 
    {
      if (!anyNA(c(cf[i], dr[i], capex[i], subsidies[i],opex[i], exp_life[i])))
      {
        npv_value <- sum((cf[i] + subsidies[i]-opex[i]) / (1 + dr[i] / 100)^(1:exp_life[i]))
        npv[i] <- npv_value - capex[i]-decom[i]/((1 + dr[i] / 100)^exp_life[i])
      }
    }
  }
  
  #  (2) Levelized Cost of Energy (LCOE)
  required_cols_lcoe <- c("AEP(MWh/yr)", "OPEX (€/yr)", "Discount Rate (%/year)", "CAPEX (€)", "Life Exp. (years)")
  numerator<-numeric(r)
  denominator<-numeric(r)
  if (all(required_cols_lcoe %in% colnames(data))) 
  {
    aep <- data[,"AEP(MWh/yr)"]
    opex <- data[,"OPEX (€/yr)"]
    
    for (i in 1:r) 
    {
      if (!anyNA(c(aep[i], opex[i], dr[i], capex[i], exp_life[i]))) 
      {
        numerator[i] <- sum(opex[i] / (1 + dr[i] / 100)^(1:exp_life[i]))
        denominator[i] <- sum(aep[i] / (1 + dr[i] / 100)^(1:exp_life[i]))
        lcoe[i] <- (capex[i]/(1+dr[i]/100) + numerator[i]) / denominator[i]
      }
    }
  }
  
  #  (3) Capacity Factor (CF)
  required_cols_cf <- c("AEP(MWh/yr)", "Rated Power (MW)")
  
  if (all(required_cols_cf %in% colnames(data))) 
  {
    rated_power <- data[,"Rated Power (MW)"]
    
    for (i in 1:r) 
    {
      if (!anyNA(c(aep[i], rated_power[i]))) 
      {
        capacity_factor[i] <- (aep[i] / (rated_power[i] * 8760)) * 100
      }
    }
  }
  
  #  (4) Power Coefficient (Cp)
  required_cols_cp <- c("Blade Length (m)", "Altitude (m)", "Hub Height (m)", "Avg. Wind v (m/s)", "AEP(MWh/yr)")
  
  if (all(required_cols_cp %in% colnames(data))) 
  {
    radius <- data[,"Blade Length (m)"]
    h <- data[,"Altitude (m)"]
    hub_height <- data[,"Hub Height (m)"]
    avg_wind_speed <- data[,"Avg. Wind v (m/s)"]
    
    for (i in 1:r) 
    {
      if (!anyNA(c(radius[i], h[i], hub_height[i], avg_wind_speed[i], aep[i]))) 
      {
        A <- pi * (radius^2)
        rho_0 <- 1.225 # kg/m^3
        T0<-288.15     # KELVIN
        a<-0.0065      # Kelvin/meter
        g0<-9.80665    # m/s^2
        R<-287.053     #Joule/kg/Kelvin
        height<-h+hub_height
        rho<-rho_0*(T0/(T0-a*height))*(1-a*height/T0)^(g0/(a*R))
        P_available <- (1/2) * rho * A * avg_wind_speed^3  # in W
        P_actual <- aep / 8760  # in MW = 10^6 W
        Cp <- (P_actual / P_available) * 10^6
      }
    }
  }
  
  #  Create Result Table
  result <- data.frame(
    "Project"=alternative_name,
    "Capacity Factor (%)" = capacity_factor,
    "L.C.O.E (€/MWh)" = lcoe,
    "NPV (€)" = npv,
    "Power Coefficient" = Cp,
    check.names = FALSE
  )
  
  cat("\n\n")
  cat("-------------------------------------------------\n")
  cat("The following matrix shows the analysis results :\n")
  cat("-------------------------------------------------\n")
  print(kable(result, format = "markdown", digits = 3))
  
  return(result)
}




#---------------------------------------------------------------------------------------------




save_results_to_excel <- function(normalized_data, criteria_weights, sorted_weights, analysis_results, plot) 
{
  # Load the working excel file
  wb <- loadWorkbook(file_path)
  #  Function to Apply Formatting to a Sheet
  format_sheet <- function(wb, sheet_name, data)
  {
    if (sheet_name %in% names(wb))
    {
      removeWorksheet(wb, sheet_name)  # ✅ Remove old sheet if it exists
    }
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet_name, data, rowNames = FALSE)
    # Define styles
    header_style <- createStyle(fontSize = 12, textDecoration = "bold", halign = "center", fgFill = "#D9EAD3")
    body_style <- createStyle(halign = "center")
    # Apply styles
    addStyle(wb, sheet_name, header_style, rows = 1, cols = 1:ncol(data), gridExpand = TRUE)
    addStyle(wb, sheet_name, body_style, rows = 2:(nrow(data) + 1), cols = 1:ncol(data), gridExpand = TRUE)
    # Auto-fit columns for better readability
    setColWidths(wb, sheet_name, cols = 1:ncol(data), widths = "auto")
  }
  #  Convert Dataframes to Ensure Row Names are Included
  criteria_weights_df <- data.frame(Criteria = names(criteria_weights), Weight = criteria_weights)
  sorted_weights_df   <- data.frame( "Sorted Criteria" = names(sorted_weights), Weight = sorted_weights,check.names = FALSE)
  #  Add Sheets with Formatting
  format_sheet(wb, "Normalized Data", normalized_data)
  format_sheet(wb, "Criteria Weights", criteria_weights_df)  # Fixed Row Names
  format_sheet(wb, "Sorted Weights", sorted_weights_df)      # Fixed Row Names
  format_sheet(wb, "Analysis Results", analysis_results)
  #  Conditional Formatting (Highlight NA values in red)
  na_style <- createStyle(fontColour = "#FFFFFF", fgFill = "#FF0000")  # Red background
  conditionalFormatting(wb, "Analysis Results", cols = 2:ncol(analysis_results), rows = 2:(nrow(analysis_results) + 1),
                        rule = "ISNA(A1)", style = na_style)
  #  Step 1: Save the Plot as an Image
  plot_file <- "Criteria_Weights_Plot.png"
  ggsave(plot_file, plot = plot, width = 35, height = 25, dpi = 300)
  #  Step 2: Add a Sheet for the Plot and Insert the Image
  #  Step 2: **Remove "Plot" Sheet First** & Insert the Image
  if ("Plot" %in% names(wb))
  {
    removeWorksheet(wb, "Plot")  # ✅ Ensure we remove the old sheet
  }
  addWorksheet(wb, "Plot")
  insertImage(wb, "Plot", plot_file, width = 30, height = 20, startRow = 2, startCol = 2)
  #  Save the Excel file
  saveWorkbook(wb, file_path, overwrite = TRUE)
  cat("\n📁 All results saved in '",file_path,"'")
}





#---------------------------------------------------------------------------------------------




suppressMessages(suppressWarnings({
  cat("\n + Running Normalization...\n")
  normalized_data <- normalize_data(data)
  
  cat("\n + Running Criteria Weight Calculation...\n")
  weights_result <- criteria_weight(data)
  criteria_weights <- weights_result$W
  sorted_weights <- sort(criteria_weights,decreasing = TRUE)
  
  #  Extract the plot from `criteria_weight()`
  W_df <- data.frame(Criteria = names(criteria_weights), Weight = criteria_weights)
  plot <- ggplot(W_df, aes(x = reorder(Criteria, -Weight), y = Weight, fill = Criteria)) +
    geom_bar(stat = "identity", color = "black") +
    geom_text(aes(label = format(Weight, digits = 5, nsmall = 4)), vjust = -0.5, size = 8) +
    theme_minimal() +
    labs(title = "Final Weights of Each Criterion", x = "Criteria", y = "Weight") +
    theme(legend.position = "none",  # Hide the legend
        axis.text.x = element_text(size = 14, angle = 45, hjust = 1),  # Increase x-axis label size and rotate
        axis.text.y = element_text(size = 14),  # Increase y-axis label size
        axis.title.x = element_text(size = 16, face = "bold"),  # Increase x-axis title size
        axis.title.y = element_text(size = 16, face = "bold"),  # Increase y-axis title size
        plot.title = element_text(size = 18, face = "bold", hjust = 0.5)  # Increase title size)
    )
        
  cat("\n + Running Financial & Performance Analysis...\n")
  analysis_results <- analysis(data)
  
  # Save everything to Excel
  save_results_to_excel(normalized_data, criteria_weights, sorted_weights, analysis_results,plot)
}))

