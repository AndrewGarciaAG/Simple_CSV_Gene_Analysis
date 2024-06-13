# Install and load necessary package
if (!require("readxl")) install.packages("readxl", dependencies = TRUE)
library(readxl)

# Increase max.print limit
options(max.print = 10000)

# Function to process sheets and find matching genes
process_sheets <- function(file_path, sheet_names, comparison_sheets) {
  # List to store matching genes from each sheet for comparison
  all_matching_genes <- list()
  
  for (sheet_name in sheet_names) {
    cat("Processing sheet:", sheet_name, "\n")
    
    # Read data from the current sheet
    excel_data <- read_excel(file_path, sheet = sheet_name)
    
    # Print column names
    cat("Column names for", sheet_name, ":\n")
    print(colnames(excel_data))
    cat("\n")
    
    # Convert gene names to lowercase if the columns exist
    if ("Target_genes" %in% colnames(excel_data) && "gene" %in% colnames(excel_data)) {
      excel_data$Target_genes <- tolower(as.character(excel_data$Target_genes))
      excel_data$gene <- tolower(as.character(excel_data$gene))
      
      # Find matching genes
      matching_genes <- intersect(excel_data$Target_genes, excel_data$gene)
      
      # Add matching genes to the list for cross-sheet comparison
      if (sheet_name %in% comparison_sheets) {
        all_matching_genes[[sheet_name]] <- matching_genes
      }
      
      # Print matching genes for the current sheet
      cat("Matching genes for", sheet_name, ":\n")
      print(matching_genes)
      cat("\n")
      
      # Write matching genes to a CSV file
      write.csv(matching_genes, file = paste0("matching_genes_", sheet_name, ".csv"))
      
      # Tally the number of matching genes
      num_matches <- length(matching_genes)
      cat("Number of matches for", sheet_name, ":", num_matches, "\n\n")
    } else {
      cat("One or both of the columns 'Target_genes' and 'gene' not found in the sheet", sheet_name, "\n\n")
    }
  }
  
  # Find matches common to comparison sheets
  common_matches <- Reduce(intersect, all_matching_genes)
  
  # Print common matches
  cat("Common matches across", paste(comparison_sheets, collapse = ", "), ":\n")
  print(common_matches)
  
  # Write common matches to a CSV file
  write.csv(common_matches, file = "common_matches_across_selected_sheets.csv")
}

# Function to find min and max values for specific columns
find_min_max_columns <- function(file_path, sheet_names, column_indices) {
  min_values <- rep(Inf, length(column_indices))
  max_values <- rep(-Inf, length(column_indices))
  
  for (sheet_name in sheet_names) {
    # Read data from the current sheet
    excel_data <- read_excel(file_path, sheet = sheet_name)
    
    # Check if the sheet has enough columns
    if (ncol(excel_data) >= max(column_indices)) {
      for (i in seq_along(column_indices)) {
        col_index <- column_indices[i]
        excel_data[[col_index]] <- as.numeric(excel_data[[col_index]])
        
        # Update min and max values
        min_values[i] <- min(min_values[i], min(excel_data[[col_index]], na.rm = TRUE))
        max_values[i] <- max(max_values[i], max(excel_data[[col_index]], na.rm = TRUE))
      }
    } else {
      cat("Sheet", sheet_name, "does not have enough columns.\n")
    }
  }
  
  return(list(min_values = min_values, max_values = max_values))
}

# Function to create a table for matched genes and their scores
create_matched_gene_table <- function(file_path, sheet_name, matched_genes) {
  # Read data from the sheet
  excel_data <- read_excel(file_path, sheet = sheet_name)
  
  # Ensure there are at least 2 columns
  if (ncol(excel_data) < 2) stop("Sheet", sheet_name, "does not have at least 2 columns.")
  
  # Convert gene names to lowercase for matching
  excel_data$gene <- tolower(as.character(excel_data$gene))
  
  # Filter rows with matching genes
  matched_data <- excel_data[excel_data$gene %in% matched_genes, ]
  
  # Get the name of the second column (binding score)
  binding_score_col <- names(excel_data)[2]
  
  # Select only the columns for gene names and binding scores
  result_table <- matched_data[, c("gene", binding_score_col)]
  
  # Order by binding score in descending order
  result_table <- result_table[order(-result_table[[binding_score_col]]), ]
  
  # Return the result table
  return(result_table)
}

# Main function to run the analysis
run_analysis <- function(file_path) {
  sheet_names <- c("ATF3vGPS2si", "ATF4vGPS2si", "ATF5vGPS2si", "GPS2vGPS2si")
  comparison_sheets <- c("ATF3vGPS2si", "ATF4vGPS2si", "ATF5vGPS2si")
  
  process_sheets(file_path, sheet_names, comparison_sheets)
  
  column_indices <- c(2, 4)
  min_max_values <- find_min_max_columns(file_path, comparison_sheets, column_indices)
  cat("Range for columns", paste(column_indices, collapse = ", "), "across sheets:\n")
  cat("Min values:", min_max_values$min_values, "\n")
  cat("Max values:", min_max_values$max_values, "\n")
  
  for (sheet_name in comparison_sheets) {
    matched_genes <- all_matching_genes[[sheet_name]]
    result_table <- create_matched_gene_table(file_path, sheet_name, matched_genes)
    
    cat(sheet_name, "Matched Genes and Scores:\n")
    print(result_table)
    write.csv(result_table, file = paste0(sheet_name, "_Matched_Genes_and_Scores.csv"))
  }
}

# Example usage
file_path <- "path/to/your/excel/file.xlsx" # Update this with the actual file path
run_analysis(file_path)
