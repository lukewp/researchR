### FUNCTION BLOCK:
#' spss_to_excel
#' 
#' Generates an Excel representation of an SPSS response file with separate sheets 
#' for the raw responses, variable descriptions, and codes.
#' 
#' @param x name of the haven-generated object where the data will go
#' @param spss_path string with filename/path of the source SPSS .sav file OR list of strings with filename/path of the source SPSS.sav file, in which case files must be the same format and will be concatenated
#' @param excel_path string with filename/path of the destination Excel .xlsx file
#' @param overwrite_excel_path logical indicating whether to overwrite (erase) any existing at excel_path
#' @param return_data logical indicating whether to return an R object containing the data or not
#' @param source_field_name string used to name a field that'll be used if multiple files passed with spss_path. FALSE (default) to not create a source variable.
#' 
#' @returns nothing if successful; errors from child functions if exist; dataframe if "return_dataframe" is TRUE
#' 
#' @export
spss_to_excel <- function(x, spss_path, 
                          excel_path,
                          overwrite_excel_path = FALSE,
                          return_data = FALSE,
                          source_field_name = FALSE) {
  
  library(haven, openxlsx)
  
  # Function to read and process SPSS data
  read_and_process_spss <- function(file_path) {
    data <- haven::read_spss(file_path)
    if (source_field_name != FALSE) {
      data[[source_field_name]] <- file_path
    }
    return(data)
  }
  
  # Load spss data from all files
  raw_data <- do.call(rbind, lapply(spss_path, read_and_process_spss))
  
  # initialize in-memory workbook:
  wb <- openxlsx::createWorkbook()
  
  # initialize sheets:
  sheet_names <- c("raw", "rawVars", "rawVals")
  lapply(sheet_names, function(sheet) openxlsx::addWorksheet(wb, sheet))
  
  # write full dataset to "raw" sheet in a table:
  openxlsx::writeDataTable(wb, "raw", x = raw_data, rowNames = TRUE, xy = c("B", 5))
  
  # create data frame from SPSS variable descriptions and write to "rawVars" table/sheet:
  raw_vars <- unlist(lapply(raw_data, attr, "label"))
  openxlsx::writeDataTable(wb, "rawVars", x = data.frame(raw_vars), rowNames = TRUE, xy = c("B", 5))
  
  # create data frame from SPSS value descriptions and write to "rawVals" table/sheet:
  raw_vals <- unlist(lapply(raw_data, attr, "labels"))
  openxlsx::writeDataTable(wb, "rawVals", x = data.frame(raw_vals), rowNames = TRUE, xy = c("B", 5))
  
  # Save workbook and return data if specified:
  if (return_data) {
    openxlsx::saveWorkbook(wb, excel_path, overwrite = overwrite_excel_path)
    return(raw_data)
  }
  
  # Save the workbook and exit, delivering response from the openxlsx::saveWorkbook operation
  return(openxlsx::saveWorkbook(wb, excel_path, overwrite = overwrite_excel_path))  
}

#' recenter_recode_vars
#' 
#' Useful for Likert-style statement-agreement blocks where a series of questions 
#' or statements are asked in a block on an online survey, as the respondent will 
#' tend to ground their answers in the visible frame of the survey from one pole
#' to the other of the Likert scale. Methodologically, smarter to do these one block
#' at a time vs. attacking all Likert-style statement-agreement blocks in one go,
#' because respondents' means will often drift from block to block, especially in 
#' online surveys.
#' 
#' Takes an SPSS dataset, a list of variable names we want to recode, a list of 
#'  variable names to which we want them recoded, calculates the mean-centered 
#' value for the variables and returns a new dataset with those new vars appended.
#' new var names (second column of vars_names) must not already exist in the dataset
#' or they'll be overwritten in-place. 
#' 
#' @param x haven-generated SPSS dataset we're working on
#' @param existing_vars list of input variables. must be same length/order as recode_vars
#' @param recode_vars list of target variable names. must be same length/order as existing_vars
#' @param recenter logical TRUE to mean-center the value rather than just recode/rename the variable, TRUE by default
#' @param na_value integer value for recoding data to NAs. 0 by default.
#' 
#' @returns a haven-generated SPSS dataset with new fields as indicated in vars_names, inheriting the original attributes of the old variables.
#' 
#' @export

recenter_recode_vars <- function(x, existing_vars,
                                 recode_vars,
                                 recenter = TRUE,
                                 na_value = 0) {
  
  # make data_frame from vars (really just to throw an error if they're not the same length):
  var_frame <- data.frame(existing_vars <- existing_vars,
                          recode_vars <- recode_vars)
  
  # create a temp df to work in so we're safely handling the input:
  target_data <- x[, existing_vars]
  
  # handle NA recode if needed:
  if (na_value > 0) {
    target_data[, existing_vars][target_data[, existing_vars] == na_value] <- NA
  }
  
  # gather row means and assign to temp table: 
  target_data$avg_value <- rowMeans(target_data[, existing_vars], na.rm=TRUE)
  
  # loop over the inputs and create new recode vars on the temp df
  for (i in 1:nrow(var_frame)) {
    existing_var_name <- var_frame$existing_vars[i]
    recode_var_name <- var_frame$recode_vars[i]
    
    # Copy the values of the existing variable to the new variable
    target_data[recode_var_name] <- target_data[existing_var_name]
    
    if (recenter == TRUE) {
      # Copy the values of the existing variable to the new variable
      target_data[recode_var_name] <- target_data[existing_var_name] - target_data$avg_value
    }
    
    # Copy the attributse from the existing variable to the new variable
    attributes(target_data[[recode_var_name]]) <- attributes(target_data[[existing_var_name]])
  }
  
  # return the input data with the new recoded variables attached to the end!
  return(cbind(x, target_data[var_frame$recode_vars]))
}

#' split_recode_nonnegative_vars
#' 
#' Useful to transform Likert-scale question responses into polar variables for 
#' NMF analysis where all inputs have to be positive numbers. In this case a 1-5
#' scale question would be recoded so 1s and 2s would be 2s and 1s for a negative
#' response variable and 0s for a positive response variable; 4s and 5s would be 
#' recoded to 1s and 2s for a positive response variable and 0s for a negative response 
#' variable, 3s would be recoded as a 0 for both responses. It is based on means,
#' so even-numbered scales end up with half-values. This is mathematically irrelevant
#' for both correlation-based clustering and and NMF: consistent scales matter most.
#' 
#' Examples:
#' 1-4 scale (mean=2.5) would be 1.5,0.5,0,0 (no) and 0,0,.5,1.5 (yes)
#' 1-5 scale (mean=3) would be 2,1,0,0,0 (no) and 0,0,0,1,2 (yes)
#' 1-6 scale (mean=3.5) would be 2.5,1.5,0.5,0,0,0 (no) and 0,0,0,0.5,1.5,2.5 (yes)
#' 
#' Takes an SPSS dataset, a list of variable names we want to recode, a list of 
#' variable names to which we want them recoded, calculates the mean-centered value 
#' for the variables and returns a new dataset with those new vars appended, with 
#' string appendages for the top and bottom scale values ("_yes" and "_no" are default 
#' values here). New var names with appendages must not already exist in the dataset 
#' or they'll be overwritten in-place. 
#' 
#' Note: Do not mix variables with different response values or scales here. For
#' example, if you have Q11_1, Q11_2, Q11_3 on a 6-point Likert with 6 = 'strongly 
#' agree' and 1 = 'strongly disagree', and 7 = "not applicable", and Q12_1, Q12_2, 
#' Q12_3 on a 5-point Likert with 14 = 'strongly disagree', 16 = 'neither agree 
#' nor disagree', 18 = 'strongly agree', and 19 = 'not applicable' ... do not recode
#' them in the same set. I don't know why you would but it's definitely going to 
#' yield undesirable results.
#' 
#' @param x haven-generated SPSS dataset to manipulate
#' @param existing_vars list of variable names valid for x that we want to recode (must be same length and order as recode_vars)
#' @param recode_vars list of variable names to which we want to recode (must be same length and order as existing_vars)
#' @param scale list of integer values for existing_vars, e.g. c(1:6) for 6-point scale from 1-6, c(18:22) for 5-point scale from 18-22
#' @param na_value = 0 if there's a hard-coded response value we want to set to NA, indicate it here
#' @param flip = FALSE logical indicating whether to flip the scale in case the higher numeric response value corresponds to the lower response
#' @param low_appendage = "_no" the lower response
#' @param high_appendage = "_yes" the higher response
#' 
#' @returns a haven-generated SPSS dataset with your new recoded variables appended to the back.
#' 
#' @export
split_recode_nonnegative_vars <- function(x, existing_vars, 
                                          recode_vars, 
                                          scale,
                                          na_value = 0,
                                          flip = FALSE,
                                          low_appendage = "_no",
                                          high_appendage = "_yes") {
  
  # make data_frame from vars (really just to throw an error if they're not the same length):
  var_frame <- data.frame(existing_vars <- existing_vars,
                          recode_vars <- recode_vars)
  
  # create a temp df to work in so we're safely handling the input:
  target_data <- x[, existing_vars]
  
  # initialize an output colname list:
  output_cols <- c()
  
  # handle NA recode if needed:
  #TODO: revise to include "zapping" function from haven
  if (na_value > 0) {
    target_data[, existing_vars][target_data[, existing_vars] == na_value] <- NA
  }
  
  # handle flip, just reversing the string appendages:
  if (flip == TRUE) {
    tempvar <- low_appendage
    low_appendage <- high_appendage
    high_appendage <- tempvar
  }
  
  # determine center:
  center <- mean(scale)
  
  for (i in 1:nrow(var_frame)) {
    # Figure out names:
    existing_var_name <- var_frame$existing_vars[i]
    recode_var_name_high <- paste(var_frame$recode_vars[i], high_appendage, sep = "")
    recode_var_name_low <- paste(var_frame$recode_vars[i], low_appendage, sep = "")
    
    # Instantiate the variables:
    target_data[[recode_var_name_high]] <- 0
    target_data[[recode_var_name_low]] <- 0
    
    # Copy the attributes from the existing variable to the new variables
    attributes(target_data[[recode_var_name_high]]) <- attributes(target_data[[existing_var_name]])
    attributes(target_data[[recode_var_name_low]]) <- attributes(target_data[[existing_var_name]])
    
    # Transform the values of the existing variable to the new high and low variables
    target_data[target_data[[existing_var_name]] > center, recode_var_name_high] <- 
      target_data[target_data[[existing_var_name]] > center, existing_var_name]
    
    target_data[target_data[[existing_var_name]] < center, recode_var_name_low] <- 
      -1 * (target_data[target_data[[existing_var_name]] < center, existing_var_name] - center)
    
    # Add to the output cols variable:
    output_cols <- c(output_cols, recode_var_name_high, recode_var_name_low)
  }
  
  return(cbind(x, target_data[output_cols]))
  
}
### END FUNCTION BLOCK
