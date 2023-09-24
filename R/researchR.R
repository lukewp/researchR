### FUNCTION BLOCK:
#' spss_to_excel
#' 
#' Generates an Excel representation of an SPSS response file with separate sheets 
#' for the raw responses, variable descriptions, and codes.
#' 
#' @param x name of the haven-generated object where the data will go
#' @param spss_path string with filename/path of the source SPSS .sav file
#' @param excel_path string with filename/path of the destination Excel .xlsx file
#' @param overwrite_excel_path logical indicating whether to overwrite (erase) any existing at excel_path
#' @param return_data logical indicating whether to return an R object containing the data or not
#' 
#' @returns nothing if successful; errors from child functions if exist; dataframe if "return_dataframe" is TRUE
#' 
#' @export
spss_to_excel <- function(x, spss_path, 
                          excel_path,
                          overwrite_excel_path = FALSE,
                          return_data = FALSE) {

  library(haven, openxlsx)
  
  # load spss data to temp table:
  raw_data <- haven::read_spss(spss_path)
  
  # initialize in-memory workbook:
  wb <- openxlsx::createWorkbook()
  
  # initialize sheets:
  openxlsx::addWorksheet(wb, "raw")
  openxlsx::addWorksheet(wb, "rawVars")
  openxlsx::addWorksheet(wb, "rawVals")
  
  # write full dataset to "raw" sheet in a table:
  openxlsx::writeDataTable(wb, "raw", x = raw_data, rowNames = TRUE, xy = c("B", 5))
  
  # create data frame from SPSS variable descriptions and write to "rawVars" table/sheet:
  raw_vars <- lapply(raw_data, attr, "label")
  raw_vars <- unlist(raw_vars)
  openxlsx::writeDataTable(wb, "rawVars", x = data.frame(raw_vars), rowNames = TRUE, xy = c("B", 5))
  
  # create data frame from SPSS value descriptions and write to "rawVals" table/sheet:
  raw_vals <- lapply(raw_data, attr, "labels")
  raw_vals <- unlist(raw_vals)
  openxlsx::writeDataTable(wb, "rawVals", x = data.frame(raw_vals), rowNames = TRUE, xy = c("B", 5))
  
  if (return_data == TRUE) {
    openxlsx::saveWorkbook(wb, excel_path, overwrite = overwrite_excel_path)
    return(raw_data)
  }

    # write the workbook and exit, delivering response from the openxlsx::saveWorkbook operation
  if (return_data == FALSE) {
    return(openxlsx::saveWorkbook(wb, excel_path, overwrite = overwrite_excel_path))  
  }
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
#' @param recenter logical TRUE to mean-center the value rather than just recode/rename the variable
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
#' @param flip = FALSE logical indicating whether to flip the scale
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
  # testing variables:
  # library(haven)
  # x <- haven::read_spss("892579-data-Eggs_and_Syrup-09-11-2023-CX--Default.sav")
  # existing_vars <- c('Q11_7',	'Q11_8',	'Q11_9',	'Q11_10',	'Q11_11',	'Q11_12',	'Q11_13',	'Q11_14',	'Q11_15',	'Q11_16',	'Q11_17',	'Q11_18',	'Q11_19',	'Q11_20',	'Q11_21',	'Q11_22',	'Q11_23',	'Q11_24',	'Q11_25',	'Q11_26',	'Q11_27',	'Q11_28',	'Q11_29',	'Q11_30',	'Q11_31',	'Q11_32',	'Q11_33',	'Q11_34',	'Q11_35',	'Q12_7',	'Q12_8',	'Q12_9',	'Q12_10',	'Q12_11',	'Q12_12',	'Q12_13',	'Q12_14',	'Q12_15',	'Q12_16',	'Q12_17',	'Q12_18',	'Q12_19',	'Q12_20',	'Q12_21',	'Q12_22',	'Q12_23',	'Q12_24',	'Q12_25',	'Q12_26',	'Q12_27',	'Q12_28',	'Q12_29',	'Q12_30',	'Q12_31',	'Q12_32',	'Q12_33',	'Q12_34',	'Q12_35',	'Q12_36',	'Q12_37',	'Q12_38',	'Q12_39',	'Q13_7',	'Q13_8',	'Q13_9',	'Q13_10',	'Q13_11',	'Q13_12',	'Q13_13',	'Q13_14',	'Q13_15',	'Q13_16',	'Q13_17',	'Q13_18',	'Q13_19',	'Q13_20',	'Q13_21',	'Q13_22',	'Q13_23',	'Q14_7',	'Q14_8',	'Q14_9',	'Q14_10',	'Q14_11',	'Q14_12',	'Q14_13',	'Q14_14',	'Q14_15',	'Q14_16',	'Q14_17',	'Q14_18',	'Q14_19',	'Q14_20',	'Q14_21',	'Q14_22',	'Q14_23',	'Q14_24',	'Q14_25',	'Q14_26',	'Q14_27',	'Q14_28',	'Q14_29',	'Q14_30',	'Q14_31',	'Q14_32',	'Q14_33',	'Q14_34',	'Q14_35',	'Q14_36',	'Q14_37',	'Q14_38',	'Q14_39',	'Q14_40',	'Q14_41',	'Q14_42',	'Q14_43')
  # recode_vars <- c('buy_american',	'born_usa',	'parents_born_usa',	'i_recycle',	'i_vote_federal',	'i_vote_state',	'i_vote_local',	'go_to_work_sick',	'diet_informed',	'pay_for_environmentally_friendly_products',	'family_is_most_important',	'in_control_of_weight',	'existential_book_angst',	'spanish_in_home',	'english_in_home',	'nature_understanding',	'comfortable_being_unconventional',	'respect_customs_and_beliefs',	'exercises_regularly',	'sacrifice_family_for_career',	'breakfast_most_important',	'lunch_most_important',	'dinner_most_important',	'optimist',	'participate_in_protests',	'patronize_environmentally_friendly_companies',	'environmental_soundness_good_for_business',	'travel_abroad',	'prefers_travel_domestic',	'enjoys_cooking',	'tries_new_recipes',	'guilty_about_sweets',	'flavor_as_important_as_healthiness',	'no_guilt_over_fat_indulgence',	'eats_healthy',	'shops_organic_natural',	'not_price_sensitive_around_food',	'enjoys_foreign_foods',	'fast_food_is_junk',	'like_healthy_fast_food',	'eats_gourmet_food',	'snacks_healthy',	'nonartificial_preference',	'eats_precooked',	'local_eggs',	'organic_eggs',	'local_fruits_veggies',	'organic_fruits_veggies',	'local_meats',	'organic_meats',	'local_dairy',	'organic_dairy',	'local_organic_supports_local_farmers',	'local_organic_source_knowledge',	'local_organic_are_healthier',	'local_organic_carbon_footprint',	'farmers_markets',	'belongs_csa',	'avoids_sugar_unhealthy',	'avoids_sugar_unethical',	'prefers_non_sugar_artificial_sweeteners',	'prefers_natural_alternative_sugars',	'tries_new_drinks',	'buys_alcohol',	'collects_wine',	'winemaking_knowledge',	'restaurant_wine_chooser',	'collects_whiskey',	'whiskeymaking_knowledge',	'explores_beer',	'beermaking_knowledge',	'low_sugar_alcohol',	'low_carb_alcohol',	'low_cal_alcohol',	'na_beer',	'na_wine',	'na_cocktails',	'avoids_alcohol_unhealthy',	'fancy_alcohol_too_expensive',	'breakfast_syrup_user',	'breakfast_syrup_careful_user',	'breakfast_syrup_collector',	'breakfast_syrup_low_cal',	'breakfast_syrup_flavored_preference',	'table_syrup_user',	'pure_maple_user',	'pure_maple_souvenir',	'pure_maple_gift',	'pure_maple_buy_grocery',	'pure_maple_buy_farm_stand',	'pure_maple_buy_online',	'explores_breakfast_syrup',	'explores_pure_maple',	'pure_maple_discerning_palate',	'sugarmaking_knowledge',	'breakfast_syrup_mixology',	'breakfast_syrup_non_breakfast',	'pure_maple_mixology',	'pure_maple_non_breakfast',	'table_syrup_mixology',	'table_syrup_non_breakfast',	'enjoys_baking',	'baking_subs_pure_maple',	'table_syrup_prefers_local',	'maple_syrup_prefers_local',	'table_syrup_prefers_organic',	'pure_maple_prefers_organic',	'pure_maple_expensive_treat',	'pure_maple_unhealthy_treat',	'table_syrup_good_for_corn_farmers',	'pure_maple_good_for_sugarmakers',	'open_to_maple_subsidies',	'pure_maple_not_available',	'pure_maple_organic_not_available',	'pure_maple_local_not_available',	'american_pure_maple_preference')
  # scale <- c(1:6)
  # na_value = 0
  # flip = FALSE
  # low_appendage = "_no"
  # high_appendage = "_yes"

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
