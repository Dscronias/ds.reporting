#' Oneway report of multiple categorical variables to Excel
#'
#' This functions creates an Excel workbook and exports a oneway table
#' of multiple variables with counts and percentages
#'
#' @param Data dataframe
#' @param var Variable to use
#' @param rounding number of digits for rounding (int)
#' @return Count table
#' @export

svy_ow <- function(data, var, rounding = 2) {
    table <- data %>%
        survey_count({{var}}, .drop = FALSE) %>%
        select(-n_se) %>%
        mutate(
            percent = n / sum(n)
        ) %>%
        adorn_pct_formatting(rounding, , TRUE, percent)

    return(table)
}