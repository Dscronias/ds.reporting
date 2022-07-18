#' Oneway report of multiple categorical variables to Excel
#'
#' This functions returns a oneway table of a variable
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param var Variable to use
#' @param rounding number of digits for rounding (int)
#' @param hide_prct_char hides the "%" characted in the percentages column
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @return Count table
#' @export

svy_ow <- function(data, var, rounding = 2, hide_prct_char = TRUE, lang="en") {
    # Select header name
    prct <- switch(
        lang,
        "en" = "Percentage",
        "fr" = "Pourcentage",
        "math" = "%"
    )

    table <- data %>%
        survey_count({{var}}, .drop = FALSE) %>%
        select(-n_se) %>%
        mutate(
            {{prct}} := n / sum(n)
        ) %>%
        adorn_pct_formatting(rounding, , TRUE, {{prct}}) %>%
        {
            if (hide_prct_char)
                mutate(
                    .,
                    {{prct}} := str_remove(.data[[prct]], "%")
                )
            else .
        } %>%
        rename(N = n)

    return(table)
}