#' Tidy results to use with regression_report
#'
#' This function tidies the raw output of some regressions using
#' broom, shows robust estimates if required, and does some
#' formatting so the output can be used with regression_report()
#'
#' @param data Dataframe to use
#' @param robust Whether to include robust results (uses lmtest package)
#' @param robust_method Which robust method to use (see vcovHC in sandwich package)
#' @return Formatted regression table to use with regression_report
#' @export

regression_tidy_results <- function(data, exponentiate = FALSE, robust = FALSE, robust_method = "HC3") {
    
    data %>%
        # Tidy results + keep original results object
        tidy_and_attach(
            conf.int = TRUE,
            exponentiate = exponentiate
        ) %>%
        # Get reference categories
        tidy_add_reference_rows() %>%
        # Replace results with robust ones if robust == TRUE
        {
            if (robust) 
                # Getting rid of original estimates...
                select(., -estimate, -std.error, -statistic, -p.value, -conf.low, -conf.high) %>%
                # and left_joining the new ones
                left_join(
                    data %>%
                        coeftest(., vcov = vcovHC(., robust_method)) %>%
                        # NB: exponentiate is not supported by tidy for coeftest objects
                        tidy(
                            conf.int = TRUE
                        ) %>%
                        # Do it by hand then
                        {
                            if (exponentiate)
                                mutate(
                                    .,
                                    across(
                                        c("estimate", "conf.low", "conf.high"),
                                        ~ exp(.x)
                                    )
                                )
                            else
                                .
                        },
                    by = "term"
                )
            else
                .
        } %>%
        # Create new column with the label of the category
        mutate(
            category = str_remove(term, variable),
            .after = variable
        ) %>%
        # Confidence intervals formatting, stars (i.e. p-values) next to estimates, and formatted estimates
        mutate(
            conf_int = case_when(
                is.na(conf.low) & is.na(conf.high) ~ "-",
                TRUE ~ paste(
                    "[",
                    conf.low %>% round(2) %>% format(2) %>% as.character(),
                    "; ",
                    conf.high %>% round(2) %>% format(2) %>% as.character(),
                    "]",
                    sep = ""
                )
            ),
            estimate = case_when(
                is.na(estimate) ~ "Réf.",
                TRUE ~ estimate %>% round(2) %>% format(2)
            ),
            estimate = case_when(
                p.value < 0.001 ~ paste(estimate, "***", sep = ""),
                p.value < 0.01 ~ paste(estimate, "**", sep = ""),
                p.value < 0.05 ~ paste(estimate, "*", sep = ""),
                TRUE ~ estimate
            )
        )
}
