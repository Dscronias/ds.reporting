#' Oneway report of multiple categorical variables to Excel
#'
#' This functions returns a oneway table of a variable
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param var Variable to use
#' @param rounding_prct number of digits for rounding Ns (int)
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param hide_prct_char hides the "%" characted in the percentages column
#' @param no_na_prct Removes NAs from %ages calculations
#' @param cond_prct create a new percentage column on some specific, with % computed conditionally on a threshold (bool)
#' @param min_cond_prct Threshold to compute the new %s (int)
#' @param cond_excluded_labels Restrict the exclusion of %s calculations to specific value labels (list)
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @return Count table
#' @export

svy_ow <- function(data, var, rounding_n = 0, rounding_prct = 2, hide_prct_char = TRUE, 
no_na_prct = TRUE, cond_prct = FALSE, min_cond_prct = 0.05, cond_excluded_labels = NULL, lang = "en") {

    # Select header and NAs labels
    prct <- switch(
        lang,
        "en" = "Percentage",
        "fr" = "Pourcentage",
        "math" = "%"
    )
    prct_cond <- switch(
        lang,
        "en" = "Percentage*",
        "fr" = "Pourcentage*",
        "math" = "%*"
    )
    na_label <- switch(
        lang,
        "en" = "Missing values",
        "fr" = "Valeurs manquantes",
        "math" = "Missing values"
    )

    # Exclusions
    # If no excluded labels are specified, all labels will potentially be deleted
    if (is.null(cond_excluded_labels)) {
        cond_excluded_labels <- data %>%
            pull({{var}}) %>%
            levels()
    }

    # Create OW table
    table <- data %>%
        survey_count({{var}}, .drop = FALSE) %>%
        select(-n_se) %>%
        # Make missing values explicit
        mutate(
            {{var}} := fct_explicit_na({{var}}, na_level = na_label)
        ) %>%
        # Case when we do not want to calculate NAs percentages
        {if (no_na_prct)
            mutate(
                .,
                {{prct}} := case_when(
                    {{var}} != na_label ~ n / sum(n[which({{var}} != na_label)]),
                    TRUE ~ NaN
                )
            )
        else
            mutate(
                .,
                {{prct}} := n / sum(n)
            )
        }

    # CONDITIONAL %AGE COLUMN
    # New column with percentages computed only from
    # percentage >= [threshold]
    if (cond_prct) {
        # Get Ns from non_excluded labels
        n_cond <- table %>%
            # Delete NAs, if needed
            {if (no_na_prct)
                filter(
                    .,
                    !({{var}} == na_label)
                )
            else
                .
            } %>%
            filter(
                .,
                !({{var}} %in% cond_excluded_labels &
                !!sym(prct) < min_cond_prct)
            ) %>%
            pull(n)

        # Calculate new %s without unwanted categories
        table <- table %>%
            {if (no_na_prct)
                mutate(
                    .,
                    {{prct_cond}} := case_when(
                        !({{var}} == na_label) &
                        !(
                            {{var}} %in% cond_excluded_labels &
                            !!sym(prct) < min_cond_prct
                        ) ~ (n / sum(n_cond)),
                        TRUE ~ NaN
                    )
                )
            else
                mutate(
                    .,
                    {{prct_cond}} := case_when(
                        !(
                            {{var}} %in% cond_excluded_labels &
                            !!sym(prct) < min_cond_prct
                        ) ~ (n / sum(n_cond)),
                        TRUE ~ NaN
                    )
                )
            }
    }

    # Ns and %ages formatting
    table <- table %>%
        mutate(
            n = round(n, rounding_n)
        ) %>%
        rename(N = n) %>%
        adorn_pct_formatting(., rounding_prct, , TRUE, {{prct}}) %>%
            {
                if (hide_prct_char)
                    mutate(
                        .,
                        {{prct}} := str_remove(.data[[prct]], "%")
                    )
                else .
            } %>%
            {
                if (cond_prct)
                    adorn_pct_formatting(., rounding_prct, , TRUE, {{prct_cond}}) %>%
                    {
                        if (hide_prct_char)
                            mutate(
                                .,
                                {{prct_cond}} := str_remove(.data[[prct_cond]], "%")
                            )
                        else .
                    }
                else .
            }
    return(table)
}