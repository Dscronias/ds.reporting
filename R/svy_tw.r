#' Twoway report of multiple categorical variables to Excel
#'
#' This functions returns a twoway table of one column variable and one row variable
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param colvar Variable to use as the column variable
#' @param rowvar Variable to use as the row variable
#' @param type_prct Type of percentage shown (one of "col", "row" or "cell")
#' @param row_total Add total column for row variable
#' @param rounding_prct number of digits for rounding Ns (int)
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param hide_prct_char hides the "%" characted in the percentages column
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @return Count table
#' @export

svy_tw <- function(
    data,
    colvar,
    rowvar,
    type_prct = "col",
    row_total = FALSE,
    rounding_n = 0,
    rounding_prct = 2,
    hide_prct_char = TRUE,
    lang = "en"
) {
    # Header label
    prct <- switch(
        lang,
        "en" = "Percentage",
        "fr" = "Pourcentage",
        "math" = "%"
    )

    # Table
    # Switch case depending on percentage type (col, row or cell)
    table <- switch(
        type_prct,
        "cell" = data %>%
            group_by({{colvar}}, {{rowvar}}) %>%
            summarise(
                N = survey_total()
            ) %>%
            ungroup() %>%
            mutate(
                {{prct}} := N / sum(N)
            ),
        "col" = data %>%
            group_by({{colvar}}, {{rowvar}}) %>%
            summarise(
                N = survey_total(),
                {{prct}} := survey_prop(),
            ) %>%
            ungroup(),
        "row" = data %>%
            group_by({{rowvar}}, {{colvar}}) %>%
            summarise(
                N = survey_total(),
                {{prct}} := survey_prop(),
            ) %>%
            ungroup()
)

    # Some table formatting
    table <- table %>%
        # Show empty categories
        complete({{rowvar}}, {{colvar}}, fill = list(N = 0)) %>%
        # Percentage column as %age
        adorn_pct_formatting(rounding_prct, , TRUE, {{prct}}) %>%
        {
            if (hide_prct_char)
                mutate(
                    .,
                    {{prct}} := str_remove(.data[[prct]], "%")
                )
            else .
        } %>%
        mutate(N = round(N, rounding_n)) %>%
        select(-N_se) %>%
        {if (type_prct %in% c("col", "row"))
            select(., -glue("{prct}_se"))
        else .} %>%
        pivot_wider(
            names_from = {{colvar}},
            values_from = c("N", all_of(prct))
        )

    # Reorganize columns : N % N % instead of N N % %
    for (cat in data %>% pull({{colvar}}) %>% levels()) {
        n_cat <- glue("N_{cat}")
        percentage_cat <- glue("{prct}_{cat}")

        table <- table %>%
            relocate(
                !!sym(percentage_cat),
                .after = !!sym(n_cat)
            )
    }

    # Add total column
    if (row_total) {
        total_prct <- paste0(prct, "_Total")

        table_row_total <- data %>%
            group_by({{rowvar}}) %>%
            summarise(
                N_Total = survey_total(),
                {{total_prct}} := survey_prop()
            ) %>%
            select(-N_Total_se, -glue("{total_prct}_se")) %>%
            adorn_pct_formatting(rounding_prct, , TRUE, {{total_prct}}) %>%
            {
            if (hide_prct_char)
                mutate(
                    .,
                    {{total_prct}} := str_remove(.data[[total_prct]], "%")
                )
            else .
            } %>%
            mutate(N_Total = round(N_Total, rounding_n))

        table <- table %>%
            left_join(
                table_row_total,
                by = englue("{{rowvar}}")
            )
    }

    return(table)
}