#' Twoway table
#'
#' @description
#' This functions returns a tibble with one or multiple two-way tables between one column variable and multiple row variables
#'
#' @param data Dataframe to use
#' @param colvar Variable to use as the column variable
#' @param rowvar list of variables to use as row variables
#' @param type_prct Type of percentage shown (one of "col", "row" or "cell")
#' @param add_row_total Add total column for row variable
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param missing include NAs from the table (bool)
#' @param chisq_n_thres Minimum n for chi-squared test, if >= 1. Minimum percentage for chi-squared test, if < 1
#' @param chisq_show_excluded Show which variables have excluded categories
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @return Tibble with one or multiple two-way tables
#' @export

tw <- function(
    data,
    colvar,
    rowvars,
    type_prct = "col",
    add_row_total = FALSE,
    add_col_total = FALSE,
    rounding_prct = 2,
    missing = TRUE,
    chisq_n_thres = 5,
    chisq_show_excluded = TRUE,
    lang = "en"
) {
    # Language-dependent labels
    ## Missing values label
    na_label <- switch(
        lang,
        "en" = "Missing values",
        "fr" = "Valeurs manquantes",
        "math" = "Missing values"
    )

    # Results dataframe
    tw_result <- data %>%
        dplyr::select(dplyr::all_of(rowvars)) %>%
        purrr::imap_dfr(
            ~ tw_compute(
                data = data,
                colvar = {{colvar}},
                rowvar = !!rlang::sym(.y),
                type_prct = type_prct,
                add_row_total = add_row_total,
                rounding_prct = rounding_prct,
                missing = missing,
                chisq_n_thres = chisq_n_thres,
                chisq_show_excluded = chisq_show_excluded,
                na_label = na_label
            ),
        ) %>%
        `attr<-`(
            "colvar_categories",
            data %>%
                dplyr::pull({{colvar}}) %>%
                {if (missing) forcats::fct_explicit_na(., na_level = na_label) else .} %>%
                levels() %>%
                {if (add_row_total) append(., "Total") else .}
        ) %>%
        `attr<-`("report_type", "tw") %>%
        `attr<-`("missing", missing) %>%
        `attr<-`("lang", lang) %>%
        `attr<-`("chisq_rejection_threshold", chisq_n_thres) %>%
        `attr<-`("type_prct", type_prct) %>%
        `attr<-`("has_row_totals", add_row_total)

    # Add counts for column variable in a new row
    if (add_col_total) {
        tw_result <- tw_add_col_total(data, tw_result, missing, rounding_prct, {{colvar}}, na_label)
    }

    return(tw_result)
}

tw_compute <- function(
    data,
    colvar,
    rowvar,
    type_prct,
    add_row_total,
    rounding_prct,
    missing,
    chisq_n_thres,
    chisq_show_excluded,
    na_label
) {

    # Base count table
    table <- data %>%
        dplyr::count({{colvar}}, {{rowvar}}) %>%
        {
            if (!missing)
                tidyr::drop_na(.)
            else
                dplyr::mutate(
                    .,
                    {{colvar}} := {{colvar}} %>% forcats::fct_explicit_na(na_level = na_label),
                    {{rowvar}} := {{rowvar}} %>% forcats::fct_explicit_na(na_level = na_label)
                )
        } %>%
        tidyr::complete({{colvar}}, {{rowvar}}, fill = list(n = 0)) %>%
        dplyr::mutate(prct = n / sum(n))

    # Chi-squared test
    ## Display variables rejected categories
    chisq_table_rejected <- table %>%
        { if (chisq_n_thres >= 1)
            dplyr::filter(., n < chisq_n_thres)
        else
            dplyr::filter(., prct < chisq_n_thres)
        }
    if (nrow(chisq_table_rejected) > 0 && chisq_show_excluded) {
        print(glue::glue("[{rlang::englue('{{rowvar}}')}] Some categories were excluded"))
    }

    ## Calculate test and get p-value
    chisq_p <- table %>%
        { if (chisq_n_thres >= 1)
            dplyr::filter(., n >= chisq_n_thres)
        else
            dplyr::filter(., prct >= chisq_n_thres)
        } %>%
        dplyr::select(!prct) %>%
        tidyr::pivot_wider(
            names_from = {{colvar}},
            values_from = n
        ) %>%
        dplyr::select(!{{rowvar}})
    ### NAs in the chi-squared matrix will output an error
    if (any(is.na(chisq_p))) {
      message(
        glue::glue(
        "[{rlang::englue('{{rowvar}}')}] Chi-squared test matrix has NAs. P-value will be NA. Modifying the chisq_n_thres (or setting it to 0) will correct this."
        )
      )
      chisq_p <- NA
    } else {
      chisq_p <- chisq_p %>%
        chisq.test() %>%
        broom::glance() %>%
        dplyr::pull(p.value)
    }

    # Results formatting
    ## Percentage type
    table <- switch(
        type_prct,
        "cell" = table,
        "col" = table %>%
            dplyr::group_by({{colvar}}) %>%
            dplyr::mutate(prct = (n / sum(n)) %>% tidyr::replace_na(0)),
        "row" = table %>%
            dplyr::group_by({{rowvar}}) %>%
            dplyr::mutate(prct = (n / sum(n)) %>% tidyr::replace_na(0))
    ) %>%
    dplyr::ungroup()

    # %ages formatting
    table <- table %>%
        dplyr::mutate(prct = (prct * 100) %>% round(rounding_prct) %>% format(rounding_prct)) %>%
        tidyr::pivot_wider(
            names_from = {{colvar}},
            values_from = c("n", "prct")
        )

    # Reordering columns : N N % % ==> N % N %
    for (cat in data %>% dplyr::pull({{colvar}}) %>% levels()) {
        n_cat <- glue::glue("n_{cat}")
        percentage_cat <- glue::glue("prct_{cat}")

        table <- table %>%
            dplyr::relocate(
                !!rlang::sym(percentage_cat),
                .after = !!rlang::sym(n_cat)
            )
    }

    #Add total column
    if (add_row_total) {
        table <- table %>%
            dplyr::mutate(
                n_Total = rowSums(dplyr::select(., tidyselect::starts_with("n_"))),
                prct_Total = ((n_Total / sum(n_Total)) * 100) %>% round(rounding_prct) %>% format(rounding_prct)
            )
    }

    # List formatting
    table <- list(table)
    names(table) <- rlang::englue("{{rowvar}}")
    chisq_table_rejected <- list(chisq_table_rejected)
    names(chisq_table_rejected) <- rlang::englue("{{rowvar}}")

    # Results to tibble
    return(
        tibble::tibble(
            var_col = rlang::englue("{{colvar}}"),
            var_row = rlang::englue("{{rowvar}}"),
            tw_table = table,
            excluded_cat = chisq_table_rejected,
            p_value = chisq_p
        )
    )
}

tw_add_col_total <- function(data, tw_result, missing, rounding_prct, colvar, na_label) {
    # Adds a new row to the results dataframe
    # With the counts of the column variable
    col_total <- data %>%
        {
            if (missing)
                dplyr::mutate(., {{colvar}} := {{colvar}} %>% forcats::fct_explicit_na(na_level = na_label))
            else
                tidyr::drop_na(., {{colvar}})
        } %>%
        dplyr::count({{colvar}}) %>%
        dplyr::mutate(prct = ((n/sum(n)) * 100) %>% round(rounding_prct) %>% format(rounding_prct)) %>%
        tidyr::pivot_wider(
            names_from = {{colvar}},
            values_from = c("n", "prct")
        )

    # Reordering columns : N N % % ==> N % N %
    for (cat in data %>% dplyr::pull({{colvar}}) %>% levels()) {
        n_cat <- glue::glue("n_{cat}")
        percentage_cat <- glue::glue("prct_{cat}")

        col_total <- col_total %>%
            dplyr::relocate(
                !!rlang::sym(percentage_cat),
                .after = !!rlang::sym(n_cat)
            )
    }

    col_total <- list(col_total)
    names(col_total) <- rlang::englue("{{colvar}}")

    # Add new row to tw_results
    return(tw_result %>%
        dplyr::add_row(
            var_col = rlang::englue("{{colvar}}"),
            var_row = "total",
            tw_table = col_total,
            excluded_cat = NA,
            p_value = NA
        )
    )
}