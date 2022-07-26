#' Twoway report of multiple categorical variables to Excel, with survey design
#'
#' This functions creates an Excel workbook and exports a oneway table
#' of multiple variables with counts and percentages (with a survey design)
#'
#' @param data Dataframe to use
#' @param workbook name of the workbook (string)
#' @param worksheet name of the worksheet (string)
#' @param colvar Variable to use as the column variable
#' @param rowvars Variables to use as row variables
#' @param type_prct Type of percentage shown (one of "col", "row" or "cell")
#' @param rounding_prct number of digits for rounding Ns (int)
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param hide_prct_char hides the "%" characted in the percentages column
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @param data_label Dataframe of value labels from which to retrieve variable labels (optionnal)
#' @param label_from Column (in data_label) of variable names
#' @param label_to Column (in data_label) of variable labels
#' @param open_on_finish open excel file on finish (bool)
#' @param overwrite_file Overwrite existing file (bool)
#' @param filename Name of excel file to export to (string)
#' @return Excel file with twoway table
#' @export

svy_tw_report <- function(
    data,
    workbook,
    worksheet,
    colvar,
    rowvars,
    col_total = FALSE,
    row_total = FALSE,
    type_prct = "col",
    rounding_n = 0,
    rounding_prct = 2,
    hide_prct_char = TRUE,
    lang = "en",
    data_label,
    label_from,
    label_to,
    open_on_finish = TRUE,
    overwrite_file = TRUE,
    filename
) {
    ###########################################################################
    # Setup and styles
    wb <- createWorkbook(workbook)
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = "Top",
        valign = "center"
    )
    bs1 <- createStyle(
        border = "Bottom",
        valign = "center"
    )
    r_halign <- createStyle(
        halign = "right",
        valign = "center"
    )
    c_halign <- createStyle(
        halign = "center",
        valign = "center"
    )
    v_valign <- createStyle(
        valign = "center"
    )
    indent_style <- createStyle(
        indent = 1,
        valign = "center"
    )
    bold_font <- createStyle(
        textDecoration = "bold"
    )
    # Indicates, in the header, the type of %age used
    type_prct_label <-
        if (lang %in% c("en", "math") && type_prct == "row") "row" else
        if (lang %in% c("en", "math") && type_prct == "col") "column" else
        if (lang %in% c("en", "math") && type_prct == "cell") "row" else
        if (lang == "fr" && type_prct == "row") "ligne" else
        if (lang == "fr" && type_prct == "col") "colonne" else
        if (lang == "fr" && type_prct == "cell") "cellule" else
        "Problem at type_prct_label"
    # Various explanations in the footer
    footer_label <- switch(
        lang,
        "en" = "P-value: chi-squared test with Rao-Scott correction. P-values in bold indicate a result significant at the <0.05 level.",
        "fr" = "P-value : test du Khi-deux avec correction de Rao-Scott. Les p-values en gras indiquent des tests significatifs (< 0.05).",
        "math" = "P-value: chi-squared test with Rao-Scott correction. P-values in bold indicate a result significant at the <0.05 level."
    )

    # Get var label
    if (!missing(data_label) && !missing(label_from) && !missing(label_to)) {
        get_var_label <- TRUE
    } else {
        get_var_label <- FALSE
    }

    # Get colvar categories
    colvar_cats <- data %>%
        pull({{colvar}}) %>%
        levels()

    table_col_end <- 1 + (2 * length(colvar_cats)) + row_total * 2

    # Counters
    row_index <- 1

    ###########################################################################
    # HEADER
    # Length of top border: 1 + 2*length(colvar_cats)
    # Then : col variable name (or label) in column 2 to end of table (merged)
    # Then : Variable categories (size 2 columns, merged)
    # Then : N % N % (using rep)
    # Then border

    ## Top table borders
    addStyle(wb, worksheet, cols = 1:table_col_end,
        rows = row_index, style = hs1, stack = TRUE)

    ## Variable name
    if (get_var_label) {
        writeData(
            wb,
            worksheet,
            x = put_label(data = NULL,
                    var = {{colvar}},
                    data_label = data_label,
                    label_from = {{label_from}},
                    label_to = {{label_to}}
                ) %>% as.character(),
            startCol = 2,
            startRow = row_index
        )
    } else {
        writeData(wb, worksheet, x = englue("{{colvar}}"), startCol = 2, startRow = row_index)
    }
    mergeCells(wb, worksheet, cols = 2:table_col_end, rows = row_index)
    ## Top & bottom borders
    addStyle(wb, worksheet, cols = 2:table_col_end,
        rows = row_index, style = bs1, stack = TRUE)
    ## Center align
    addStyle(wb, worksheet, cols = 2:table_col_end,
        rows = row_index, style = c_halign, stack = TRUE)
    row_index <- row_index + 1

    ## Column variable labels
    ## Put labels
    col_index <- 2
    for (cat in colvar_cats) {
        writeData(wb, worksheet, x = cat, startCol = col_index, startRow = row_index)
        mergeCells(wb, worksheet, cols = col_index:(col_index + 1), rows = row_index)
        col_index <- col_index + 2
    }
    if (row_total) {
        writeData(wb, worksheet, x = "Total", startCol = col_index, startRow = row_index)
        mergeCells(wb, worksheet, cols = col_index:(col_index + 1), rows = row_index)
    }
    addStyle(wb, worksheet, cols = 2:table_col_end,
        rows = row_index, style = c_halign, stack = TRUE)
    row_index <- row_index + 1
    ### N & %
    writeData(wb, worksheet, x = (rep(c("N", glue("% ({type_prct_label})")), length(colvar_cats) + row_total * 1)) %>% t(), startCol = 2, startRow = row_index, colNames = FALSE)
    addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = bs1, stack = TRUE)
    addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = r_halign, stack = TRUE)
    row_index  <- row_index + 1

    ###########################################################################
    # TABLE CONTENT

    # Column total in very first row, if required
    if (col_total) {
        table_col_total <- data %>%
            svy_ow(
                {{colvar}},
                rounding_n = rounding_n,
                rounding_prct = rounding_prct,
                hide_prct_char = hide_prct_char,
                lang = "math"
            ) %>%
            pivot_wider(
                names_from = {{colvar}},
                values_from = c("N", "%")
            )

        for (cat in colvar_cats) {
            n_cat <- glue("N_{cat}")
            percentage_cat <- glue("%_{cat}")

            table_col_total <- table_col_total %>%
                relocate(
                    !!sym(percentage_cat),
                    .after = !!sym(n_cat)
                ) %>%
                mutate(
                    Total = "Total"
                ) %>%
                select(Total, everything())
        }
        writeData(wb = wb, sheet = worksheet, x = table_col_total,
            startRow = row_index, colNames = FALSE
        )
        addStyle(wb, worksheet, cols = 2:table_col_end,
            rows = row_index,
            style = r_halign, stack = TRUE, gridExpand = TRUE
        )
        row_index <- row_index + 1
    }

    for (var in rowvars) {
        # Create table
        table <- data %>%
            svy_tw(
                colvar = {{colvar}},
                rowvar = !!sym(var),
                type_prct = type_prct,
                row_total = row_total,
                rounding_n = rounding_n,
                rounding_prct = rounding_prct,
                hide_prct_char = hide_prct_char,
                lang = lang
            )

        # Get variable name or label
        if (get_var_label) {
            var_label <- put_label(
                data = NULL,
                var = !!sym(var),
                data_label = data_label,
                label_from = {{label_from}},
                label_to = {{label_to}}
            )
        } else {
            var_label <- var
        }

        # Write variable name or label in XL
        writeData(wb = wb, sheet = worksheet, x = var_label %>% as.character(),
            startRow = row_index
        )
        addStyle(wb, worksheet, cols = 1:table_col_end,
            rows = row_index, style = v_valign, stack = TRUE,
            gridExpand = TRUE
        )
        mergeCells(wb, worksheet, cols = 1:table_col_end, rows = row_index)
        row_index <- row_index + 1

        # Write table content
        writeData(wb, worksheet, table, startRow = row_index, colNames = FALSE)
        addStyle(wb, worksheet, cols = 2:table_col_end,
            rows = row_index:(row_index + nrow(table) - 1),
            style = r_halign, stack = TRUE, gridExpand = TRUE
        )
        addStyle(wb, worksheet, cols = 1,
            rows = row_index:(row_index + nrow(table) - 1),
            style = indent_style,
            stack = TRUE
        )
        row_index <- row_index + nrow(table)

        # Chi-squared test
        chisq <- data %>% svychisq(
            as.formula(glue("~ {englue('{{colvar}}')} + {var}")),
            .
        )

        if (chisq$p.value < 0.001) {
            pval <- "< 0.001"
        } else if (chisq$p.value < 0.05) {
            pval <- chisq$p.value %>% round(3)
        } else {
            pval <- chisq$p.value %>% round(2)
        }

        writeData(wb, worksheet, "P-value: ", startRow = row_index, colNames = FALSE)
        addStyle(wb, worksheet, cols = 1,
            rows = row_index,
            style = indent_style,
            stack = TRUE
        )
        writeData(wb, worksheet, pval %>% as.character(), startRow = row_index, startCol = 2, colNames = FALSE)
        if (pval < 0.05) {
            addStyle(wb, worksheet, cols = 2, rows = row_index, style = bold_font, stack = TRUE)
        }
        addStyle(wb, worksheet, cols = 2, rows = row_index, style = v_valign, stack = TRUE)
        row_index <- row_index + 1
    }

    ###########################################################################
    # FOOTER

    # Border to bottom of table
    addStyle(wb, worksheet, cols = 1:table_col_end,
        rows = row_index - 1, style = bs1, stack = TRUE)
    # Some additional explanations
    writeData(wb, worksheet, footer_label, startRow = row_index, startCol = 1, colNames = FALSE)
    addStyle(wb, worksheet, cols = 1, rows = row_index, style = v_valign, stack = TRUE)
    mergeCells(wb, worksheet, cols = 1:table_col_end, rows = row_index)

    ###########################################################################
    # WRITE TABLE
    if (open_on_finish) {
        openXL(wb)
    }
    saveWorkbook(wb, filename, overwrite = overwrite_file)
}