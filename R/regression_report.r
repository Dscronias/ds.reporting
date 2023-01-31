#' Regression table report to XL
#'
#' This functions creates an Excel workbook and exports a regression table
#'
#' @param data Dataframe to use
#' @param workbook name of the workbook (string)
#' @param new_wb create new workbook (or open filename)
#' @param worksheet name of the worksheet (string)
#' @param depvar Dependent variable name (used to retrieve its label)
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @param data_label Dataframe of value labels from which to retrieve variable labels (optionnal)
#' @param label_from Column (in data_label) of variable names
#' @param label_to Column (in data_label) of variable labels
#' @param open_on_finish open excel file on finish (bool)
#' @param overwrite_file Overwrite existing file (bool)
#' @param filename Name of excel file to export to (string)
#' @return Excel file with twoway table
#' @export

regression_report <- function(data, workbook, new_wb = TRUE, worksheet, model_label = "Model", lang = "en", depvar, data_label, label_from, label_to, open_on_finish = TRUE, overwrite_file = TRUE, filename) {
    
    ###########################################################################
    # Setup and styles
    if (new_wb) {
        wb <- createWorkbook(workbook)
    } else {
        wb <- loadWorkbook(filename)
    }
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

    # Various explanations in the footer
    pval_stars <- attr(data, "pval_stars")
    footer_label <- switch(
        lang,
        "en" = glue::glue("***: p-value < {pval_stars[1]}; **: p-value < {pval_stars[2]}; *: p-value < {pval_starts[3]}. Estimates in bold are statistically significant (p-value < 0.05)."),
        "fr" = glue::glue("***: p-value < {pval_stars[1]}; **: p-value < {pval_stars[2]}; *: p-value < {pval_starts[3]}. Les estimations en gras sont statistiquement significatives (p-value < 0.05)."),
        "math" = glue::glue("***: p-value < {pval_stars[1]}; **: p-value < {pval_stars[2]}; *: p-value < {pval_starts[3]}. Estimates in bold are statistically significant (p-value < 0.05).")
    )

    # Get var label
    if (!missing(data_label) && !missing(label_from) && !missing(label_to)) {
        get_var_label <- TRUE
    } else {
        get_var_label <- FALSE
    }

    # Last column pos
    table_col_end <- 3

    # Row index
    row_index <- 1

    ###########################################################################
    # HEADER

    ## Top table borders
    addStyle(wb, worksheet, cols = 1:table_col_end,
        rows = row_index, style = hs1, stack = TRUE)

    ## Variable name
    if (!missing(depvar)) {
        if (get_var_label) {
            writeData(
                wb,
                worksheet,
                x = put_label(
                    data = NULL,
                    var = {{depvar}},
                    data_label = data_label,
                    label_from =  {{label_from}},
                    label_to = {{label_to}}
                ) %>% as.character()
            )
        } else {
            writeData(
                wb,
                worksheet,
                x = englue('{{depvar}}'),
                startCol = 1,
                startRow = row_index
            )
        }
    }
    mergeCells(
        wb, worksheet,
        cols = 1,
        rows = row_index:(row_index + 1)
    )
    addStyle(wb, worksheet, cols = 1,
        rows = row_index, style = v_valign, stack = TRUE
    )


    ## Model title
    writeData(
        wb,
        worksheet,
        x = model_label,
        startCol = 2,
        startRow = row_index
    )
    mergeCells(
        wb, worksheet,
        cols = 2:table_col_end,
        rows = row_index
    )
    addStyle(wb, worksheet, cols = 2:table_col_end,
        rows = row_index, style = c_halign, stack = TRUE
    )
    row_index <- row_index + 1


    ## Coef + CI
    writeData(
        wb,
        worksheet,
        x = c("Coef.", "95% CI") %>% t(),
        startCol = 2,
        startRow = row_index,
        colNames = FALSE
    )
    addStyle(wb, worksheet, cols = 2:table_col_end,
        rows = row_index, style = c_halign, stack = TRUE
    )

    ## Varlist
    varlist <- data %>%
        distinct(variable) %>%
        filter(variable != "(Intercept)") %>%
        pull()

    ## Bottom header border
    addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = bs1, stack = TRUE)
    row_index <- row_index + 1
    
    ###########################################################################
    # TABLE CONTENT

    for (var in varlist) {
        table <- data %>% 
            filter(variable == var) %>%
            select(category, estimate, conf_int, p.value)

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
        writeData(
            wb,
            worksheet,
            table %>% select(-p.value),
            startRow = row_index,
            colNames = FALSE
        )
        addStyle(wb, worksheet, cols = 2:table_col_end,
            rows = row_index:(row_index + nrow(table) - 1),
            style = c_halign, stack = TRUE, gridExpand = TRUE
        )
        addStyle(wb, worksheet, cols = 1,
            rows = row_index:(row_index + nrow(table) - 1),
            style = indent_style,
            stack = TRUE
        )
        
        # Significant p-values in bold
        pval_list <- table %>% pull(p.value)
        row_index_pval <- row_index
        for (el in pval_list) {
            if (!is.na(el) & el < pval_stars[3]) {
                addStyle(
                    wb, worksheet, cols = 2:table_col_end,
                    rows = row_index_pval,
                    style = bold_font,
                    stack = TRUE
                )
            }
            row_index_pval <- row_index_pval + 1
        }
        row_index <- row_index + nrow(table)
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
