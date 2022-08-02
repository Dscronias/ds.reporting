#' Report of a quantitative variable by multiple groups, with t-tests
#'
#' This functions returns a by-group summary of means for a continuous variable (with t-test of difference in means using a regression)
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param workbook name of the workbook (string)
#' @param new_wb create new workbook (or open filename)
#' @param worksheet name of the worksheet (string)
#' @param quant_var Quantitative variable to use
#' @param by_vars Variable to use for group_by
#' @param rounding_mean number of digits for rounding mean (int)
#' @param rounding_se number of digits for rounding mean standard error (int)
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param rounding_ci hides the "%" characted in the percentages column
#' @param show_se shows standard error of the mean
#' @param show_ci shows the confidence interval of the mean
#' @param lang changes some presentation texts (depending on your language). Three values : "en" (default), "fr", or "math"
#' @param data_label Dataframe of value labels from which to retrieve variable labels (optionnal)
#' @param label_from Column (in data_label) of variable names
#' @param label_to Column (in data_label) of variable labels
#' @param open_on_finish open excel file on finish (bool)
#' @param overwrite_file Overwrite existing file (bool)
#' @param filename Name of excel file to export to (string)
#' @return Table with means, SE, CI and p-values
#' @export

svy_bygroup_mean_report <- function(data, workbook, new_wb = TRUE, worksheet, quant_var, by_vars, rounding_mean = 2, rounding_se = 2, rounding_prct = 2, rounding_ci = 2, show_se = FALSE, show_ci = TRUE, lang = "en", data_label, label_from, label_to, open_on_finish = FALSE, overwrite_file = TRUE, filename) {
    ###########################################################################
    # Setup and styles
    if (new_wb) {
        wb <- createWorkbook(workbook)
    } else {
        wb <- loadWorkbook(filename)
    }
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = c("Top", "Bottom"),
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
    footer_label <- switch(
        lang,
        "en" = "P-value: t-test of difference in means between the category and the reference category (weighted linear regression). P-values in bold indicate a result significant at the <0.05 level.",
        "fr" = "P-value : t-test de comparaison des moyennes entre le groupe en ligne et le groupe de référence (régression linéaire pondérée). Les p-values en gras indiquent des tests significatifs (< 0.05).",
        "math" = "P-value: t-test of difference in means between the category and the reference category (weighted linear regression). P-values in bold indicate a result significant at the <0.05 level."
    )

    # Get var label
    if (!missing(data_label) && !missing(label_from) && !missing(label_to)) {
        get_var_label <- TRUE
    } else {
        get_var_label <- FALSE
    }

    # Get last column number
    # 4 columns are always there : variable, proportion, mean, p-val*
    # 2 optional : SE and CI
    table_col_end <- 4 + 1 * show_se + 1 * show_ci

    # Counters
    row_index <- 1

    ###########################################################################
    # HEADER

    # Build header
    table_header <- data %>%
        svy_bygroup_mean(
            {{quant_var}},
            !!sym(by_vars[1]),
            show_se = show_se,
            show_ci = show_ci,
            lang = lang
        ) %>%
        slice(0)
    ## Write
    writeData(
        wb = wb,
        sheet = worksheet,
        x = table_header,
        startRow = row_index,
    )

    ## Set continuous variable name (or label)
    if (get_var_label) {
        writeData(
            wb = wb,
            sheet = worksheet,
            x = put_label(
                    data = NULL,
                    var = {{quant_var}},
                    data_label = data_label,
                    label_from = {{label_from}},
                    label_to = {{label_to}}
                ) %>% as.character(),
            startCol = 1,
            startRow = row_index
        )
    } else {
        writeData(
            wb,
            worksheet,
            x = englue("{{quant_var}}"),
            startCol = 1,
            startRow = row_index
        )
    }

    ## Borders
    addStyle(
        wb,
        worksheet,
        cols = 1:table_col_end,
        rows = row_index,
        style = hs1,
        stack = TRUE
    )

    ## Increment row
    row_index <- row_index + 1

    ###########################################################################
    # TABLE CONTENT

    for (var in by_vars) {

        # Table
        table <- data %>%
        svy_bygroup_mean(
            {{quant_var}},
            !!sym(var),
            rounding_mean = rounding_mean,
            rounding_se = rounding_se,
            rounding_prct = rounding_prct,
            rounding_ci = rounding_ci,
            show_se = show_se,
            show_ci = show_ci,
            lang = lang
        ) %>%
        # P-val formatting
        mutate(
            pval_true = `P-value*`,
             `P-value*` = case_when(
                `P-value*` < 0.001 ~ "< 0.001",
                `P-value*` < 0.05 ~ round(`P-value*`, 3) %>% format(nsmall = 3) %>% as.character(),
                `P-value*` >= 0.05 ~ round(`P-value*`, 2) %>% format(nsmall = 2) %>% as.character(),
                is.na(`P-value*`) ~ "-"
            )
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
        writeData(wb, worksheet, table %>% select(-pval_true), startRow = row_index, colNames = FALSE)
        addStyle(wb, worksheet, cols = 2:table_col_end,
            rows = row_index:(row_index + nrow(table) - 1),
            style = r_halign, stack = TRUE, gridExpand = TRUE
        )
        addStyle(wb, worksheet, cols = 1,
            rows = row_index:(row_index + nrow(table) - 1),
            style = indent_style,
            stack = TRUE
        )

        # Significant p-values in bold
        pval_list <- table %>% pull(pval_true)
        row_index_pval <- row_index
        for (el in pval_list) {
            if (!is.na(el) & el < 0.05) {
                addStyle(
                    wb, worksheet, cols = table_col_end,
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