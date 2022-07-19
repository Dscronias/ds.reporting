#' Oneway report of multiple categorical variables to Excel, with survey design
#'
#' This functions creates an Excel workbook and exports a oneway table
#' of multiple variables with counts and percentages (with a survey design)
#'
#' @param Data dataframe
#' @param workbook name of the workbook (string)
#' @param worksheet name of the worksheet (string)
#' @param vars Variables to report (string)
#' @param rounding_n number of digits for rounding Ns (int)
#' @param rounding number of digits for rounding percentages (int)
#' @param cond_prct create a new percentage column conditionally, with % computed conditionally on a threshold (bool)
#' @param min_cond_prct Threshold to compute the new %s (int)
#' @param open_on_finish open excel file on finish (bool)
#' @param overwrite_file Overwrite existing file (bool)
#' @param lang Language of header (one of "en", "fr" or "math")
#' @param filename Name of excel file to export to (string)
#' @return Excel file with oneway table
#' @export

svy_ow_report <- function(data, workbook, worksheet, vars, 
    rounding_n = 0, rounding_prct = 2,
    cond_prct = FALSE, min_cond_prct = 5,
    cond_excluded_labels = NULL,
    data_label, label_from, label_to, open_on_finish = TRUE,
    overwrite_file = TRUE, lang="en", filename) {

    # SETUP & STYLES
    wb <- createWorkbook(workbook)
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = c("Top", "Bottom")
    )
    bs1 <- createStyle(
        border = "Bottom"
    )
    r_align <- createStyle(
        halign = "right"
    )
    indent_style <- createStyle(
        indent = 1
    )

    # Get var label
    if (!missing(data_label) & !missing(label_from) & !missing(label_to)) {
        get_var_label <- TRUE
    }

    # Counters
    row_counter <- 1

    ###########################################################################
    # HEADER

    # Build header
    table_header <- data %>%
        svy_ow(var = !!sym(vars[1]), cond_prct = cond_prct, min_cond_prct = min_cond_prct, lang=lang) %>%
        slice(0)

    ## Write header
    writeData(wb = wb, sheet = worksheet, x = table_header,
        startRow = row_counter,
    )
    ## Remove variable name in header
    writeData(wb = wb, sheet = worksheet, x = " ", startRow = row_counter)

    ## Header borders
    addStyle(wb, worksheet, cols = 1:length(table_header),
        rows = row_counter, style = hs1, stack = TRUE)
    ## Header right align
    addStyle(wb, worksheet, cols = 1:length(table_header),
        rows = row_counter, style = r_align, stack = TRUE)
    row_counter <- row_counter + 1
    ###########################################################################
    # TABLE CONTENT
    for (var in vars) {

        # Create table
        table <- data %>%
            svy_ow(var = !!sym(var),
                rounding_n = rounding_n,
                rounding_prct = rounding_prct,
                cond_prct = cond_prct,
                min_cond_prct = min_cond_prct,
                cond_excluded_labels = cond_excluded_labels,
                hide_prct_char = TRUE,
                lang=lang
            )

        # Table
        ## Write variable name
        ## Get variable label, if we have this information
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
        writeData(wb = wb, sheet = worksheet, x = var_label %>% as_character(),
            startRow = row_counter
        )
        mergeCells(wb, worksheet, cols = 1:length(table_header), rows = row_counter)
        row_counter <- row_counter + 1
        ## Write content
        writeData(wb, worksheet, table, startRow = row_counter,
            colNames = FALSE
        )
        ## Right align results
        addStyle(wb, worksheet, cols = 2:length(table_header),
            rows = row_counter:(row_counter + nrow(table) - 1),
            style = r_align, stack = TRUE, gridExpand = TRUE)
        ## Indent rowvar categories
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter:(row_counter + nrow(table) - 1),
            style = indent_style
        )
        row_counter <- row_counter + nrow(table)
    }

    ###########################################################################
    # FOOTER

    # Border to bottom of table
    addStyle(wb, worksheet, cols = 1:length(table_header),
        rows = row_counter-1, style = bs1, stack = TRUE)

    ###########################################################################
    # WRITE TABLE
    if (open_on_finish) {
        openXL(wb)
    }
    saveWorkbook(wb, filename, overwrite = overwrite_file)
}
