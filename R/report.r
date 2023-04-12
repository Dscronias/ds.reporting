#' Export report (for tw tables only for now)
#'
#' @description
#' This functions exports an excel file of a tw table
#'
#' @param data Dataframe to use
#' @param workbook name of the workbook (string)
#' @param worksheet name of the worksheet (string)
#' @param label_data Dataframe of value labels from which to retrieve variable labels (optional)
#' @param label_from Column (in data_label) of variable names
#' @param label_to Column (in data_label) of variable labels
#' @param open_on_finish open excel file on finish (bool)
#' @param append_to_existing_file Append report ton an existing excel file
#' @param filename Name of excel file to export to (string)
#' @return An excel file + the original table
#' @export

report <- function(data, workbook, worksheet, label_data, label_from, label_to, open_on_finish = FALSE, append_to_existing_file = FALSE, filename) {
    # Report expects a tibble with a "report_type" attributes indicating the reporting function used to create the input data ("tw"...)
    # Throws an error if the report_type attribute is missing
    if (is.null(attributes(data)$report_type)) stop(glue::glue("attributes({deparse(substitute(data))})$report_type is missing"))

    switch(
        attributes(data)$report_type,
        "tw" = report.tw(
            data = data,
            workbook = workbook,
            worksheet = worksheet,
            label_data = label_data,
            label_from = {{label_from}},
            label_to = {{label_to}},
            open_on_finish = open_on_finish,
            append_to_existing_file = append_to_existing_file,
            filename = filename
        ),
        stop(glue::glue("Input data not supported (attributes({deparse(substitute(data))})$report_type not recognized)"))
    )
}

report.tw <- function(data, workbook, worksheet, label_data, label_from, label_to, open_on_finish, append_to_existing_file, filename) {

    # SETUP ###################################################################
    ## Init
    if (append_to_existing_file) {
        wb <- openxlsx::loadWorkbook(filename)
    } else {
        wb <- openxlsx::createWorkbook(workbook)
    }
    openxlsx::addWorksheet(wb, worksheet)
    row_index <- 1


    ## Styles
    border_top <- openxlsx::createStyle(
        border = "Top"
    )
    border_bottom <- openxlsx::createStyle(
        border = "Bottom"
    )
    align_horizontal_right <- openxlsx::createStyle(
        halign = "right",
    )
    align_horizontal_center <- openxlsx::createStyle(
        halign = "center",
    )
    align_vertical_center <- openxlsx::createStyle(
        valign = "center",
    )
    indent_one <- openxlsx::createStyle(
        indent = 1
    )
    font_bold <- openxlsx::createStyle(
        textDecoration = "bold"
    )
    # Get variables labels
    if (!missing(label_data) && !missing(label_from) && !missing(label_to)) {
        data <- data %>%
            dplyr::left_join(
                label_data %>% dplyr::select({{label_from}}, {{label_to}}),
                by = c("var_row" = rlang::englue("{{label_from}}"))
            )
        label_colvar <- label_data %>%
            dplyr::filter(
                {{label_from}} == data %>% dplyr::distinct(var_col) %>% dplyr::pull()
            ) %>%
            dplyr::pull({{label_to}})
    } else {
        label_colvar <- data %>% dplyr::distinct(var_col) %>% dplyr::pull()
    }

    # Prct label
    type_prct_label <-
        if (attr(data, "lang") %in% c("en", "math") && attr(data, "type_prct") == "row") "row" else
        if (attr(data, "lang") %in% c("en", "math") && attr(data, "type_prct") == "col") "column" else
        if (attr(data, "lang") %in% c("en", "math") && attr(data, "type_prct") == "cell") "cell" else
        if (attr(data, "lang") == "fr" && attr(data, "type_prct") == "row") "ligne" else
        if (attr(data, "lang") == "fr" && attr(data, "type_prct") == "col") "colonne" else
        if (attr(data, "lang") == "fr" && attr(data, "type_prct") == "cell") "cellule" else
        "Problem at type_prct_label"

    # Explanations in footer
    footer_label <- switch(
        attr(data, "lang"),
        "en" = "P-value: chi-squared test. P-values in bold indicate a result significant at the <0.05 level.",
        "fr" = "P-value : test du Khi-deux. Les p-values en gras indiquent des tests significatifs (< 0.05).",
        "math" = "P-value: chi-squared test. P-values in bold indicate a result significant at the <0.05 level."
    )

    # Last column
    table_col_end <- 1 + 2 * length(attr(data, "colvar_categories"))

    # HEADER ##################################################################
    ## Top border
    openxlsx::addStyle(wb, worksheet, cols = 1:table_col_end,
        rows = row_index, style = border_top, stack = TRUE)
    ## Colvar label
    openxlsx::writeData(wb, worksheet, x = label_colvar, startCol = 2, startRow = row_index)
    openxlsx::mergeCells(wb, worksheet, cols = 2:table_col_end, rows = row_index)
    ### Center align
    openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_horizontal_center, stack = TRUE)
    openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_vertical_center, stack = TRUE)
    ### Bottom border
    openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = border_bottom, stack = TRUE)
    row_index <- row_index + 1
    ## Colvar categories
    col_index <- 2
    for (cat in attr(data, "colvar_categories")) {
        openxlsx::writeData(wb, worksheet, x = cat, startCol = col_index, startRow = row_index)
        openxlsx::mergeCells(wb, worksheet, cols = col_index:(col_index + 1), rows = row_index)
        col_index <- col_index + 2
    }
    openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_horizontal_center, stack = TRUE)
    openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_vertical_center, stack = TRUE)
    row_index <- row_index + 1
    ## N & %
    openxlsx::writeData(wb, worksheet, x = (rep(c("N", glue::glue("% ({type_prct_label})")), length(attr(data, "colvar_categories")))) %>% t(), startCol = 2, startRow = row_index, colNames = FALSE)
    openxlsx::addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = border_bottom, stack = TRUE)
    openxlsx::addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = align_horizontal_center, stack = TRUE)
    row_index  <- row_index + 1

    ###########################################################################
    # TABLE CONTENT

    ## Colvar totals
    if (data %>% dplyr::filter(var_row == "total") %>% nrow() == 1) {
        data_current_row <- data %>% dplyr::filter(var_row == "total")
        data_current_var <- data_current_row %>% dplyr::pull(var_col)

        openxlsx::writeData(wb = wb, sheet = worksheet, x = "Total", startRow = row_index)
        openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_vertical_center, stack = TRUE)
        openxlsx::writeData(
            wb = wb,
            sheet = worksheet,
            x = data_current_row %>% dplyr::pull(tw_table) %>% .[[data_current_var]],
            startRow = row_index,
            startCol = 2,
            colNames = FALSE
        )
        openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_vertical_center, stack = TRUE)
        openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index, style = align_horizontal_center, stack = TRUE)
        row_index <- row_index + 1
    }

    ## Rowvars
    for (var in data %>% filter(var_row != "total") %>% dplyr::pull(var_row)) {
        data_current_row <- data %>% dplyr::filter(var_row == var)

        ### Row variable label
        if (!missing(label_data) && !missing(label_from) && !missing(label_to)) {
            openxlsx::writeData(wb = wb, sheet = worksheet, x = data_current_row %>% dplyr::pull({{label_to}}), startRow = row_index)
        } else {
            openxlsx::writeData(wb = wb, sheet = worksheet, x = data_current_row %>% dplyr::pull(var_row), startRow = row_index)
        }
        openxlsx::mergeCells(wb, worksheet, cols = 1:table_col_end, rows = row_index)
        openxlsx::addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = align_vertical_center, stack = TRUE)
        row_index <- row_index + 1

        ### Table contents
        data_current_row_table <- data_current_row %>% dplyr::pull(tw_table) %>% .[[var]]
        openxlsx::writeData(wb = wb, sheet = worksheet, x = data_current_row_table, startRow = row_index, colNames = FALSE)
        openxlsx::addStyle(wb, worksheet, cols = 1, rows = row_index:(row_index+nrow(data_current_row_table)), style = indent_one, stack = TRUE)
        openxlsx::addStyle(wb, worksheet, cols = 2:table_col_end, rows = row_index:(row_index+nrow(data_current_row_table)), style = align_horizontal_center, stack = TRUE, gridExpand = TRUE)
        row_index <- row_index + nrow(data_current_row_table)

        ### P-value
        data_current_row_p <- data_current_row %>% dplyr::pull(p_value)
        if (is.na(data_current_row_p)) {
            data_current_row_p_fmt <- "NA"
        } else if (data_current_row_p < 0.001) {
            data_current_row_p_fmt <- "< 0.001"
        } else if (data_current_row_p < 0.05) {
            data_current_row_p_fmt <- data_current_row_p %>% round(3) %>% format(3)
        } else {
            data_current_row_p_fmt <- data_current_row_p %>% round(2) %>% format(2)
        }
        openxlsx::writeData(wb = wb, sheet = worksheet, x = "P-value", startRow = row_index, colNames = FALSE)
        openxlsx::writeData(wb = wb, sheet = worksheet, x = data_current_row_p_fmt, startRow = row_index, startCol = 2, colNames = FALSE)
        openxlsx::addStyle(wb, worksheet, cols = 1, rows = row_index, style = indent_one, stack = TRUE)
        if (!is.na(data_current_row_p) && data_current_row_p < 0.05) {
            openxlsx::addStyle(wb, worksheet, cols = 2, rows = row_index, style = font_bold, stack = TRUE)
        }

        ### Next variable
        row_index <- row_index + 1
    }

    ###########################################################################
    # FOOTER

    ## Bottom border
    openxlsx::addStyle(wb, worksheet, cols = 1:table_col_end, rows = row_index, style = border_top, stack = TRUE)
    openxlsx::writeData(wb, worksheet, footer_label, startRow = row_index)
    openxlsx::mergeCells(wb, worksheet, 1:table_col_end, row_index)

    ###########################################################################
    # WRITE TABLE
    if (open_on_finish) {
        openxlsx::openXL(wb)
    }
    openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)

    return(data)
}

#####################################################################
