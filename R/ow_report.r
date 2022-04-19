#' Oneway report of multiple categorical variables to Excel
#'
#' This functions creates an Excel workbook and exports a oneway table
#' of multiple variables with counts and percentages
#'
#' @param Data dataframe
#' @param workbook name of the workbook (string)
#' @param worksheet name of the worksheet (string)
#' @param vars Variables to report (string)
#' @param rounding number of digits for rounding (int)
#' @param total_col add a total column (bool)
#' @param filename Name of excel file to export to
#' @return Excel file with oneway table
#' @export

ow_report <- function(data, workbook, worksheet, vars, rounding = 2,
    total = TRUE, filename) {

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

    # Counters
    row_counter <- 1

    ###########################################################################
    # HEADER

    # Build header
    table_header <- data %>%
        tabyl(!!sym(vars[1])) %>%
        as_tibble() %>%
        rename(N = n, Percent = percent) %>%
        #! Tabyl adds a column "valid_percent" if there are NAs
        { 
            if ("valid_percent" %in% colnames(.))
                rename(., `Percent (no NA)` = "valid_percent")
            else
                .
        } %>%
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
    addStyle(wb, worksheet, cols = 1:4,
        rows = row_counter, style = r_align, stack = TRUE)
    row_counter <- row_counter + 1
    ###########################################################################
    # TABLE CONTENT
    for (var in vars) {
        # Create table
        table <- data %>%
                    tabyl(!!sym(var)) %>%
                    adorn_pct_formatting(digits = rounding) %>%
                    # Delete "%" character
                    mutate(percent = str_remove(percent, "%")) %>%
                    {
                        if ("valid_percent" %in% colnames(.))
                            mutate(.,
                                valid_percent = str_remove(valid_percent, "%"),
                                !!sym(var) := fct_explicit_na(!!sym(var), "Missing values")
                            )
                        else
                            .
                    } %>%
                    as.tibble()

        if (total) {
            table <- bind_rows(
                data %>%
                    tabyl(!!sym(var)) %>%
                    adorn_totals() %>%
                    adorn_pct_formatting(2) %>%
                    mutate(percent = str_remove(percent, "%")) %>%
                    { 
                        if ("valid_percent" %in% colnames(.))
                            rename(., `Percent (no NA)` = "valid_percent")
                        else
                            .
                    } %>%
                    slice_tail(),
                table
            )
        }

        # Table
        ## Write variable name
        writeData(wb = wb, sheet = worksheet, x = var,
            startRow = row_counter
        )
        row_counter <- row_counter + 1
        ## Write content
        writeData(wb, worksheet, table, startRow = row_counter,
            colNames = FALSE
        )
        ## Right align results
        addStyle(wb, worksheet, cols = 2:4,
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
    openXL(wb)
    saveWorkbook(wb, filename, overwrite = TRUE)
}