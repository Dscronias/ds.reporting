#' Twoway report of multiple categorical variables to Excel
#'
#' This functions creates an Excel workbook and exports a twoway table
#' of one column variable and multiple row variables, with chi-squared tests.
#'
#'  @param Data dataframe
#'  @param workbook name of the workbook (string)
#'  @param worksheet name of the worksheet (string)
#'  @param colvar name of column variable (string, max one)
#'  @param rowvars names of row variables (vector of strings)
#'  @param rounding number of digits for rounding
#'  @param percentage_style "row", "col" or "both" (= cell) percentages
#'  @param counts show counts in parenthesis
#'  @param separate_counts show counts in a new column
#'  @param total_col add a total column (bool)
#'  @param pval_col put the p-values in a new column instead of new rows (bool)
#'  @param chi_cat_reject Automatic drop of categories with low counts
#'  @param chi_cat_reject_threshold Threshold at which categories are dropped (default:  categories with n <= 30 are dropped). Use a value >= to 1 to drop based on counts (n), or < 1 to drop based on percentages (e.g. 0.01)
#'  @param show_chi_table Returns the table used to compute the chi-squared test
#'  @param filename Name of excel file to export to
#'  @return Excel file with twoway table
#'  @export

tw_report <- function(data, workbook, worksheet, colvar, rowvars, rounding = 2,
    percentage_style = "row", counts = FALSE,
    separate_counts = FALSE, total_col = TRUE, pval_col = FALSE,
    chi_cat_reject = FALSE, chi_cat_reject_threshold = 30, show_chi_table = FALSE, filename) {

    # SETUP & STYLES
    wb <- createWorkbook(workbook)
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = "Top"
    )
    bs1 <- createStyle(
        border = "Bottom"
    )
    indent_style <- createStyle(
        indent = 1
    )

    # Counters
    row_counter <- 1

    # Missing values to explicit category
    data <- data %>%
        mutate(
            across(
                c(
                    all_of(colvar),
                    all_of(rowvars)
                ),
                ~fct_explicit_na(., na_level = "Missing values")
            )
        )


    # Column variable categories
    colvar_categories <- data %>%
        pull(colvar) %>%
        levels()

    # Subfunction
    # separate_results <- function(data, column) {
    #     column_n <- paste0(column, "_N")
    #     column_perc <- paste0(column, "_N")

    #     data %>%
    #         mutate(!!sym(column) := str_remove(!!sym(column), "\\)")) %>%
    #         separate(
    #             !!sym(column),
    #             c(paste0(column, "_N"), paste0(column, "_%")),
    #             sep = "\\("
    #         ) %>%
    #         mutate(!!sym(column_n) := str_trim(!!sym(column_n)))
    # }

    ###########################################################################

    
    ###########################################################################
    # HEADER

    # Build header
    table_header <- data %>%
        tabyl(!!sym(rowvars[1]), !!sym(colvar)) %>%
        as_tibble() %>%
        slice(0)
    # Total column
    if (total_col) {
        table_header <- table_header %>%
        mutate(Total = NA)
    }
    ## % (N) in header
    if (counts & separate_counts) {
        header_legend <- rep(c("%", "N"), times = length(table_header)-1) %>%
            t()
    } else {
        header_legend <- rep(c("% (N)"), times = length(table_header)-1) %>%
            t()
    }
    
    # P-value in a column
    if (pval_col) {
        table_header <- table_header %>%
        mutate(`P-value` = NA)
    }
    ## COLUMN VARIABLE TREATMENT
    colvar_count <- data %>%
        tabyl(!!sym(colvar))
    ### FOR CHI-SQUARED TESTS
    ### Used to filter some columns in the chi-squared tests
    if (chi_cat_reject) {
        colvar_cats_chi_reject <- colvar_count %>%
        {if (chi_cat_reject_threshold >= 1)
            filter(., n < chi_cat_reject_threshold) else
            filter(., percent < chi_cat_reject_threshold)} %>%
        pull(colvar) %>%
        as.vector()
        message("Rejected column categories for chi-sq test: ", colvar_cats_chi_reject)
    }
    ### Get counts (N + %) of column variable
    colvar_count <- colvar_count %>%
        {if (total_col) adorn_totals(.) else .} %>%
        adorn_pct_formatting(2) %>%
        mutate(n = paste0("(", as.character(n), ")")) %>%
        unite("percent", percent:n, sep = (" ")) %>%
        mutate(percent = str_remove(percent, "%")) %>%
        pull(percent) %>%
        as_tibble() %>%
        rename(Total = value) %>%
        t()

    if (counts & separate_counts) {
        colvar_count <- colvar_count %>% as_tibble()
        for (var in colnames(colvar_count)) {
            colvar_count <- colvar_count %>% separate_results(var)
        }
        colvar_count <- colvar_count %>%
            mutate(
                rowname = "Total"
            ) %>%
            column_to_rownames("rowname")
    }

    ## Write colvar name
    writeData(wb = wb, sheet = worksheet, x = colvar, startCol = 2)
    addStyle(wb, worksheet, cols = 1:(length(colvar_count)+1),
        rows = 1, style = hs1
    )
    mergeCells(wb, worksheet, cols = (1:length(colvar_count)) + 1, rows = 1)
    row_counter <- row_counter + 1
    ## Write colvar categories
    writeData(wb = wb, sheet = worksheet, x = table_header,
        startRow = row_counter
    )
    ## Remove variable name in header
    writeData(wb = wb, sheet = worksheet, x = " ", startRow = row_counter)
    
    ### Alternative header if N and % are separated
    if (counts & separate_counts) {
        alt_header_column = 2
        for (cat in colvar_categories) {
            writeData(wb = wb, sheet = worksheet, x = cat, 
                startRow = row_counter, startCol = alt_header_column)
            mergeCells(wb, worksheet, cols = (alt_header_column:(alt_header_column+1)),
                rows = row_counter)
            alt_header_column <- alt_header_column + 2
        }
        if (total_col) {
            writeData(wb = wb, sheet = worksheet, x = "Total", 
                startRow = row_counter, startCol = alt_header_column)
            mergeCells(wb, worksheet, cols = (alt_header_column:(alt_header_column+1)),
                rows = row_counter)
            alt_header_column <- alt_header_column + 2
        }
    }

    row_counter <- row_counter + 1
    ## Add header legend: % (N)
    writeData(wb = wb, sheet = worksheet, x = header_legend, startCol = 2,
        startRow = row_counter, colNames = FALSE)
    addStyle(wb, worksheet, cols = 1:(length(colvar_count)+1), style = bs1,
        rows = row_counter)
    row_counter <- row_counter + 1
    ## Add counts of col var
    writeData(wb = wb, sheet = worksheet, x = colvar_count, startCol = 1,
        startRow = row_counter, colNames = FALSE, rowNames = TRUE)
    row_counter <- row_counter + 1

    ###########################################################################
    # TABLE CONTENT
    for (rowvar in rowvars) {
    message("Cross-tabulating ", colvar, " and ", rowvar)
        # Create table
        table <- data %>%
            tabyl(!!sym(rowvar), !!sym(colvar)) %>%
            adorn_percentages(percentage_style) %>%
            adorn_pct_formatting(digits = rounding) %>%
            {if (counts) adorn_ns(.) else .} %>%
            mutate(across(all_of(colvar_categories), ~ str_remove(.x, "%"))) %>%
            as_tibble()

        # Total column
        if (total_col) {
            table_col_total <- data %>%
                tabyl(!!sym(rowvar)) %>%
                adorn_pct_formatting(2) %>%
                mutate(n = paste0("(", as.character(n), ")")) %>%
                unite("percent", percent:n, sep = (" ")) %>%
                mutate(percent = str_remove(percent, "%")) %>%
                rename(Total = percent)

            table <- table %>%
                left_join(table_col_total)

        }
        
        # Separate counts & percentages
        if (counts & separate_counts) {
            for (var in colvar_categories) {
                table <- table %>% separate_results(var)
            }
            if (total_col) {
                table <- table %>% separate_results("Total")
            }
        }

        # Delete small categories for row var
        rowvar_count <- data %>%
                tabyl(!!sym(rowvar))
        ### FOR CHI-SQUARED TESTS
        ### Used to filter some columns in the chi-squared tests
        if (chi_cat_reject) {
            rowvar_cats_chi_reject <- rowvar_count %>%
            {if (chi_cat_reject_threshold >= 1)
                filter(., n < chi_cat_reject_threshold) else
                filter(., percent < chi_cat_reject_threshold)} %>%
            pull(rowvar) %>%
            as.vector()
            message("Rejected row categories for chi-sq test: ", rowvar_cats_chi_reject)
        }

        # Calculate chisq test
        table_chi <- data %>%
            mutate(
                # Tabyl shows unused categories
                # These should not be used in the chisq test
                !!sym(rowvar) := fct_drop(!!sym(rowvar))
            ) %>%
            tabyl(!!sym(rowvar), !!sym(colvar)) %>%
            {
                if (chi_cat_reject)
                    select(., -all_of(colvar_cats_chi_reject)) %>%
                    filter(!(!!sym(rowvar) %in% rowvar_cats_chi_reject))
                else
                    .
            } %>%
            select(-!!sym(rowvar)) %>%
            as_tibble()

        # DIAGNOSTIC
        if (show_chi_table) {
            message("Observed counts used for chi-squared test :")
            print(table_chi)
        }

        chi_pval <- if (chisq.test(table_chi)$p.value < 0.001) {
            "< 0.001"
        } else if (chisq.test(table_chi)$p.value < 0.05) {
            chisq.test(table_chi)$p.value %>% round(3)
        } else {
            chisq.test(table_chi)$p.value %>% round(2)
        }

        # Table
        ## Write variable name
        writeData(wb = wb, sheet = worksheet, x = rowvar,
            startRow = row_counter
        )
        row_counter <- row_counter + 1
        ## Write content
        writeData(wb, worksheet, table, startRow = row_counter,
            colNames = FALSE
        )
        ## Indent rowvar categories
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter:(row_counter + nrow(table) - 1),
            style = indent_style
        )
        row_counter <- row_counter + nrow(table)
        ## Add p-value
        writeData(wb, worksheet, "P-value", startRow = row_counter)
        writeData(wb, worksheet, chi_pval, startRow = row_counter, startCol = 2)
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter,
            style = indent_style
        )
        row_counter <- row_counter + 1
    }

    ###########################################################################
    # FOOTER

    # Border to bottom of table
    addStyle(wb, worksheet, cols = 1:(length(colvar_count)+1),
        rows = row_counter-1, style = bs1, stack = TRUE)

    # Footer info about categories not used in chi-squared test calculations
    if (chi_cat_reject) {
        if (chi_cat_reject_threshold >= 1) {
            writeData(wb, worksheet, paste("Categories with counts <= ", chi_cat_reject_threshold, " were not used in the chi-squared test"), startRow = row_counter)
        } else {
            writeData(wb, worksheet, paste("Categories <= ", chi_cat_reject_threshold*100, "% were not used in the chi-squared test"), startRow = row_counter)
        }
    }

    ###########################################################################
    # WRITE TABLE
    openXL(wb)
    saveWorkbook(wb, filename, overwrite = TRUE)
}