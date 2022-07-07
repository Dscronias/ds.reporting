#' Report of multiple quantitative variables by group to Excel
#'
#' This functions creates an Excel workbook and exports a twoway table
#' of one column variable and multiple row variables, with chi-squared tests.
#'
#' @param Data dataframe
#' @param workbook name of the workbook (string)
#' @param worksheet name of the worksheet (string)
#' @param colvar name of column variable (string, max one)
#' @param quantvars names of row variables (vector of strings)
#' @param rounding number of digits for rounding
#' @param percentage_style "row", "col" or "both" (= cell) percentages
#' @param counts show counts in parenthesis
#' @param separate_counts show counts in a new column
#' @param total_col add a total column (bool)
#' @param pval_col put the p-values in a new column instead of new rows (bool)
#' @param chi_cat_reject Automatic drop of categories with low counts
#' @param chi_cat_reject_threshold Threshold at which categories are dropped (default:  categories with n <= 30 are dropped). Use a value >= to 1 to drop based on counts (n), or < 1 to drop based on percentages (e.g. 0.01)
#' @param show_chi_table Returns the table used to compute the chi-squared test
#' @param filename Name of excel file to export to
#' @return Excel file with twoway table
#' @export

quant_group_report <- function(
    data,
    workbook,
    worksheet,
    groupvar,
    quantvars,
    rounding = 2,
    filename
) {

    # SETUP & STYLES
    wb <- createWorkbook(workbook)
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = "TopBottom"
    )
    bs1 <- createStyle(
        border = "Bottom"
    )
    indent_style <- createStyle(
        indent = 1
    )
    indent_style2 <- createStyle(
        indent = 2
    )

    # Counters
    row_counter <- 1

    # Missing values to explicit category
    data <- data %>%
        mutate(
            across(
                c(
                    all_of(groupvar)
                ),
                ~fct_explicit_na(., na_level = "Missing values")
            )
        )

    # Column variable categories
    groupvar_categories <- data %>%
        pull(groupvar) %>%
        levels()
    n_groupvar_categories <- length(groupvar_categories)

    ###########################################################################
    # Init first table

    table <- data %>%
        group_by(!!sym(groupvar)) %>%
        summarise(
            Full_N = n(),
            N = !!sym(quantvars[1]) %>% .[!is.na(!!sym(quantvars[1]))] %>% length(),
            Missing = !!sym(quantvars[1]) %>% .[is.na(!!sym(quantvars[1]))] %>% length(),
            Mean = mean(!!sym(quantvars[1]), na.rm = TRUE),
            SD = sd(!!sym(quantvars[1]), na.rm = TRUE),
            Minimum = min(!!sym(quantvars[1]), na.rm = TRUE),
            `1%` = quantile(!!sym(quantvars[1]), 0.01, na.rm = TRUE),
            `5%` = quantile(!!sym(quantvars[1]), 0.05, na.rm = TRUE),
            `25%` = quantile(!!sym(quantvars[1]), 0.25, na.rm = TRUE),
            `50%` = quantile(!!sym(quantvars[1]), 0.50, na.rm = TRUE),
            `75%` = quantile(!!sym(quantvars[1]), 0.75, na.rm = TRUE),
            `95%` = quantile(!!sym(quantvars[1]), 0.95, na.rm = TRUE),
            `99%` = quantile(!!sym(quantvars[1]), 0.99, na.rm = TRUE),
            Maximum = max(!!sym(quantvars[1]), na.rm = TRUE)
        ) %>%
        mutate(
            Missing = glue('{Missing} ({round((Missing/Full_N)*100, 2)}%)'),
            !!sym(groupvar) := glue('{Tobacco} (N = {Full_N})')
        ) %>%
        mutate(
            `95% CI` = glue(
                '[{round(Mean - qt(1-(0.05/2), N-1) * SD / sqrt(N), 2)}; {round(Mean + qt(1-(0.05/2), N-1) * SD / sqrt(N), 2)}]'
            ),
            .after = Mean
        ) %>%
        select(-Full_N)

    # HEADER
    writeData(wb = wb, sheet = worksheet, x = table %>% slice(0), startCol = 1,
        startRow = row_counter, colNames = TRUE)
    addStyle(wb, worksheet, cols = 1:(length(table)),
        rows = 1, style = hs1
    )
    ## Increment
    row_counter <- row_counter + 1

    ###########################################################################
    # First table
    ## Variable name
    writeData(wb = wb, sheet = worksheet, x = quantvars[1], startCol = 1,
        startRow = row_counter, colNames = FALSE)
    row_counter <- row_counter + 1
    ## Table
    writeData(wb = wb, sheet = worksheet, x = table, startCol = 1,
        startRow = row_counter, colNames = FALSE)
    ## Indent column categories
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter:(row_counter + nrow(table) - 1),
            style = indent_style
        )
    row_counter <- row_counter + nrow(table)

    ## T-test (if two categories in column variable)
    if (n_groupvar_categories == 2) {
        ttest_data <- data %>% 
            group_by(!!sym(groupvar)) %>%
            select(!!sym(quantvars[1])) %>%
            group_split()
        ttest_result <- t.test(
            ttest_data[[1]] %>% pull(!!sym(quantvars[1])),
            ttest_data[[2]] %>% pull(!!sym(quantvars[1])),
            paired = FALSE,
            var.equal = TRUE
        ) %>%
        broom::tidy()
        
        # Format P-val
        ttest_pval <- if (ttest_result$`p.value` < 0.001) {
            "< 0.001"
        } else if (ttest_result$`p.value` < 0.05) {
            ttest_result$`p.value` %>% round(3)
        } else {
            ttest_result$`p.value` %>% round(2)
        }

        # Write results
        writeData(wb = wb, sheet = worksheet, x = "Difference in means:", startCol = 1,
            startRow = row_counter, colNames = FALSE)
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter,
            style = indent_style2
        )
        writeData(wb = wb, sheet = worksheet, x = (ttest_result$estimate1 - ttest_result$estimate2), startCol = 4,
            startRow = row_counter, colNames = FALSE)
        writeData(wb = wb, sheet = worksheet, x = glue('[{round(ttest_result$conf.low, 2)}; {round(ttest_result$conf.high, 2)}]'), startCol = 5,
            startRow = row_counter, colNames = FALSE)
        row_counter <- row_counter + 1
        writeData(wb = wb, sheet = worksheet, x = "P-value:", startCol = 1,
            startRow = row_counter, colNames = FALSE)
        addStyle(wb, worksheet, cols = 1,
            rows = row_counter,
            style = indent_style2
        )
        writeData(wb = wb, sheet = worksheet, x = ttest_pval, startCol = 4,
            startRow = row_counter, colNames = FALSE)
        row_counter <- row_counter + 1
    }

    # Other variables in quantvars
    if (length(quantvars) > 1) {
        for (quantvar in quantvars[-1]) {
            # Create table
            table <- data %>%
                group_by(!!sym(groupvar)) %>%
                summarise(
                    Full_N = n(),
                    N = !!sym(quantvar) %>% .[!is.na(!!sym(quantvar))] %>% length(),
                    Missing = !!sym(quantvar) %>% .[is.na(!!sym(quantvar))] %>% length(),
                    Mean = mean(!!sym(quantvar), na.rm = TRUE),
                    SD = sd(!!sym(quantvar), na.rm = TRUE),
                    Minimum = min(!!sym(quantvar), na.rm = TRUE),
                    `1%` = quantile(!!sym(quantvar), 0.01, na.rm = TRUE),
                    `5%` = quantile(!!sym(quantvar), 0.05, na.rm = TRUE),
                    `25%` = quantile(!!sym(quantvar), 0.25, na.rm = TRUE),
                    `50%` = quantile(!!sym(quantvar), 0.50, na.rm = TRUE),
                    `75%` = quantile(!!sym(quantvar), 0.75, na.rm = TRUE),
                    `95%` = quantile(!!sym(quantvar), 0.95, na.rm = TRUE),
                    `99%` = quantile(!!sym(quantvar), 0.99, na.rm = TRUE),
                    Maximum = max(!!sym(quantvar), na.rm = TRUE)
                ) %>%
                mutate(
                    Missing = glue('{Missing} ({round((Missing/Full_N)*100, 2)}%)')
                ) %>%
                mutate(
                    `95% CI` = glue(
                        '[{round(Mean - qt(1-(0.05/2), N-1) * SD / sqrt(N), 2)}; {round(Mean + qt(1-(0.05/2), N-1) * SD / sqrt(N), 2)}]'
                    ),
                    .after = Mean
                ) %>%
                select(-Full_N)

            # Put variable name
            writeData(wb = wb, sheet = worksheet, x = quantvar, startCol = 1,
                startRow = row_counter, colNames = FALSE)
            row_counter <- row_counter + 1
            ## Table
            writeData(wb = wb, sheet = worksheet, x = table, startCol = 1,
                startRow = row_counter, colNames = FALSE)
            ## Indent column categories
            addStyle(wb, worksheet, cols = 1,
                rows = row_counter:(row_counter + nrow(table) - 1),
                style = indent_style
            )
            row_counter <- row_counter + nrow(table)

            ## T-test (if two categories in column variable)
            if (n_groupvar_categories == 2) {
                ttest_data <- data %>%
                    group_by(!!sym(groupvar)) %>%
                    select(!!sym(quantvar)) %>%
                    group_split()
                ttest_result <- t.test(
                    ttest_data[[1]] %>% pull(!!sym(quantvar)),
                    ttest_data[[2]] %>% pull(!!sym(quantvar)),
                    paired = FALSE,
                    var.equal = TRUE
                ) %>%
                broom::tidy()

                # Format P-val
                ttest_pval <- if (ttest_result$`p.value` < 0.001) {
                    "< 0.001"
                } else if (ttest_result$`p.value` < 0.05) {
                    ttest_result$`p.value` %>% round(3)
                } else {
                    ttest_result$`p.value` %>% round(2)
                }

                # Write t-test results
                writeData(wb = wb, sheet = worksheet, x = "Difference in means:", startCol = 1,
                    startRow = row_counter, colNames = FALSE)
                addStyle(wb, worksheet, cols = 1,
                    rows = row_counter,
                    style = indent_style2
                )
                writeData(wb = wb, sheet = worksheet, x = (ttest_result$estimate1 - ttest_result$estimate2), startCol = 4,
                    startRow = row_counter, colNames = FALSE)
                writeData(wb = wb, sheet = worksheet, x = glue('[{round(ttest_result$conf.low, 2)}; {round(ttest_result$conf.high, 2)}]'), startCol = 5,
                    startRow = row_counter, colNames = FALSE)
                row_counter <- row_counter + 1
                writeData(wb = wb, sheet = worksheet, x = "P-value:", startCol = 1,
                    startRow = row_counter, colNames = FALSE)
                addStyle(wb, worksheet, cols = 1,
                    rows = row_counter,
                    style = indent_style2
                )
                writeData(wb = wb, sheet = worksheet, x = ttest_pval, startCol = 4,
                    startRow = row_counter, colNames = FALSE)
                row_counter <- row_counter + 1
            }
        }
    }

    # Bottom border
    addStyle(wb, worksheet, cols = 1:(length(table)),
        rows = row_counter-1, style = bs1, stack = TRUE)

    ###########################################################################
    # FOOTER
    if (n_groupvar_categories == 2) {
        writeData(
            wb = wb,
            sheet = worksheet,
            x = "CI: Confidence Interval; P-values: Welch t-test of difference in means",
            startRow = row_counter
        )
    } else {
        writeData(
            wb = wb,
            sheet = worksheet,
            x = "CI: Confidence Interval",
            startRow = row_counter
        )
    }
    ###########################################################################
    # WRITE TABLE
    openXL(wb)
    saveWorkbook(wb, filename, overwrite = TRUE)
}