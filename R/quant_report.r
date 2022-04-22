#' Descriptive statistics for continuous variables
#'
#' This functions returns a dataframe with some descriptive statistics from continuous vars
#' name of variables, labels, number of factor levels...
#'
#' @param data Dataframe to use
#' @param columns String of list of strings with columns to describe
#' @return A dataframe with descriptive statistics
#' @export

quant_report <- function(data, workbook, worksheet, vars, filename,
    open_table_on_completion = FALSE) {

    # HEADER STYLE
    wb <- createWorkbook(workbook)
    addWorksheet(wb, worksheet)

    hs1 <- createStyle(
        border = c("Top", "Bottom")
    )
    bs1 <- createStyle(
        border = "Bottom"
    )

    # TABLE
    report <- data %>% quant_summary(vars)

    # TO EXCEL
    ## TABLE
    writeData(
        wb = wb,
        sheet = worksheet,
        x = report
    )
    ## STYLE
    ### HEADER
    addStyle(
        wb,
        worksheet,
        cols = 1:(length(report)),
        rows = 1,
        style = hs1
    )
    ### FOOTER
    addStyle(
        wb,
        worksheet,
        cols = 1:(length(report)),
        rows = nrow(report)+1,
        style = bs1
    )

    # WRITE TABLE
    if (open_table_on_completion) {openXL(wb)}
    saveWorkbook(wb, filename, overwrite = TRUE)
}
