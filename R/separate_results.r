separate_results <- function(data, column) {
        column_n <- paste0(column, "_N")
        column_perc <- paste0(column, "_%")

        data %>%
            mutate(!!sym(column) := str_remove(!!sym(column), "\\)")) %>%
            separate(
                !!sym(column),
                c(paste0(column, "_N"), paste0(column, "_%")),
                sep = "\\("
            ) %>%
            mutate(!!sym(column_n) := str_trim(!!sym(column_n)))
    }