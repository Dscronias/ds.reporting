#' Descriptive statistics for continuous variables
#'
#' This functions returns a dataframe with some descriptive statistics from continuous vars
#' name of variables, labels, number of factor levels...
#'
#' @param data Dataframe to use
#' @param columns String of list of strings with columns to describe
#' @return A dataframe with descriptive statistics
#' @export

quant_summary <- function(data, columns) {
    data %>%
        summarise(
            across(
                all_of(columns),
                list(
                    N = ~ na.omit(.x) %>% length(),
                    Mean = ~ mean(.x, na.rm = TRUE),
                    SD = ~ sd(.x, na.rm = TRUE),
                    Minimum = ~ min(.x, na.rm = TRUE),
                    `1%` = ~ quantile(.x, 0.01, na.rm = TRUE),
                    `5%` = ~ quantile(.x, 0.05, na.rm = TRUE),
                    `25%` = ~ quantile(.x, 0.25, na.rm = TRUE),
                    `50%` = ~ quantile(.x, 0.50, na.rm = TRUE),
                    `75%` = ~ quantile(.x, 0.75, na.rm = TRUE),
                    `95%` = ~ quantile(.x, 0.95, na.rm = TRUE),
                    `99%` = ~ quantile(.x, 0.99, na.rm = TRUE),
                    Maximum = ~ max(.x, na.rm = TRUE)
                )
            )
        ) %>%
        pivot_longer(
            cols = everything(),
            names_sep = "_(?!.*_)",
            names_to = c("Variable", ".value")
        ) %>%
        mutate(
            Missing = glue('{round((nrow(df)-N)/nrow(df)*100, 2)}%'),
            .after = N
        )
}
