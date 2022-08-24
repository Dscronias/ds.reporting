#' Quick description of dataframe
#'
#' This functions returns a dataframe with some information:
#' name of variables, labels, number of factor levels...
#'
#' @param data dataframe to describe
#' @return A new dataframe with informations about the input dataframe
#' @export

describe_df <- function(data) {
    bind_cols(
        variable = data %>% colnames(),
        class = data %>% map_chr(~ class(.) %>% paste(collapse = "; ")),
        count = data %>%
            summarise(
                across(
                    everything(),
                    ~sum(!is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        missing = data %>%
            summarise(
                across(
                    everything(),
                    ~sum(is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        nlevels = data %>% summarise(
            across(
                everything(),
                ~levels(.) %>% length()
            )
        ) %>%
        t() %>%
        as.vector(),
        levels = data %>% summarise(
            across(
                everything(),
                ~levels(.) %>% paste(collapse = "; ")
            )
        ) %>%
        t() %>%
        as.vector()
    ) %>%
    mutate(
        count = glue("{count} ({round(count/nrow(df)*100, 2)}%)"),
        missing = glue("{missing} ({round(missing/nrow(df)*100, 2)}%)")
    )
}
