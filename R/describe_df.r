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
        variable = df %>% colnames(),
        class = df %>% map_chr(~ class(.) %>% paste(collapse = "; ")),
        count = df %>%
            summarise(
                across(
                    everything(),
                    ~sum(!is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        missing = df %>%
            summarise(
                across(
                    everything(),
                    ~sum(is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        nlevels = df %>% summarise(
            across(
                everything(),
                ~levels(.) %>% length()
            )
        ) %>%
        t() %>%
        as.vector(),
        levels = df %>% summarise(
            across(
                everything(),
                ~levels(.) %>% paste(collapse = "; ")
            )
        ) %>%
        t() %>%
        as.vector()
    )
}
