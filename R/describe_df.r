#' Quick description of dataframe
#'
#' This functions returns a dataframe with some information:
#' name of variables, labels, number of factor levels...
#'
#' @param data dataframe to describe
#' @return A new dataframe with informations about the input dataframe
#' @export

describe_df <- function(data) {
    dplyr::bind_cols(
        variable = data %>% colnames(),
        class = data %>% purrr::map_chr(~ class(.) %>% paste(collapse = "; ")),
        count = data %>%
            dplyr::summarise(
                dplyr::across(
                    dplyr::everything(),
                    ~sum(!is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        missing = data %>%
            dplyr::summarise(
                dplyr::across(
                    dplyr::everything(),
                    ~sum(is.na(.))
                )
            ) %>%
            t() %>%
            as.vector(),
        nlevels = data %>% dplyr::summarise(
            dplyr::across(
                dplyr::everything(),
                ~levels(.) %>% length()
            )
        ) %>%
        t() %>%
        as.vector(),
        levels = data %>% dplyr::summarise(
            dplyr::across(
                dplyr::everything(),
                ~levels(.) %>% paste(collapse = "; ")
            )
        ) %>%
        t() %>%
        as.vector()
    ) %>%
    dplyr::mutate(
        count = glue::glue("{count} ({round(count/nrow(df)*100, 2)}%)"),
        missing = glue::glue("{missing} ({round(missing/nrow(df)*100, 2)}%)")
    )
}
