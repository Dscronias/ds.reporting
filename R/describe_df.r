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
        class = df %>% sapply(class),
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
        nlevels = df %>% summarise_if(
            is.factor,
            ~levels(.) %>% length()
        ) %>%
        t() %>%
        as.vector(),
        levels = df %>% summarise_if(
            #? Does not work properly using across
            is.factor,
            ~levels(.) %>% paste(., collapse = "; ")
        ) %>%
        t() %>%
        as.vector()
    )
}
