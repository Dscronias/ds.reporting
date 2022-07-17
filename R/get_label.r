#' Replace a variable name (var) with a label in another dataframe
#'
#' In oneway or twoway tables, replace the name of the variable with its label
#' The label must be stored in another dataframe, with: a column with the list of variables
#' and another column with its associated label
#' @param data dataframe
#' @param var Variable to rename
#' @param data_label dataframe with the list of variables in data (in one column), and their label (in another column)
#' @param label_from column in data_label of variables from data
#' @param label_to column in data_label of labels
#' @return Same table, with a renamed variable
#' @export

put_label <- function(data, var, data_label, label_from, label_to) {
    var_label <- data_label %>% 
        filter({{label_from}} == englue("{{var}}")) %>% 
        pull({{label_to}})

    # Check that var exists in data_label$label_from
    assert_that(length(var_label) > 0, msg = glue("{englue('{{var}}')} does not exist in {englue('{{label_from}}')} (label_from) column"))

    data <- data %>%
        rename(
            !!sym(var_label) := {{var}}
        )
    return(data)
}