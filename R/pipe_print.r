#' Print an object mid-chain
#'
#' Stolen from: https://stackoverflow.com/questions/46123285/printing-intermediate-results-without-breaking-pipeline-in-tidyverse
#'
#' @param .data Input data from chaining
#' @param content Content to display using print (e.g. a message, a table...)
#' @return Input dataframe and a message
#' @export

pipe_print <- function(.data, content) {print(content); .data}
