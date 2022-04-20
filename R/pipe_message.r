#' Display a message mid-chain
#'
#' Stolen from: https://stackoverflow.com/questions/46123285/printing-intermediate-results-without-breaking-pipeline-in-tidyverse
#'
#' @param .data Input data from chaining
#' @param status Message to display
#' @return Input dataframe and a message
#' @export

pipe_message <- function(.data, status) {message(status); .data}
