#' Oneway report of multiple categorical variables to Excel
#'
#' This functions returns a oneway table of a variable
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param quant_var Quantitative variable to use
#' @param by_var Variable to use for group_by
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param rounding_ci hides the "%" characted in the percentages column
#' @param show_se shows standard error of the mean
#' @param show_ci shows the confidence interval of the mean
#' @return Table with means, SE, CI and p-values
#' @export

svy_bygroup_mean <- function(data, quant_var, by_var, rounding_prct = 2, rounding_ci = 2, show_se = FALSE, show_ci = TRUE) {
    table <- data %>%
        group_by({{by_var}}) %>%
        summarise(
            `Proportion` = survey_prop(),
            Mean = survey_mean({{quant_var}}, vartype = c("se", "ci"))
        ) %>%
        mutate(
            Proportion = (Proportion * 100) %>% round(., rounding_prct) %>% format(., rounding_prct) %>% paste(., "%", sep = ""),
            `CI 95%` = glue("[{round(Mean_low, {rounding_ci}) %>% format(nsmall = {rounding_ci})}; {round(Mean_upp, {rounding_ci}) %>% format(nsmall = {rounding_ci})}]")
        ) %>%
        rename(
            `Standard error` = Mean_se
        ) %>%
        select(-Mean_low, -Mean_upp, -Proportion_se) %>%
        {if (show_ci == FALSE)
            select(., !`CI 95%`)
        else .} %>%
        {if (show_se == FALSE)
            select(., !`Standard error`)
        else .}

    regression <- data %>%
        survey::svyglm(as.formula(glue("{englue('{{quant_var}}')} ~ {englue('{{by_var}}')}")), design = .) %>%
        tidy() %>%
        mutate(
            term = sub(englue('{{by_var}}'), "", term),
        ) %>%
        filter(
            term != "(Intercept)"
        ) %>%
        select(term, p.value) %>%
        rename(
            {{by_var}} := term,
            `P-value*` = p.value
        )

    table <- table %>%
        left_join(
            regression,
            by = englue('{{by_var}}')
        )

    return(table)
}