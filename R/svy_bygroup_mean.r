#' Oneway report of multiple categorical variables to Excel
#'
#' This functions returns a by-group summary of means for a continuous variable (with t-test of difference in means using a regression)
#' Must be used on a tbl_svy dataframe
#'
#' @param data Dataframe to use
#' @param quant_var Quantitative variable to use
#' @param by_var Variable to use for group_by
#' @param rounding_mean number of digits for rounding mean (int)
#' @param rounding_se number of digits for rounding mean standard error (int)
#' @param rounding_prct number of digits for rounding percentages (int)
#' @param rounding_ci number of digits for rounding confidence intervals (int)
#' @param show_se shows standard error of the mean
#' @param show_ci shows the confidence interval of the mean
#' @param lang changes the name of the percentage column (depending on your language). Three values : "en" (default), "fr", or "math"
#' @return Table with means, SE, CI and p-values
#' @export

svy_bygroup_mean <- function(data, quant_var, by_var, rounding_mean = 2, rounding_se = 2, rounding_prct = 2, rounding_ci = 2, show_se = FALSE, show_ci = TRUE, lang = "en") {
    reference_label <- switch(
        lang,
        "en" = "(réf.)",
        "fr" = "(ref.)",
        "math" = "(ref.)"
    )
    
    table <- data %>%
        group_by({{by_var}}) %>%
        summarise(
            `Proportion` = survey_prop(),
            Mean = survey_mean({{quant_var}}, vartype = c("se", "ci"))
        ) %>%
        mutate(
            Proportion = (Proportion * 100) %>% round(., rounding_prct) %>% format(., rounding_prct),
            `CI 95%` = glue("[{round(Mean_low, {rounding_ci}) %>% format(nsmall = {rounding_ci})}; {round(Mean_upp, {rounding_ci}) %>% format(nsmall = {rounding_ci})}]"),
            Mean = Mean %>% round(rounding_mean) %>% format(rounding_mean),
            Mean_se = Mean_se %>% round(rounding_se) %>% format(rounding_se)
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
        ) %>%
        mutate(
            {{by_var}} := case_when(
                is.na(`P-value*`) & !is.na({{by_var}}) ~ paste({{by_var}}, reference_label),
                TRUE ~ {{by_var}}
            )
        )

    return(table)
}