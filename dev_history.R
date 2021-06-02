available::available("yesSIR")
usethis::create_package("C:/SSD/Stage AO/yesSIR")
usethis::use_build_ignore("dev_history.R")

usethis::use_package("flextable")
usethis::use_package("SensoMineR")
usethis::use_package("FactoMineR")
usethis::use_package("openxlsx")
usethis::use_package("officer")
usethis::use_tidy_description()

attachment::att_to_description()
usethis::use_r("yes_MCA")

library(SensoMineR)
data(chocolates)
usethis::use_data(sensochoc)

usethis::use_r("exportxlsx")

usethis::use_git_ignore(".Rhistory")

data(tea)
usethis::use_data(tea)

usethis::use_r("Yes_decat")

usethis::use_tidy_description()
attachment::att_amend_desc()
attachment::att_amend_desc(extra.suggests = "SensoMineR")

usethis::use_r("Yes_textual")
usethis::use_r("Yes_PCA")

attachment::att_from_rscripts()

usethis::use_vignette("powerpoint-reporting")
