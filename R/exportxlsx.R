#' Export a flextable into a .xlsx file
#'
#' @param table A flextable
#' @param path Path to the Excel file to be created
#' @return Returns an .xlsx file based on a flextable
#' @export
#'
#' @examples
#' \dontrun{
#' ft <- flextable::flextable(head(mtcars))
#' # color some cells in blue
#' ft <- flextable::bg(ft, i=ft$body$dataset$disp>200, j=3, bg = "#7ed6df", part = "body")
#' # color a few cells in yellow
#' ft <- flextable::bg(ft, i=ft$body$dataset$vs==0, j=8, bg = "#FCEC20", part = "body")
#' # export your flextable as a .xlsx in the current working directory
#' exportxlsx(ft, filename ="myFlextable", path="path/to/the/excel_file.xlsx")
#' }

exportxlsx = function(table, path) {

  # setwd(path) # Indique le repertoire ou sera enregistrer le fichier excel

  data = table$body$dataset
  bgcolor = as.data.frame(table$body$styles$cells$background.color$data)

  wb = openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "feuille1")
  openxlsx::writeData(wb,1,data)

  for (desc in 2:ncol(data)) {

    for (prod in 1:nrow(data)) {

      # on recupere la couleur de la cellule
      cell.bgcolor = bgcolor[prod,desc]

      # on cree un style pour la cellule
      cell.style = openxlsx::createStyle(numFmt = "0.000", border = c("top", "bottom", "left", "right"), borderColour = "black", fgFill = ifelse(cell.bgcolor=="transparent","#FFFFFF",cell.bgcolor), halign = "center")

      # on applique le style a la cellule
      openxlsx::addStyle(wb,sheet=1,style = cell.style, rows = prod+1, cols = desc)

    }
  }

  openxlsx::saveWorkbook(wb,path, overwrite = TRUE)

}
