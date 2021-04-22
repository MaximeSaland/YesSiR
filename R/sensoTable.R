library(SensoMineR)
library(flextable)
library(openxlsx)

#' Color the cells of a data frame according to 2 threshold levels
#'
#' @param res.decat result of a SensoMineR::decat
#' @param thres the threshold under which cells are colored col.neg (or col.pos) if the tested coefficient is significantly lower (or higher) than the average
#' @param thres2 the threshold under which cells are colored col.neg2 (or col.pos2); this threshold should be lower than thres
#' @param col.neg the color used for thres when the tested coefficient is negative
#' @param col.neg2 the color used for thres2 when the tested coefficient is negative
#' @param col.pos the color used for thres when the tested coefficient is positive
#' @param col.pos2 the color used for thres2 when the tested coefficient is positive
#' @importFrom FactoMineR PCA
#' @importFrom flextable flextable bg compose as_paragraph colformat_double align
#' @return sensoTable() returns a flextable
#' @export
#' @description Return a colored flextable based on the result of a SensoMineR::decat according to 2 threshold levels.
#' @details This function is useful to highlight elements which are significant, especially when there are many values to check
#'
#' @examples
#' ### Example 1
#' data("sensochoc")
#' # Use the decat function
#' resdecat <-SensoMineR::decat(sensochoc, formul="~Product+Panelist", firstvar = 5, graph = FALSE)
#' sensoTable(resdecat)
#'
#' ### Example 2
#' data("sensochoc")
#' resdecat2 <-SensoMineR::decat(sensochoc, formul="~Product+Panelist", firstvar = 5, graph = FALSE)
#' sensoTable(resdecat2,thres2=0.01) # Add a second level of significance

sensoTable = function(res.decat, thres=0.05,thres2=0, col.neg="#ff7979",  col.neg2="#eb4d4b", col.pos="#7ed6df", col.pos2="#22a6b3") {

  if (thres<=thres2) {
    stop("thres must be greater than thres2")
  }
  else {
    prodadjmean = res.decat$adjmean

    # On trie selon les coordonnees des produits et des descripteurs sur la premiere dimension
    res.pca = PCA(prodadjmean, graph = FALSE)
    prodsort = names(sort(res.pca$ind$coord[,1]))
    attsort = names(sort(res.pca$var$coord[,1]))
    adjmeantable = as.data.frame(prodadjmean)
    adjmeantable = adjmeantable[prodsort,attsort]

    # creation de la flextable
    adjmeantable = flextable(cbind(row.names(adjmeantable),adjmeantable))

    for (desc in attsort) {

      for (prod in prodsort) {


        probacrit = res.decat$resT[[prod]][desc,3]*sign(res.decat$resT[[prod]][desc,4])

        if (is.na(probacrit)==TRUE) {

          color = "#ffffff"
        }
        else {
          color = ifelse(abs(probacrit)<thres2 && sign(probacrit)<0,col.neg2,
                  ifelse(abs(probacrit)<thres2 && sign(probacrit)>0,col.pos2,
                  ifelse(abs(probacrit)<thres && sign(probacrit)<0,col.neg,
                  ifelse(abs(probacrit)<thres && sign(probacrit)>0,col.pos,"#ffffff"
                  ))))
        }

        adjmeantable = bg(adjmeantable,i=prod,j=desc,bg = color)

      }

    }

    # Mise en forme de la flextable
    adjmeantable = compose(adjmeantable,i = 1, j = 1, value = as_paragraph(""), part = "header") # Cache le nom de la premiere colonne
    adjmeantable = colformat_double(adjmeantable,j=-1,digits = 3)
    adjmeantable = align(adjmeantable,align = "center")

    return(adjmeantable)
  }
}

#' Export a flextable into a .xlsx file
#'
#' @param table A flextable
#' @param filename Name of the .xlsx file
#' @param path Path to the directory stocking the .xlsx file
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @return Returns an .xlsx file based on a flextable
#' @export
#'
#' @examples
#' ft <- flextable::flextable(head(mtcars))
#' # color some cells in blue
#' ft <- flextable::bg(ft, i=ft$body$dataset$disp>200, j=3, bg = "#7ed6df", part = "body")
#' # color a few cells in yellow
#' ft <- flextable::bg(ft, i=ft$body$dataset$vs==0, j=8, bg = "#FCEC20", part = "body")
#' # export your flextable as a .xlsx in the current working directory
#' exportxlsx(ft, filename ="myFlextable", path=getwd())

exportxlsx = function(table, filename, path) {

  setwd(path) # Indique le repertoire ou sera enregistrer le fichier excel

  data = table$body$dataset
  bgcolor = as.data.frame(table$body$styles$cells$background.color$data)

  wb = createWorkbook()
  addWorksheet(wb, filename)
  writeData(wb,1,data)

  for (desc in 2:ncol(data)) {

    for (prod in 1:nrow(data)) {

      # on recupere la couleur de la cellule
      cell.bgcolor = bgcolor[prod,desc]

      # on cree un style pour la cellule
      cell.style = createStyle(numFmt = "0.000", border = c("top", "bottom", "left", "right"), borderColour = "black", fgFill = ifelse(cell.bgcolor=="transparent","#FFFFFF",cell.bgcolor), halign = "center")

      # on applique le style a la cellule
      addStyle(wb,sheet=filename,style = cell.style, rows = prod+1, cols = desc)

    }
  }

  saveWorkbook(wb,paste0(filename,".xlsx"), overwrite = TRUE)

}
