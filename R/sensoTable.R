library(SensoMineR)
library(flextable)
library(openxlsx)

#' Title
#'
#' @param res.decat Result of a SensoMineR::decat
#' @param level.upper numeric
#' @param level.upper2 numeric
#' @param col.lower string
#' @param col.lower2 string
#' @param col.upper string
#' @param col.upper2 string
#' @importFrom FactoMineR PCA
#' @importFrom flextable flextable bg compose as_paragraph colformat_double align
#' @return sensoTable() returns a flextable
#' @export
#'
#' @examples
sensoTable = function(res.decat, level.upper=Inf,level.upper2=0.05, col.lower="#ff7979",  col.lower2="#eb4d4b", col.upper="#7ed6df", col.upper2="#22a6b3") {

  if (level.upper<=level.upper2) {
    stop("level.upper must be greater than level.upper2")
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
          color = ifelse(abs(probacrit)<level.upper2 && sign(probacrit)<0,col.lower2,
                  ifelse(abs(probacrit)<level.upper2 && sign(probacrit)>0,col.upper2,
                  ifelse(abs(probacrit)<level.upper && sign(probacrit)<0,col.lower,
                  ifelse(abs(probacrit)<level.upper && sign(probacrit)>0,col.upper,"#ffffff"
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

#' Title
#'
#' @param table flextable
#' @param filename String
#' @param path String
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @return Returns an .xlsx file based on a flextable
#' @export
#'
#' @examples
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
      cell.style = createStyle(numFmt = "0.000", border = c("top", "bottom", "left", "right"), borderColour = "black", fgFill = cell.bgcolor, halign = "center")

      # on applique le style a la cellule
      addStyle(wb,sheet=filename,style = cell.style, rows = prod+1, cols = desc)

    }
  }

  saveWorkbook(wb,paste0(filename,".xlsx"), overwrite = TRUE)

}

#
# data(chocolates)
#
# res.decat = decat(sensochoc,formul="~Product+Panelist",firstvar=5,lastvar=ncol(sensochoc), graph = TRUE)
#
# res = sensoTable(res.decat,0.05,0.01)
#
# exportxlsx(res, "testbis","C:/SSD")
