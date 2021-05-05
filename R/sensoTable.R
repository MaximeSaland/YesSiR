#' Color the cells of a data frame according to 2 threshold levels
#'
#' @param res.decat result of a SensoMineR::decat
#' @param thres the threshold under which cells are colored col.neg (or col.pos) if the tested coefficient is significantly lower (or higher) than the average
#' @param thres2 the threshold under which cells are colored col.neg2 (or col.pos2); this threshold should be lower than thres
#' @param col.neg the color used for thres when the tested coefficient is negative
#' @param col.neg2 the color used for thres2 when the tested coefficient is negative
#' @param col.pos the color used for thres when the tested coefficient is positive
#' @param col.pos2 the color used for thres2 when the tested coefficient is positive
#' @return Returns a formatted flextable
#' @export
#' @description Colors a flextable based on the result of a SensoMineR::decat according to 2 threshold levels.
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
    res.pca = FactoMineR::PCA(prodadjmean, graph = FALSE)
    prodsort = names(sort(res.pca$ind$coord[,1]))
    attsort = names(sort(res.pca$var$coord[,1]))
    adjmeantable = as.data.frame(prodadjmean)
    adjmeantable = adjmeantable[prodsort,attsort]

    # creation de la flextable
    adjmeantable = flextable::flextable(cbind(row.names(adjmeantable),adjmeantable))

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

        adjmeantable = flextable::bg(adjmeantable,i=prod,j=desc,bg = color)

      }

    }

    # Mise en forme de la flextable
    adjmeantable = flextable::compose(adjmeantable,i = 1, j = 1, value =  flextable::as_paragraph(""), part = "header") # Cache le nom de la premiere colonne
    adjmeantable = flextable::colformat_double(adjmeantable,j=-1,digits = 3)
    adjmeantable = flextable::align(adjmeantable,align = "center")

    return(adjmeantable)
  }
}
