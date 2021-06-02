#' PowerPoint reporting of a description of categories
#'
#' @param res decat result
#' @param yes_study_name title displayed on the first slide
#' @param path PowerPoint file to be created
#' @param file.name name of the PowerPoint file
#' @param x1 dimension to plot on the x-axis
#' @param x2 dimension to plot on the y-axis
#' @param size_tab maximum number of rows of a table per slide
#' @param col.neg color for negative values (on the wordclouds)
#' @param col.pos color for positive values (on the wordclouds)
#' @export
#'
#' @examples
#' \dontrun{
#' data("sensochoc")
#' res.decat <- SensoMineR::decat(sensochoc, formul="~Product+Panelist", firstvar = 5, graph = FALSE)
#' # Create the PowerPoint in the current working directory
#' Yes_decat(res.decat)
#' }

Yes_decat <- function(res, yes_study_name = "Quantitative description of products", path=getwd(), file.name = "decat_results.pptx", x1=1, x2=2, size_tab=10, col.neg="red", col.pos="blue"){

  yes_temp = system.file("YesSiR_template.pptx", package = "YesSiR")
  #######################################
  ### Init: first slide ----
  yes_slide_num=0
  set.seed(42)
  essai <- officer::read_pptx(yes_temp)
  essai <- officer::remove_slide(essai)
  essai <- officer::add_slide(essai, layout = "Diapositive de titre", master = "YesSiR")
  essai <- officer::ph_with(essai, value = yes_study_name, location = officer::ph_location_type(type = "ctrTitle"))

  #######################################

  prods <- row.names(res$adjmean)

  #######################################
  ### Sensory profile for each product ----

  if (is.null(res$resT)==FALSE) {
    # resT contains all the results from res$resT in a single table used for creating the wordclouds
    resT <- res$resT
    for (prod in names(res$resT)) {
      resT[[prod]] <- dplyr::mutate(resT[[prod]], descriptor=row.names(resT[[prod]]), .before=1)
    }
    resT <- dplyr::bind_rows(resT,.id = "product")
    for (prod in prods) {
      if (is.null(res$resT[[prod]])==FALSE) {
        # yes_resT used to print the tables
        yes_resT <- NULL
        yes_resT <- cbind(row.names(res$resT[[prod]]), res$resT[[prod]])
        colnames(yes_resT) <- c("Descriptor",colnames(yes_resT)[-1])
        # creation of the wordcloud
        yes_wordcloud <- (ggplot2::ggplot(resT[resT$product==prod,], ggplot2::aes(label = descriptor, size=Vtest, colour=Vtest)) +
                            ggwordcloud::geom_text_wordcloud(show.legend = TRUE) +
                            ggplot2::scale_size_area(max_size = 15, limits=c(min(resT$Vtest),max(resT$Vtest))) +
                            ggplot2::theme_minimal() +
                            ggplot2::scale_color_gradient2(low = col.neg, mid = "grey", high = col.pos, limits=c(min(resT$Vtest),max(resT$Vtest))) +
                            ggplot2::guides(size = FALSE))
        if (dim(yes_resT)[1]<=size_tab){
          ft <- flextable::flextable(yes_resT)
          ft <- flextable::colformat_double(ft,j=-1,digits = 3)
          ft <- flextable::align(ft,align = "center")
          ft <- flextable::autofit(ft)
          essai <- officer::add_slide(essai,layout ="Deux contenus", master = "YesSiR")
          essai <- officer::ph_with(essai, value = paste("Result of the T-test for",prod), location = officer::ph_location_type(type="title"))
          essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
          essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
          yes_slide_num <- yes_slide_num+1
          essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
        } else {
          n_slide=NULL
          n_slide <- dim(yes_resT)[1]%/%size_tab
          for (i in 0:(n_slide-1)){
            ft <- flextable::flextable(yes_resT[(i*size_tab+1):(size_tab*(i+1)),])
            ft <- flextable::colformat_double(ft,j=-1,digits = 3)
            ft <- flextable::align(ft,align = "center")
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Deux contenus", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Result of the T-test for",prod), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
            # essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
          }
          if (dim(yes_resT)[1]%%size_tab!=0){
            ft <- flextable::flextable(yes_resT[(n_slide*size_tab+1):dim(yes_resT)[1],])
            ft <- flextable::colformat_double(ft,j=-1,digits = 3)
            ft <- flextable::align(ft,align = "center")
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Deux contenus", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Result of the T-test for",prod), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
            essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
          } else {
            essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
          }
        }
      }
    }
    # All the wordclouds on the same page
    yes_wordclouds <- (ggplot2::ggplot(resT, ggplot2::aes(label = descriptor, size=Vtest, colour=Vtest)) +
                         ggwordcloud::geom_text_wordcloud(show.legend = TRUE) +
                         ggplot2::scale_size_area(max_size = 10) +
                         ggplot2::theme_minimal() +
                         ggplot2::facet_wrap(~product) +
                         ggplot2::scale_color_gradient2(low = col.neg, mid = "grey", high = col.pos) +
                         ggplot2::guides(size=FALSE))
    essai <- officer::add_slide(essai, layout = "Image avec legende", master = "YesSiR")
    essai <- officer::ph_with(essai, value = "All sensory profiles illustrated with wordclouds", location = officer::ph_location_type(type = "title"))
    essai <- officer::ph_with(essai, value = yes_wordclouds, location = officer::ph_location_type(type = "pic"))
    yes_slide_num <- yes_slide_num+1
    essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  }
  #######################################



  #######################################
  # Product space ----

  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Product space", location = officer::ph_location_type(type = "title"))

  #######################################
  # Attributes structuring the product space ----

  if (is.null(res$resF)==FALSE) {
    yes_resF <- cbind(row.names(res$resF), res$resF)
    colnames(yes_resF) = c("Descriptor", colnames(yes_resF)[-1])
    if (dim(yes_resF)[1]<size_tab){
      ft <- flextable::flextable(yes_resF)
      ft <- flextable::colformat_double(ft,j=-1,digits = 3)
      ft <- flextable::align(ft,align = "center")
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai,layout ="Titre et contenu", master = "YesSiR")
      essai <- officer::ph_with(essai, value = "Attributes structuring the product space", location = officer::ph_location_type(type="title"))
      essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type="body"))
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
    } else {
      n_slide=NULL
      n_slide <- dim(yes_resF)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(yes_resF[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft,j=-1,digits = 3)
        ft <- flextable::align(ft,align = "center")
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
        essai <- officer::ph_with(essai, value = "Attributes structuring the product space", location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
      if (dim(yes_resF)[1]%%size_tab!=0){
        ft <- flextable::flextable(yes_resF[(n_slide*size_tab+1):dim(yes_resF)[1],])
        ft <- flextable::colformat_double(ft,j=-1,digits = 3)
        ft <- flextable::align(ft,align = "center")
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
        essai <- officer::ph_with(essai, value = "Attributes structuring the product space", location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
    }
  }
  #######################################

  #######################################
  # Representation of the product space using a PCA

  res.pca <- FactoMineR::PCA(res$adjmean,graph = FALSE)

  yes_plotprod1 <- FactoMineR::plot.PCA(res.pca, axes = c(x1,x2), choix = "ind", title = "Representation of the products from PCA")
  essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the products from PCA", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plotprod1, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))

  yes_plotdesc1 <- FactoMineR::plot.PCA(res.pca, axes = c(x1,x2), choix = "var", title = "Representation of the descriptors from PCA")
  essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the descriptors from PCA", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plotdesc1, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))

  yes_adjmean <- YesSiR::sensoTable(res, thres2 = 0.01)
  yes_adjmean <- flextable::autofit(yes_adjmean)
  essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the sensory profiles", location = officer::ph_location_type(type="title"))
  essai <- officer::ph_with(essai, value = yes_adjmean, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################

  # Appendix
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Appendix", location = officer::ph_location_type(type = "title"))

  #######################################
  # spider plots ----

  spiderplot = function(data, product=row.names(data$adjmean),branches=row.names(data$resF)[1:ifelse(nrow(data$resF)>6,6,nrow(data$resF))], prod.col = rep("lightblue",length(product))) {
    globalmean <- colMeans(data$adjmean[,branches])
    for(i in 1:length(product)){
      plotrix::radial.plot(globalmean,labels = branches, rp.type = "p", line.col = "grey", poly.col=grDevices::adjustcolor("grey",0.5), start = pi/2, clockwise = TRUE, main = product[i], show.grid.labels = T, radial.lim = c(0,10))
      plotrix::radial.plot(data$adjmean[product[i],branches], rp.type="p", line.col = prod.col[i], poly.col = grDevices::adjustcolor(prod.col[i], 0.5), start = pi/2, clockwise = TRUE, radial.lim = c(0,10), add=T)
      plotrix::radial.plot(data$adjmean[product[i],branches], rp.type="s", point.symbols = 20, point.col = ifelse(is.na(data$resT[[product[i]]][branches,1])==TRUE, grDevices::adjustcolor("grey",0),"black"), start = pi/2, clockwise = TRUE, radial.lim = c(0,10), add=T)
      graphics::legend(-17,-9,c("Mean for all products","Adjusted mean", "Significance"), pch = 15, col=c("grey",prod.col[i], "black"), bty = "n", cex=1.3)
    }
  }

  for (prod in prods) {

    grDevices::png("spiderplot.png", height = 720, width = 1080)
    spiderplot(res,prod)
    grDevices::dev.off()
    yes_spider <- officer::external_img("spiderplot.png", height = 5.1, width = 7.7)
    essai <- officer::add_slide(essai, layout = "Titre et image", master = "YesSiR")
    essai <- officer::ph_with(essai, value = paste("Sensory profile of",prod), location = officer::ph_location_type(type = "title"))
    essai <- officer::ph_with(essai, value = yes_spider, location = officer::ph_location_type(type = "body"), use_loc_size=FALSE)
    yes_slide_num <- yes_slide_num+1
    essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
    unlink("spiderplot.png")
  }
  #######################################
  print(essai, target = paste(path,file.name, sep="/"))
}
