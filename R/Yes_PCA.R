#' PowerPont reporting of a PCA
#'
#' @param res a PCA object
#' @param yes_study_name title displayed on the first slide
#' @param path PowerPoint file to be created
#' @param file.name name of the PowerPoint file
#' @param x1 component to plot on the x-axis
#' @param x2 component to plot on the y-axis
#' @param proba the significance threshold considered to characterized the category (by default 0.05)
#' @param size_tab maximum number of rows of a table per slide
#'
#' @return Returns a .pptx file
#' @export
#'
#' @examples
#' \dontrun{
#' library(FactoMineR)
#' data(decathlon)
#' res.pca <- FactoMineR::PCA(decathlon, quanti.sup = 11:12, quali.sup=13)
#' # Create the PowerPoint in the current working directory
#' Yes_PCA(res.pca)
#' }
Yes_PCA <- function(res,yes_study_name = "PCA + HCPC",path=getwd(), file.name="PCA_results.pptx",x1=1,x2=2,proba=0.05,size_tab=10){

  location_body <- officer::ph_location_type(type = "body")
  location_title <- officer::ph_location_type(type = "title")
  location_sldNum <- officer::ph_location_type(type = "sldNum")

  yes_temp = system.file("YesSiR_template.pptx", package = "YesSiR")
  #######################################
  #Init: first slide
  yes_slide_num=0
  essai <- officer::read_pptx(yes_temp)
  essai <- officer::remove_slide(essai)
  essai <- officer::add_slide(essai, layout = "Diapositive de titre", master = "YesSiR")
  essai <- officer::ph_with(essai, value = yes_study_name, location = officer::ph_location_type(type = "ctrTitle"))
  #######################################

  #######################################
  # Eigenvalues
  lim_eig <- min(size_tab,dim(res$eig)[1])
  yes_eig <- as.data.frame(res$eig[1:lim_eig,])
  yes_eig <- round(yes_eig,4)
  yes_eig <- cbind(rownames(yes_eig),yes_eig)
  colnames(yes_eig) <- c("Dimension","Eigenvalue","Percentage of variance","Cumulative percentage of variance")
  ft <- flextable::flextable(yes_eig)
  ft <- flextable::colformat_double(ft, j=-1, digits = 3)
  ft <- flextable::autofit(ft)

  #Eigenvalues slide
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Eigenvalues Decomposition", location = location_title)
  essai <- officer::ph_with(essai, value = ft, location = location_body)
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  #######################################

  #######################################
  # Graphical representation
  yes_plot.ind1 <- FactoMineR::plot.PCA(res,choix="ind",invisible=c("quali"),label = "none",title = "Representation of the individuals",axes=c(x1,x2))

  #Slide: individuals graph
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the individuals", location = location_title)
  essai <- officer::ph_with(essai, value = yes_plot.ind1, location = location_body)
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  #######################################

  #######################################
  # Graphical representation
  yes_plot.ind2 <- FactoMineR::plot.PCA(res,choix="ind",label = "quali",title = "Representation of the individuals and the categories",axes=c(x1,x2))

  #Slide: individuals + categories graph
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the individuals and the categories", location = location_title)
  essai <- officer::ph_with(essai, value = yes_plot.ind2, location = location_body)
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  #######################################

  #######################################
  # Graphical representation
  yes_plot.var1 <- FactoMineR::plot.PCA(res,choix="var",title = "Representation of the variables",axes=c(x1,x2))

  #Slide: variable graph
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the variables", location = location_title)
  essai <- officer::ph_with(essai, value = yes_plot.var1, location = location_body)
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  #######################################
  # End of first section ----


  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1), location = location_title)

  yes_dimdesc <- FactoMineR::dimdesc(res,proba = proba, axes = c(x1,x2))

  #######################################
  #Axis x1 Description by Active Variables

  yes_dimdesc_x1 <- yes_dimdesc[[paste0("Dim.",x1)]]
  sel <- row.names(yes_dimdesc_x1$quanti)

  if (is.null(sel)==FALSE){

    sel_act_quanti <- as.data.frame(cbind(row.names(yes_dimdesc_x1$quanti),yes_dimdesc_x1$quanti))
    colnames(sel_act_quanti)[1] <- "Variable"
    sel_act_quanti[,2] <- as.numeric(sel_act_quanti[,2])
    sel_act_quanti[,3] <- as.numeric(sel_act_quanti[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_quanti)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Quantitative Variables"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)

    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_quanti)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_quanti[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Quantitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_quanti)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_quanti[(n_slide*size_tab+1):dim(sel_act_quanti)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Quantitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }

  #######################################
  #Axis x1 Description by qualitative Variables

  sel <- row.names(yes_dimdesc_x1$quali)

  if (is.null(sel)==FALSE){

    sel_act_quali <- as.data.frame(cbind(row.names(yes_dimdesc_x1$quali),yes_dimdesc_x1$quali))
    colnames(sel_act_quali)[1] <- "Variable"
    sel_act_quali[,2] <- as.numeric(sel_act_quali[,2])
    sel_act_quali[,3] <- as.numeric(sel_act_quali[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_quali)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Qualitative Variables"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)

    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_quali)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_quali[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Qualitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_quali)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_quali[(n_slide*size_tab+1):dim(sel_act_quali)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Qualitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }
  #######################################

  #######################################
  #Axis x1 Description by the categories

  sel <- row.names(yes_dimdesc_x1$category)

  if (is.null(sel)==FALSE){

    sel_act_categ <- as.data.frame(cbind(row.names(yes_dimdesc_x1$category),yes_dimdesc_x1$category))
    colnames(sel_act_categ)[1] <- "Category"
    sel_act_categ[,2] <- as.numeric(sel_act_categ[,2])
    sel_act_categ[,3] <- as.numeric(sel_act_categ[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_categ)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with the Categories"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)

    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_categ)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_categ[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with the Categories"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_categ)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_categ[(n_slide*size_tab+1):dim(sel_act_categ)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with the Categories"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }
  #######################################
  #End of second section ----

  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2), location = location_title)

  #######################################
  #Axis x2 Description by Active Variables
  yes_dimdesc_x2 <- yes_dimdesc[[paste0("Dim.",x2)]]
  sel <- row.names(yes_dimdesc_x2$quanti)

  if (is.null(sel)==FALSE){

    sel_act_quanti <- as.data.frame(cbind(row.names(yes_dimdesc_x2$quanti),yes_dimdesc_x2$quanti))
    colnames(sel_act_quanti)[1] <- "Variable"
    sel_act_quanti[,2] <- as.numeric(sel_act_quanti[,2])
    sel_act_quanti[,3] <- as.numeric(sel_act_quanti[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_quanti)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Quantitative Variables"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_quanti)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_quanti[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Quantitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_quanti)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_quanti[(n_slide*size_tab+1):dim(sel_act_quanti)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Quantitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }
  #######################################

  #######################################
  #Axis x2 Description by qualitative Variables

  sel <- row.names(yes_dimdesc_x2$quali)

  if (is.null(sel)==FALSE){

    sel_act_quali <- as.data.frame(cbind(row.names(yes_dimdesc_x2$quali),yes_dimdesc_x2$quali))
    colnames(sel_act_quali)[1] <- "Variable"
    sel_act_quali[,2] <- as.numeric(sel_act_quali[,2])
    sel_act_quali[,3] <- as.numeric(sel_act_quali[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_quali)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Qualitative Variables"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)

    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_quali)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_quali[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Qualitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_quali)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_quali[(n_slide*size_tab+1):dim(sel_act_quali)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Qualitative Variables"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }
  #######################################

  #######################################
  #Axis x2 Description by the categories

  sel <- row.names(yes_dimdesc_x2$category)

  if (is.null(sel)==FALSE){

    sel_act_categ <- as.data.frame(cbind(row.names(yes_dimdesc_x2$category),yes_dimdesc_x2$category))
    colnames(sel_act_categ)[1] <- "Category"
    sel_act_categ[,2] <- as.numeric(sel_act_categ[,2])
    sel_act_categ[,3] <- as.numeric(sel_act_categ[,3])

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_act_categ)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with the Categories"), location = location_title)
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)

    } else {
      n_slide=NULL
      n_slide <- dim(sel_act_categ)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_act_categ[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with the Categories"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
      if (dim(sel_act_categ)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_act_categ[(n_slide*size_tab+1):dim(sel_act_categ)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with the Categories"), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = location_body)
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      }
    }
  }
  #######################################
  # #End of third section ----

  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Hierarchical Clutering", location = location_title)

  ###################################################
  #    Classification                               #
  ###################################################

  res.hcpc <- FactoMineR::HCPC(res,nb.clust = -1, graph=FALSE)
  sel_act_desc <- rownames(res$var$contrib)
  data_act_class <- res.hcpc$data.clust[,c(sel_act_desc,"clust")]
  res_inter <- FactoMineR::PCA(data_act_class,quali.sup = dim(data_act_class)[2],graph = FALSE)
  yes_res.plot.hcpc <- FactoMineR::plot.PCA(res_inter,habillage = dim(data_act_class)[2],label = "none",invisible = "quali",
                                            title = "Representation of individuals according to hierarchical clustering")

  ##########################
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of individuals according to hierarchical clustering", location = location_title)
  essai <- officer::ph_with(essai, value = yes_res.plot.hcpc, location = location_body)
  ##########################

  rescat <- FactoMineR::catdes(data_act_class,num.var=dim(data_act_class)[2],proba=proba)
  #Test
  if (is.null(rescat$quanti)==FALSE){
    nlevels(data_act_class$clust)
    for (k in 1:nlevels(data_act_class$clust)){
      if (is.null(rescat$quanti[[k]])==FALSE){
        rescat$quanti[k]
        desc_classe <- as.data.frame(rescat$quanti[k])
        desc_classe <- cbind(rownames(desc_classe),desc_classe)
        colnames(desc_classe) <- c("Variable","v.test","Mean in category","Overall mean","sd in category","Overall sd","p.value")
        if (dim(desc_classe)[1]<size_tab){
          ft <- flextable::flextable(desc_classe)
          ft <- flextable::colformat_double(ft, j=-1, digits = 3)
          ft <- flextable::autofit(ft)
          essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
          essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = location_title)
          essai <- officer::ph_with(essai, value = ft, location = location_body)
          yes_slide_num <- yes_slide_num+1
          essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
        } else {
          n_slide=NULL
          n_slide <- dim(desc_classe)[1]%/%size_tab
          for (i in 0:(n_slide-1)){
            ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = location_title)
            essai <- officer::ph_with(essai, value = ft, location = location_body)
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
          }
          if (dim(desc_classe)[1]%%size_tab!=0){
            ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = location_title)
            essai <- officer::ph_with(essai, value = ft, location = location_body)
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
          }
        }
      }
    }
  }

  sel_quali_desc <- rownames(res$quali.sup$eta2)
  if (is.null(sel_quali_desc)==FALSE){
    data_ill_class <- res.hcpc$data.clust[,c(sel_quali_desc,"clust")]
    rescat <- FactoMineR::catdes(data_ill_class,num.var=dim(data_ill_class)[2],proba=proba)
    #Test
    if (is.null(rescat$category)==FALSE){
      nlevels(data_ill_class$clust)
      for (k in 1:nlevels(data_ill_class$clust)){
        if (is.null(rescat$category[[k]])==FALSE){
          rescat$category[k]
          desc_classe <- as.data.frame(rescat$category[k])
          desc_classe <- cbind(rownames(desc_classe),desc_classe)
          colnames(desc_classe) <- c("Category","Cla/Mod","Mod/Cla","Global","p.value","v.test")
          if (dim(desc_classe)[1]<size_tab){
            ft <- flextable::flextable(desc_classe)
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = location_title)
            essai <- officer::ph_with(essai, value = ft, location = location_body)
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
          } else {
            n_slide=NULL
            n_slide <- dim(desc_classe)[1]%/%size_tab
            for (i in 0:(n_slide-1)){
              ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = location_title)
              essai <- officer::ph_with(essai, value = ft, location = location_body)
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
            }
            if (dim(desc_classe)[1]%%size_tab!=0){
              ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = location_title)
              essai <- officer::ph_with(essai, value = ft, location = location_body)
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
            }
          }
        }
      }
    }
  }

  #################################################
  data_ill_class=NULL
  sel_quanti_desc <- rownames(res$quanti.sup$coord)

  if (is.null(sel_quanti_desc)==FALSE){
    data_ill_class <- res.hcpc$data.clust[,c(sel_quanti_desc,"clust")]
    rescat <- FactoMineR::catdes(data_ill_class,num.var=dim(data_ill_class)[2],proba=proba)

    if (is.null(rescat$quanti)==FALSE){
      nlevels(data_ill_class$clust)
      for (k in 1:nlevels(data_ill_class$clust)){
        if (is.null(rescat$quanti[[k]])==FALSE){
          rescat$quanti[k]
          desc_classe <- as.data.frame(rescat$quanti[k])
          desc_classe <- cbind(rownames(desc_classe),desc_classe)
          colnames(desc_classe) <- c("Variable","v.test","Mean in category","Overall mean","sd in category","Overall sd","p.value")
          if (dim(desc_classe)[1]<size_tab){
            ft <- flextable::flextable(desc_classe)
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Quantitative Illustrative Variables"), location = location_title)
            essai <- officer::ph_with(essai, value = ft, location = location_body)
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
          } else {
            n_slide=NULL
            n_slide <- dim(desc_classe)[1]%/%size_tab
            for (i in 0:(n_slide-1)){
              ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Quantitative Illustrative Variables"), location = location_title)
              essai <- officer::ph_with(essai, value = ft, location = location_body)
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
            }
            if (dim(desc_classe)[1]%%size_tab!=0){
              ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Quantitative Illustrative Variables"), location = location_title)
              essai <- officer::ph_with(essai, value = ft, location = location_body)
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
            }
          }
        }
      }
    }
  }
  print(essai, target = paste(path,file.name, sep="/"))
}
