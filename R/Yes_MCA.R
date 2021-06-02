#' PowerPoint reporting of a MCA
#'
#' @param res a MCA object
#' @param yes_study_name title displayed on the first slide
#' @param path path to the directory which will store the PowerPoint file
#' @param file.name name of the PowerPoint file
#' @param x1 component to plot on the x-axis
#' @param x2 component to plot on the y-axis
#' @param thres_x1 the threshold for the x-axis over which points are ploted translucent
#' @param thres_x2 the threshold for the y-axis over which points are ploted translucent
#' @param proba the significance threshold considered to characterized the category (by default 0.05)
#' @param size_tab maximum number of rows of a table per slide
#'
#' @return Returns a .pptx file
#' @export
#'
#' @examples
#' \dontrun{
#' data(tea)
#' res.mca <- FactoMineR::MCA(tea[,-19],quali.sup=19:35,graph=FALSE)
#' # Create the PowerPoint in the current working directory
#' Yes_MCA(res.mca)
#' }

Yes_MCA <- function(res,yes_study_name="MCA + HCPC",path=getwd(),file.name = "MCA_results.pptx",x1=1,x2=2,thres_x1=2,thres_x2=2,proba=0.05,size_tab=10){

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
  essai <- officer::ph_with(essai, value = "Eigenvalues Decomposition", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation
  yes_plot.resp1 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","quali.sup"),label = "none",title = "Representation of the respondents",axes=c(x1,x2))

  #Slide: individuals graph
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the respondents", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp1, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2
  yes_plot.resp2 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("ind","quali.sup"),
                                         label = "var",title = "Representation of the most contributive answers",axes=c(x1,x2),
                                         selectMod = "contrib 15")

  #Slide: most contributive categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the most contributive answers", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp2, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2
  yes_plot.resp3 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup"),label = "var",
                                         title = "Representation of the respondents and the most contributive answers",axes=c(x1,x2),
                                         select = "contrib 0", selectMod = "contrib 15")

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the respondents and the most contributive answers", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp3, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x1
  sel_act_x1_p <- rownames(res$var$v.test[res$var$v.test[,x1]>thres_x1,])
  sel_act_x1_n <- rownames(res$var$v.test[-res$var$v.test[,x1]>thres_x1,])
  sel_act <- unique(c(sel_act_x1_p,sel_act_x1_n))

  if (is.null(sel_act)==TRUE){
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "none",
                                           title = paste("No active answers are significant at a significant level of",thres_x1),
                                           axes=c(x1,x2), select="contrib 0 ", selectMod = sel_act)
  } else {
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "var",
                                           title = paste("Active answers at a significant level of",thres_x1,"for the x-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_act)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant answers for the x-axis", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp4, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x2
  sel_act_x2_p <- rownames(res$var$v.test[res$var$v.test[,x2]>thres_x2,])
  sel_act_x2_n <- rownames(res$var$v.test[-res$var$v.test[,x2]>thres_x2,])
  sel_act <- unique(c(sel_act_x2_p,sel_act_x2_n))

  if (is.null(sel_act)==TRUE){
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "none",
                                           title = paste("No active answers are significant at a significant level of",thres_x2),
                                           axes=c(x1,x2), select="contrib 0 ", selectMod = sel_act)
  } else {
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "var",
                                           title = paste("Active answers at a significant level of",thres_x2,"for the y-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_act)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant answers for the y-axis", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp4, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x1, thres_x2
  sel_act_x1_p <- rownames(res$var$v.test[res$var$v.test[,x1]>thres_x1,])
  sel_act_x1_n <- rownames(res$var$v.test[-res$var$v.test[,x1]>thres_x1,])

  sel_act_x2_p <- rownames(res$var$v.test[res$var$v.test[,x2]>thres_x2,])
  sel_act_x2_n <- rownames(res$var$v.test[-res$var$v.test[,x2]>thres_x2,])

  sel_act <- unique(c(sel_act_x1_p,sel_act_x1_n,sel_act_x2_p,sel_act_x2_n))

  if (is.null(sel_act)==TRUE){
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "none",
                                           title = paste("No active answers are significant at a significant level of",thres_x1,"and",thres_x2),
                                           axes=c(x1,x2), select="contrib 0 ", selectMod = sel_act)
  } else {
    yes_plot.resp4 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("quali.sup","ind"),label = "var",
                                           title = paste("Active answers at a significant level of",thres_x1,"for the x-axis, and",thres_x2,"for the y-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_act)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant answers", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp4, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x1
  sel_ill_x1_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x1]>thres_x1,])
  sel_ill_x1_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x1]>thres_x1,])
  sel_ill <- unique(c(sel_ill_x1_p,sel_ill_x1_n))

  if (is.null(sel_ill)==TRUE){
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "none",
                                           title = paste("No illustrative answers are significant at a level of",thres_x1),axes=c(x1,x2),
                                           select="contrib 0 ")
  } else {
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "quali.sup",
                                           title = paste("Illustrative answers at a significant level of",thres_x1,"for the x-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_ill)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant illustrative answers for the x-axis", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp5, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x2
  sel_ill_x2_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x2]>thres_x2,])
  sel_ill_x2_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x2]>thres_x2,])
  sel_ill <- unique(c(sel_ill_x2_p,sel_ill_x2_n))

  if (is.null(sel_ill)==TRUE){
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "none",
                                           title = paste("No illustrative answers are significant at a level of",thres_x2),axes=c(x1,x2),
                                           select="contrib 0 ")
  } else {
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "quali.sup",
                                           title = paste("Illustrative answers at a significant level of",thres_x2,"for the y-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_ill)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant illustrative answers for the y-axis", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp5, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################


  #######################################
  # Graphical representation: x1, x2, thres_x1, thres_x2
  sel_ill_x1_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x1]>thres_x1,])
  sel_ill_x1_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x1]>thres_x1,])

  sel_ill_x2_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x2]>thres_x2,])
  sel_ill_x2_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x2]>thres_x2,])

  sel_ill <- unique(c(sel_ill_x1_p,sel_ill_x1_n,sel_ill_x2_p,sel_ill_x2_n))

  if (is.null(sel_ill)==TRUE){
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "none",
                                           title = paste("No illustrative answers are significant at a significant level of",thres_x1,"and",thres_x2),axes=c(x1,x2),
                                           select="contrib 0 ")
  } else {
    yes_plot.resp5 <- FactoMineR::plot.MCA(res,choix="ind",invisible=c("var","ind"),label = "quali.sup",
                                           title = paste("Illustrative answers at a significant level of",thres_x1,"for the x-axis, and",thres_x2,"for the y-axis"),axes=c(x1,x2),
                                           select="contrib 0 ", selectMod = sel_ill)
  }

  #Slide: mix individuals and categories
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of the significant illustrative answers", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_plot.resp5, location = officer::ph_location_type(type = "body"))
  yes_slide_num <- yes_slide_num+1
  essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
  #######################################
  #End of first section

  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1), location = officer::ph_location_type(type = "title"))


  #######################################
  #Axis x1 Description by Active Variables
  sel=NULL
  sel_act_x1_p <- rownames(res$var$v.test[res$var$v.test[,x1]>thres_x1,])
  sel_act_x1_n <- rownames(res$var$v.test[-res$var$v.test[,x1]>thres_x1,])
  sel <- c(sel_act_x1_p,sel_act_x1_n)

  if (is.null(sel)==FALSE){

    sel_mod_act <- cbind(res$var$coord[,x1],res$var$cos2[,x1],res$var$v.test[,x1])
    sel_mod_act <- sel_mod_act[sel,]
    sel_mod_act <- as.data.frame(sel_mod_act[order(-sel_mod_act[,3]),])
    sel_mod_act <- cbind(rownames(sel_mod_act),sel_mod_act)
    colnames(sel_mod_act) <- c("Category","Coord","Cos2","v.test")

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_mod_act)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Active Variables"), location = officer::ph_location_type(type = "title"))
      essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))

    } else {
      n_slide=NULL
      n_slide <- dim(sel_mod_act)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_mod_act[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Active Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
      if (dim(sel_mod_act)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_mod_act[(n_slide*size_tab+1):dim(sel_mod_act)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Active Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
    }
  }
  #######################################
  #Axis x1 Description Illustrative Variables

  sel=NULL
  sel_ill_x1_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x1]>thres_x1,])
  sel_ill_x1_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x1]>thres_x1,])
  sel <- c(sel_ill_x1_p,sel_ill_x1_n)

  if (is.null(sel)==FALSE){
    sel_mod_sup <- cbind(res$quali.sup$coord[,x1],res$quali.sup$cos2[,x1],res$quali.sup$v.test[,x1])
    sel_mod_sup <- sel_mod_sup[sel,]
    sel_mod_sup <- as.data.frame(sel_mod_sup[order(-sel_mod_sup[,3]),])
    sel_mod_sup <- cbind(rownames(sel_mod_sup),sel_mod_sup)
    colnames(sel_mod_sup) <- c("Category","Coord","Cos2","v.test")

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_mod_sup)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
      essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
    } else {
      n_slide=NULL
      n_slide <- dim(sel_mod_sup)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_mod_sup[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
      if (dim(sel_mod_sup)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_mod_sup[(n_slide*size_tab+1):dim(sel_mod_sup)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x1,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
    }
  }
  #######################################
  #End of second section

  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2), location = officer::ph_location_type(type = "title"))

  #######################################
  #Axis x2 Description by Active Variables
  sel=NULL
  sel_act_x2_p <- rownames(res$var$v.test[res$var$v.test[,x2]>thres_x2,])
  sel_act_x2_n <- rownames(res$var$v.test[-res$var$v.test[,x2]>thres_x2,])
  sel <- c(sel_act_x2_p,sel_act_x2_n)

  if (is.null(sel)==FALSE){

    sel_mod_act <- cbind(res$var$coord[,x2],res$var$cos2[,x2],res$var$v.test[,x2])
    sel_mod_act <- sel_mod_act[sel,]
    sel_mod_act <- as.data.frame(sel_mod_act[order(-sel_mod_act[,3]),])
    sel_mod_act <- cbind(rownames(sel_mod_act),sel_mod_act)
    colnames(sel_mod_act) <- c("Category","Coord","Cos2","v.test")

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_mod_act)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Active Variables"), location = officer::ph_location_type(type = "title"))
      essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
    } else {
      n_slide=NULL
      n_slide <- dim(sel_mod_act)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_mod_act[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Active Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
      if (dim(sel_mod_act)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_mod_act[(n_slide*size_tab+1):dim(sel_mod_act)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Active Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
    }
  }
  #######################################
  #Axis x2 Description Illustrative Variables

  sel=NULL
  sel_ill_x2_p <- rownames(res$quali.sup$v.test[res$quali.sup$v.test[,x2]>thres_x2,])
  sel_ill_x2_n <- rownames(res$quali.sup$v.test[-res$quali.sup$v.test[,x2]>thres_x2,])
  sel <- c(sel_ill_x2_p,sel_ill_x2_n)

  if (is.null(sel)==FALSE){
    sel_mod_sup <- cbind(res$quali.sup$coord[,x2],res$quali.sup$cos2[,x2],res$quali.sup$v.test[,x2])
    sel_mod_sup <- sel_mod_sup[sel,]
    sel_mod_sup <- as.data.frame(sel_mod_sup[order(-sel_mod_sup[,3]),])
    sel_mod_sup <- cbind(rownames(sel_mod_sup),sel_mod_sup)
    colnames(sel_mod_sup) <- c("Category","Coord","Cos2","v.test")

    if (length(sel)<size_tab){
      ft <- flextable::flextable(sel_mod_sup)
      ft <- flextable::colformat_double(ft, j=-1, digits = 3)
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
      essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
      essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
    } else {
      n_slide=NULL
      n_slide <- dim(sel_mod_sup)[1]%/%size_tab
      for (i in 0:(n_slide-1)){
        ft <- flextable::flextable(sel_mod_sup[(i*size_tab+1):(size_tab*(i+1)),])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
      if (dim(sel_mod_sup)[1]%%size_tab!=0){
        ft <- flextable::flextable(sel_mod_sup[(n_slide*size_tab+1):dim(sel_mod_sup)[1],])
        ft <- flextable::colformat_double(ft, j=-1, digits = 3)
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of Dimension",x2,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
      }
    }
  }
  #######################################
  #End of third section

  #Titre de section
  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Hierarchical Clutering", location = officer::ph_location_type(type = "title"))

  ###################################################
  #    Classification                               #
  ###################################################

  res.hcpc <- FactoMineR::HCPC(res,nb.clust = -1, graph=FALSE)
  sel_act_desc <- rownames(res$var$eta2)
  data_act_class <- res.hcpc$data.clust[,c(sel_act_desc,"clust")]
  res_inter <- FactoMineR::MCA(data_act_class,quali.sup = dim(data_act_class)[2],graph = FALSE)
  yes_res.plot.hcpc <- FactoMineR::plot.MCA(res_inter,habillage = dim(data_act_class)[2],label = "none",invisible = "var",
                                            title = "Representation of respondents according to hierarchical clustering")

  ##########################
  essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Representation of respondents according to hierarchical clustering", location = officer::ph_location_type(type = "title"))
  essai <- officer::ph_with(essai, value = yes_res.plot.hcpc, location = officer::ph_location_type(type = "body"))
  ##########################

  rescat <- FactoMineR::catdes(data_act_class,num.var=dim(data_act_class)[2],proba=proba)
  #Test
  if (is.null(rescat$category)==FALSE){
    nlevels(data_act_class$clust)
    for (k in 1:nlevels(data_act_class$clust)){
      if (is.null(rescat$category[k])==FALSE){
        rescat$category[k]
        desc_classe <- as.data.frame(rescat$category[k])
        desc_classe <- cbind(rownames(desc_classe),desc_classe)
        colnames(desc_classe) <- c("Category","Cla/Mod","Mod/Cla","Global","p.value","v.test")
        if (dim(desc_classe)[1]<size_tab){
          ft <- flextable::flextable(desc_classe)
          ft <- flextable::colformat_double(ft, j=-1, digits = 3)
          ft <- flextable::autofit(ft)
          essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
          essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = officer::ph_location_type(type = "title"))
          essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
          yes_slide_num <- yes_slide_num+1
          essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
        } else {
          n_slide=NULL
          n_slide <- dim(desc_classe)[1]%/%size_tab
          for (i in 0:(n_slide-1)){
            ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
          }
          if (dim(desc_classe)[1]%%size_tab!=0){
            ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Active Variables"), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
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
        if (is.null(rescat$category[k])==FALSE){
          rescat$category[k]
          desc_classe <- as.data.frame(rescat$category[k])
          desc_classe <- cbind(rownames(desc_classe),desc_classe)
          colnames(desc_classe) <- c("Category","Cla/Mod","Mod/Cla","Global","p.value","v.test")
          if (dim(desc_classe)[1]<size_tab){
            ft <- flextable::flextable(desc_classe)
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
          } else {
            n_slide=NULL
            n_slide <- dim(desc_classe)[1]%/%size_tab
            for (i in 0:(n_slide-1)){
              ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
              essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
            }
            if (dim(desc_classe)[1]%%size_tab!=0){
              ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
              essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
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
        if (is.null(rescat$quanti[k])==FALSE){
          rescat$quanti[k]
          desc_classe <- as.data.frame(rescat$quanti[k])
          desc_classe <- cbind(rownames(desc_classe),desc_classe)
          colnames(desc_classe) <- c("Category","v.test","Mean in category","Overall mean","sd in category","Overall sd","p.value")
          if (dim(desc_classe)[1]<size_tab){
            ft <- flextable::flextable(desc_classe)
            ft <- flextable::colformat_double(ft, j=-1, digits = 3)
            ft <- flextable::autofit(ft)
            essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
            essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
            essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
            yes_slide_num <- yes_slide_num+1
            essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
          } else {
            n_slide=NULL
            n_slide <- dim(desc_classe)[1]%/%size_tab
            for (i in 0:(n_slide-1)){
              ft <- flextable::flextable(desc_classe[(i*size_tab+1):(size_tab*(i+1)),])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
              essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
            }
            if (dim(desc_classe)[1]%%size_tab!=0){
              ft <- flextable::flextable(desc_classe[(n_slide*size_tab+1):dim(desc_classe)[1],])
              ft <- flextable::colformat_double(ft, j=-1, digits = 3)
              ft <- flextable::autofit(ft)
              essai <- officer::add_slide(essai, layout = "Titre et texte vertical", master = "YesSiR")
              essai <- officer::ph_with(essai, value = paste("Description of Cluster",k,"with Illustrative Variables"), location = officer::ph_location_type(type = "title"))
              essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_type(type = "body"))
              yes_slide_num <- yes_slide_num+1
              essai <- officer::ph_with(essai, value = yes_slide_num, location = officer::ph_location_type(type = "sldNum"))
            }
          }
        }
      }
    }
  }
  print(essai, target = paste(path,file.name, sep="/"))
}#End function
