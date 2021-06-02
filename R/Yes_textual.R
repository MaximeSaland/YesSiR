#' PowerPoint reporting of a textual analysis
#'
#' @param res FactoMineR::textual result
#' @param yes_study_name title displayed on the first slide
#' @param path PowerPoint file to be created
#' @param file.name name of the PowerPoint file
#' @param proba the significance threshold considered to characterized the category (by default 0.05)
#' @param size_tab maximum number of rows of a table per slide
#' @param col.neg color for negative values (on the wordclouds)
#' @param col.pos color for positive values (on the wordclouds)
#'
#' @return Returns a .pptx file
#' @export
#'
#' @examples
#' \dontrun{
#' data(beard)
#' res.text <- FactoMineR::textual(beard,contingence.by = 1,num.text = 3)
#' # Create the PowerPoint in the current working directory
#' Yes_textual(res.text)
#' }

Yes_textual <- function(res, yes_study_name = "Textual analysis", path=getwd(), file.name = "textual_results.pptx", size_tab=10, proba=0.05,col.neg = "red", col.pos = "blue") {

  location_body <- officer::ph_location_type(type = "body")
  location_title <- officer::ph_location_type(type = "title")
  location_sldNum <- officer::ph_location_type(type = "sldNum")

  yes_temp = system.file("YesSiR_template.pptx", package = "YesSiR")
  #######################################
  ### Init: first slide ----
  yes_slide_num=0
  essai <- officer::read_pptx(yes_temp)
  essai <- officer::remove_slide(essai)
  essai <- officer::add_slide(essai, layout = "Diapositive de titre", master = "YesSiR")
  essai <- officer::ph_with(essai, value = yes_study_name, location = officer::ph_location_type(type = "ctrTitle"))
  #######################################

  #######################################
  # cont.table ----

  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Frequencies", location = location_title)
  word <- data.frame(stimuli = row.names(res$nb.words[1]), freq = res$nb.words$words)
  nbcol <- 5

  if (dim(word)[1]<=size_tab) {
    ft <- flextable::flextable(word)
    ft <- flextable::align(ft,align = "center",part ="header")
    ft <- flextable::bg(ft,j=seq(1,ncol(ft$body$dataset),2),bg="#c7ecee")
    ft <- flextable::bg(ft,j=seq(2,ncol(ft$body$dataset),2),bg="#dff9fb")
    ft <- flextable::autofit(ft)
    essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
    essai <- officer::ph_with(essai, value = ft, location = location_body)
    yes_slide_num <- yes_slide_num+1
    essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  } else if (dim(word)[1]%/%size_tab%/%nbcol==0) {
    ft <- word[1:size_tab,]
    for (i in 1:(dim(word)[1]%/%size_tab)) {
      ft <- cbind(ft,word[(i*size_tab+1):(i*size_tab+size_tab),])
      colnames(ft) <- c(utils::head(colnames(ft),2*i), paste0(colnames(word),i))
    }
    # if (dim(word)[1]%%size_tab!=0) {
    #   filler <- cbind(rep("",size_tab-dim(word)[1]%%size_tab),rep("",size_tab-dim(word)[1]%%size_tab))
    #   colnames(filler) <- colnames(word)
    #   ft <- cbind(ft,rbind(word[(i*size_tab+size_tab+1):dim(word)[1],],filler))
    #   colnames(ft) <- c(utils::head(colnames(ft),2*(i+1)), paste0(colnames(word),i+1))
    # }
    ft <- flextable::flextable(ft)
    ft <- flextable::bg(ft,j=seq(1,ncol(ft$body$dataset),2),bg="#c7ecee")
    ft <- flextable::bg(ft,j=seq(2,ncol(ft$body$dataset),2),bg="#dff9fb")
    ft <- flextable::compose(ft,j=seq(1,ncol(ft$body$dataset),2),value = flextable::as_paragraph("stimuli"), part = "header")
    ft <- flextable::compose(ft,j=seq(2,ncol(ft$body$dataset),2),value = flextable::as_paragraph("freq"), part = "header")
    ft <- flextable::align(ft,align = "center",part ="header")
    ft <- flextable::align(ft,align = "center")
    ft <- flextable::autofit(ft)
    essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
    essai <- officer::ph_with(essai, value = ft, location = location_body)
    yes_slide_num <- yes_slide_num+1
    essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
  } else {
    n_slide=NULL
    n_slide <- dim(word)[1]%/%size_tab%/%nbcol
    for (i in 0:(n_slide-1)) {
      first_row <- i*nbcol*size_tab+1
      last_row <- first_row+size_tab-1
      ft <- word[first_row:last_row,]
      for (j in 1:(nbcol-1)) {
        first_row <- i*nbcol*size_tab+j*size_tab+1
        last_row <- first_row+size_tab-1
        ft <- cbind(ft,word[first_row:last_row,])
        colnames(ft) <- c(utils::head(colnames(ft),2*j), paste0(colnames(word),j))
      }
      ft <- flextable::flextable(ft)
      ft <- flextable::bg(ft,j=seq(1,ncol(ft$body$dataset),2),bg="#c7ecee")
      ft <- flextable::bg(ft,j=seq(2,ncol(ft$body$dataset),2),bg="#dff9fb")
      ft <- flextable::compose(ft,j=seq(1,ncol(ft$body$dataset),2),value = flextable::as_paragraph("stimuli"), part = "header")
      ft <- flextable::compose(ft,j=seq(2,ncol(ft$body$dataset),2),value = flextable::as_paragraph("freq"), part = "header")
      ft <- flextable::align(ft,align = "center",part ="header")
      ft <- flextable::align(ft,align = "center")
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
    }
    if (dim(word)[1]%%(size_tab*nbcol)!=0) {
      first_row <- last_row+1
      last_row <- first_row+size_tab-1
      if (dim(word)[1]%%(size_tab*nbcol) <= size_tab) {
        ft <- word[first_row:nrow(word),]
        colnames(ft) <- colnames(word)
      } else {
        ft <- word[first_row:last_row,]
        for (i in 1:(dim(word)[1]%%(size_tab*nbcol)%/%size_tab)) {
          first_row <- last_row+1
          last_row <- first_row+size_tab-1
          ft <- cbind(ft,word[first_row:last_row,])
          colnames(ft) <- c(utils::head(colnames(ft),2*i),paste0(colnames(word),i))
        }
        # if (dim(word)[1]%%size_tab!=0){
        #   filler <- cbind(rep("",size_tab-dim(word)[1]%%size_tab),rep("",size_tab-dim(word)[1]%%size_tab))
        #   colnames(filler) <- colnames(word)
        #   ft <- cbind(ft,rbind(word[(last_row+1):nrow(word),],filler))
        #   colnames(ft) <- c(utils::head(colnames(ft),2*(i+1)),paste0(colnames(word),i+1))
        # }
      }
      ft <- flextable::flextable(ft)
      ft <- flextable::bg(ft,j=seq(1,ncol(ft$body$dataset),2),bg="#c7ecee")
      ft <- flextable::bg(ft,j=seq(2,ncol(ft$body$dataset),2),bg="#dff9fb")
      ft <- flextable::compose(ft,j=seq(1,ncol(ft$body$dataset),2),value = flextable::as_paragraph("stimuli"), part = "header")
      ft <- flextable::compose(ft,j=seq(2,ncol(ft$body$dataset),2),value = flextable::as_paragraph("freq"), part = "header")
      ft <- flextable::align(ft,align = "center",part ="header")
      ft <- flextable::align(ft,align = "center")
      ft <- flextable::autofit(ft)
      essai <- officer::add_slide(essai, layout = "Titre et contenu", master = "YesSiR")
      essai <- officer::ph_with(essai, value = ft, location = location_body)
      yes_slide_num <- yes_slide_num+1
      essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
    }
  }
  #######################################

  #######################################
  # Descfreq section ----

  essai <- officer::add_slide(essai, layout = "Titre de section", master = "YesSiR")
  essai <- officer::ph_with(essai, value = "Description of the frequencies", location = location_title)

  yes_descfreq <- FactoMineR::descfreq(res$cont.table, proba = proba)

  stimuli <- names(yes_descfreq)

  if (is.null(yes_descfreq)==FALSE) {
    for (stim in stimuli) {
      yes_desc <- as.data.frame(cbind(row.names(yes_descfreq[[stim]]),yes_descfreq[[stim]]))
      colnames(yes_desc) <- c("descriptor",colnames(yes_desc)[-1])
      yes_desc[-1] <- lapply(yes_desc[-1],as.numeric)

      # creation of the wordcloud
      yes_wordcloud <- (ggplot2::ggplot(yes_desc, ggplot2::aes(label = descriptor, size=v.test, colour=v.test)) +
                        ggwordcloud::geom_text_wordcloud(show.legend = TRUE) +
                        ggplot2::scale_size_area(max_size = 10, limits=c(min(yes_desc$v.test),max(yes_desc$v.test))) +
                        ggplot2::theme_minimal() +
                        ggplot2::scale_color_gradient2(low = col.neg, mid = "grey", high = col.pos, limits=c(min(yes_desc$v.test),max(yes_desc$v.test))) +
                        ggplot2::guides(size = FALSE))

      if (dim(yes_desc)[1]<=size_tab) {
        ft <- flextable::flextable(yes_desc)
        ft <- flextable::colformat_double(ft, j=-c(1,4,5), digits = 3)
        ft <- flextable::align(ft,align = "center")
        ft <- flextable::autofit(ft)
        essai <- officer::add_slide(essai, layout = "Deux contenus (large)", master = "YesSiR")
        essai <- officer::ph_with(essai, value = paste("Description of", stim), location = location_title)
        essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
        essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
        yes_slide_num <- yes_slide_num+1
        essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
      } else {
        n_slide=NULL
        n_slide <- dim(yes_desc)[1]%/%size_tab
        for (i in 0:(n_slide-1)){
          ft <- flextable::flextable(yes_desc[(i*size_tab+1):(size_tab*(i+1)),])
          ft <- flextable::colformat_double(ft,j=-c(1,4,5),digits = 3)
          ft <- flextable::align(ft,align = "center")
          ft <- flextable::autofit(ft)
          essai <- officer::add_slide(essai, layout = "Deux contenus (large)", master = "YesSiR")
          essai <- officer::ph_with(essai, value = paste("Description of", stim), location = location_title)
          essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
          # essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
          yes_slide_num <- yes_slide_num+1
          essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
        }
        if (dim(yes_desc)[1]%%size_tab!=0){
          ft <- flextable::flextable(yes_desc[(n_slide*size_tab+1):dim(yes_desc)[1],])
          ft <- flextable::colformat_double(ft,j=-c(1,4,5),digits = 3)
          ft <- flextable::align(ft,align = "center")
          ft <- flextable::autofit(ft)
          essai <- officer::add_slide(essai, layout = "Deux contenus (large)", master = "YesSiR")
          essai <- officer::ph_with(essai, value = paste("Description of", stim), location = location_title)
          essai <- officer::ph_with(essai, value = ft, location = officer::ph_location_label(ph_label = "Content Placeholder 2"))
          # essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
          yes_slide_num <- yes_slide_num+1
          essai <- officer::ph_with(essai, value = yes_slide_num, location = location_sldNum)
        }
        essai <- officer::ph_with(essai, value = yes_wordcloud, location = officer::ph_location_label(ph_label = "Content Placeholder 3"))
      }
    }
  }
  print(essai, target = paste(path,file.name, sep="/"))
}
