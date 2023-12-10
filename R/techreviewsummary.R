#' Generate Technical Review Documents
#'
#' This function takes generates word documents for technical reviews of specified sites
#' @param file.path String file path to the folder location where you want the technical reviews saved. Default will be to your working directory.
#' @param sitecodes Character string or vector of IBA sitecodes of tecnical reviews you want generated. Please use the original IBA code for sites in NF and LB.
#' @return Word documents for each technical review done at sites provided.
#' @keywords KBA, technical review, word document
#' @import tidyverse sf geojsonsf googlesheets4 officer officedown flextable magrittr
#' @export
#' @examples fill_forms(file.path = "D:/Key Biodiversity Areas/!Forms/FINAL", rename="_HTML" ))

techreviewsummary <- function(file.path=getwd(),sitecodes=c()){
  flextable::set_flextable_defaults(
    font.family = "Calibri",
    font.size = 11,
    border.color='black',
    split = F
  )
  technicalreview <- googlesheets4::read_sheet('https://docs.google.com/spreadsheets/d/1WPHHKqUb0elfK_FmOr4uhBcXGC9ATn_jmjvSDDujHHI/edit?usp=sharing')

  kbas <- geojsonsf::geojson_sf("https://kba-maps.deanrobertevans.ca/api/allsites") %>% suppressWarnings()

  for(SiteCode in sitecodes){

    sitereviews <- technicalreview %>% dplyr::filter(`1. Selected KBA/IBA to be commented on:` == SiteCode)

    if(nrow(sitereviews)>0){
      site <- kbas %>% dplyr::filter(sitecode==SiteCode)
      for (i in 1:nrow(sitereviews)) {
      message("Doing site: ", site$sitecode)

        reviewdoc <- officer::read_docx(system.file("extdata", "TechnicalReviewTemplate.docx", package="LandscapeR"))
        subdate <- sitereviews$Timestamp[i]
        reviewdoc <- officer::body_add_par(reviewdoc, paste0(site$name_en," Technical Review"), style = "KBAHeading")
        reviewdoc <- officer::body_add_par(reviewdoc, paste0("Survey Response Submitted: ",paste0(lubridate::month(subdate,label = T,abbr = F)," ",lubridate::day(subdate),", ",lubridate::year(subdate))), style = "KBAHeading2")
        reviewdoc <- officer::body_add_par(reviewdoc,"" , style = "KBABolded")
        reviewdoc <- officer::body_add_par(reviewdoc,paste0(site$name_en," Coordinates (lat,long)") , style = "KBABolded")
        reviewdoc <- officer::body_add_par(reviewdoc,paste0(site$latitude,", ",site$longitude) , style = "Normal")
        reviewdoc <- officer::body_add_par(reviewdoc,"" , style = "KBABolded")
        reviewdoc <- officer::body_add_par(reviewdoc,"Reviewer Information" , style = "KBABolded")
        reviewdoc <- officer::body_add_table(reviewdoc,data.frame(Name=paste0(sitereviews$`What is your first name?`[i],
                                                                     " ",
                                                                     sitereviews$`What is your last name?`[i]),
                                                         Email=sitereviews$`What is your email?`[i]),style = "KBATable")
        reviewdoc <- officer::body_add_par(reviewdoc,"" , style = "KBABolded")
        reviewdoc <- officer::body_add_par(reviewdoc,paste0("Technical Review for ",site$name_en) , style = "KBABolded")


        responsequestions <- data.frame(Question=character(),`Yes/Maybe/No`=character(),`Reviewer Comments`=character())
        responsequestions %<>% dplyr::add_row(Question="1. Do you think this site should be a KBA?",
                                       Yes.Maybe.No=dplyr::case_when(sitereviews$`Current IBA/KBA Status:`[i]=="IBA" ~ sitereviews$`2. Do you think this site should be a KBA?...7`[i], sitereviews$`Current IBA/KBA Status:`[i]=="KBA" ~sitereviews$`2. Do you think this site should be a KBA?...9`[i]),
                                       Reviewer.Comments=dplyr::case_when(sitereviews$`Current IBA/KBA Status:`[i]=="IBA"~sitereviews$`Why do you think this IBA should qualify for KBA status?`[i],
                                                                   sitereviews$`Current IBA/KBA Status:`[i]=="KBA"~sitereviews$`Why do you think this KBA might not qualify for KBA status?`[i]))

        responsequestions %<>%
          dplyr::add_row(Question="2. Do you know of any significant trigger species observations missing from this site?",
                  Yes.Maybe.No=sitereviews$`1. Do you know of any significant trigger species observations missing from this site?`[i],
                  Reviewer.Comments=sitereviews$`What species do you think are missing from this site?`[i])

        responsequestions %<>%
          dplyr::add_row(Question="3. Do you think there are any species observations at this site that no longer make sense (i.e. the species is no longer observed there)?",
                  Yes.Maybe.No=sitereviews$`2. Do you think there are any species observations at this site that no longer make sense (i.e. the species is no longer observed there)?`[i],
                  Reviewer.Comments=sitereviews$`Please list any species that do not make sense:`[i])

        responsequestions %<>%
          dplyr::add_row(Question="4. Are any species listed one-offs or vagrant species?  Even though some might be of conservation interest, KBAs can only be designated for “regularly occurring” species.",
                  Yes.Maybe.No=sitereviews$`3. Are any species listed one-offs or vagrant species?  Even though some might be of conservation interest, KBAs can only be designated for “regularly occurring” species.`[i],
                  Reviewer.Comments=sitereviews$`Please list any vagrant species you believe are listed at this site:`[i])

        responsequestions %<>%
          dplyr::add_row(Question="5. Do you know of any data that you or a colleague may have of significant observations at this site? We are looking for data gaps (i.e. newer data or missing species).",
                  Yes.Maybe.No=sitereviews$`4. Do you know of any data that you or a colleague may have of significant observations at this site? We are looking for data gaps (i.e. newer data or missing species).`[i],
                  Reviewer.Comments=sitereviews$`Please explain what data you may have that can help contribute to this KBA. If you know someone who might have data they are willing to share please provide contact information for this individual (Name and Email). Any data you would like to share can be sent to devans@birdscanada.org following the data requirements outlined in the KBA Review Document:`[i])

        responsequestions %<>%
          dplyr::add_row(Question="6. Do you think that the boundaries of this KBA/IBA are accurate? Please note we are not planning any boundary changes until all other biodiversity is included but your thoughts are appreciated.",
                  Yes.Maybe.No=sitereviews$`1. Do you think that the boundaries of this KBA/IBA are accurate? Please note we are not planning any boundary changes until all other biodiversity is included but your thoughts are appreciated.`[i],
                  Reviewer.Comments=sitereviews$`Please explain how the boundary of this KBA/IBA could be improved:`[i])

        responsequestions %<>%
          dplyr::add_row(Question="7. Do you think that the site summary of this KBA/IBA is accurate? Please note we are not updating the site summaries right now, but will contact you at a later date for information regarding changes. To see the site summary, click ‘See Corresponding IBA Website’ in the map viewer.",
                  Yes.Maybe.No=sitereviews$`2. Do you think that the site summary of this KBA/IBA is accurate? Please note we are not updating the site summaries right now, but will contact you at a later date for information regarding changes. To see the site summary, click ‘See Corresponding IBA Website’ in the map viewer.`[i],
                  Reviewer.Comments=NA)
        responsequestions %<>%dplyr:: mutate(Reviewer.Comments=dplyr::case_when(is.na(Reviewer.Comments) ~ "No response provided.",
                                                                  .default = as.character(Reviewer.Comments)))
        italicrows <- which(responsequestions$Reviewer.Comments=="No response provided.")
        names(responsequestions)[2] <-"Yes/Maybe/No"
        names(responsequestions)[3] <-"Reviewer Comments"
        responsetable <- flextable::flextable(responsequestions)
        responsetable <- flextable::theme_box(responsetable)
        responsetable <- flextable::width(responsetable,j=1,width =2)
        responsetable <- flextable::width(responsetable,j=2,width = 1.25)
        responsetable <- flextable::width(responsetable,j=3,width = 3.25)
        responsetable <- flextable::valign(responsetable,j=1:3,valign = 'top')
        responsetable <- flextable::bold(responsetable,j=1)
        responsetable <- flextable::italic(responsetable,i=italicrows,j=3)

        reviewdoc <- flextable::body_add_flextable(reviewdoc,responsetable)
        reviewdoc <- body_add_par(reviewdoc,"" , style = "KBABolded")
        comments <- sitereviews$`3. Please provide any additional comments, questions, concerns about this KBA/IBA:`[i]
        comments <- ifelse(is.na(comments),"No response provided.",comments)
        reviewdoc <- officer::body_add_par(reviewdoc,"Additional Comments:" , style = "KBABolded")
        commentflex <- flextable::flextable(data.frame(comment=comments))
        commentflex <- flextable::theme_box(commentflex)
        commentflex <- flextable::width(commentflex,j=1,width = 6.5)
        commentflex<- flextable::delete_part(commentflex, part = "header")
        commentflex<- flextable::border(commentflex, border.top = officer::fp_border(color = "black") )
        if(comments=="No response provided."){
          commentflex <- flextable::italic(commentflex,i=1,j=1)
        }
        reviewdoc <- flextable::body_add_flextable(reviewdoc,commentflex)
        reviewer <- paste0(c(sitereviews$`What is your first name?`[i],sitereviews$`What is your last name?`[i]),collapse = " ")

        print(reviewdoc, target = paste0(file.path,"/",site$name_en," Technical Review - ",reviewer," - ", lubridate::date(subdate), ".docx"))
      }
    } else {
      message("No reviews for site: ", site$sitecode)
    }
  }
}

