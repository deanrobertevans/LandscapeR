#' Fill in HTML into nomination forms
#'
#' This function fills in the HTML nessisary for the final step in sending nomination forms for upload to KBA-EBAR.
#' @param file.path String file path to the folder location where you want your data read in and saved. Default will be to your working directory.
#' @param rename Character string to append at end of original file name. If not provided the original file will be overwritten. Any files that already contain the renamed string will be ignored or overwritten. For example if "_HTML" is provided AB001.xlsm will be saved as a new file AB001_HTML.xlsm.
#' @return New xlsm files with html text provied in the appropriate cells.
#' @keywords KBA, nomination forms, html
#' @import tidyverse sf xlsx jsonlite geojsonsf
#' @export
#' @examples fill_forms(file.path = "D:/Key Biodiversity Areas/!Forms/FINAL", rename="_HTML" ))

fill_forms <- function(file.path = NULL, rename=""){

  options(java.parameters = "- Xmx5000m")
  if(missing(file.path)){message(paste0("No file.path provided, so using current working directory: ",getwd()))
    file.path <- getwd()
  }
  descriptions <- as.data.frame(jsonlite::fromJSON("https://kba-maps.deanrobertevans.ca/api/htmltext")) %>%
    dplyr::arrange(kbasite_id) %>%
    dplyr::rename(KBAID=kbasite_id,
           SiteCode=sitecode,
           SiteDescriptionE=sitedescription,
           SiteDescriptionF=sitedescription_fr,
           RationaleE=biodiversity,
           RationaleF=biodiversity_fr,
           ConservationE=conservation,
           ConservationF=conservation_fr)
  descriptions <- descriptions %>% dplyr::mutate(SiteCode=convertNL(SiteCode))
  files <-  list.files(file.path, pattern = ".xlsm", all.files = FALSE, recursive = TRUE, full.names = TRUE)
  if(rename!=""){
    files <- files %>% stringr::str_subset(.,rename,negate = T)
  }
  if(is.null(files)) stop("Error: no .xlsm files provided in filepath. Set your filepath or working directory to a folder with .xlsm file(s).")

  for(i in 1:length(files)){
    wb <- xlsx::loadWorkbook(files[i])
    sheets <- xlsx::getSheets(wb)
    #get the site code
    sitecells <- xlsx::getCells(xlsx::getRows(sheets$`2. SITE`))
    if(is.null(sitecells)){
      print(paste0("Warning: site sheet doesn't exist in .xlsm file #", i))
      next
    }
    site <- xlsx::getCellValue(sitecells$`7.3`)
    if(is.na(site)){
      print(paste0("Warning: site code doesn't exist in .xlsm file #", i))
      next
    }
    #using site code, get site description, rationale, and conservation values in english and french
    siteonly <- descriptions %>% dplyr::filter(SiteCode == site)
    if(nrow(siteonly) == 0){
      print(paste0("Warning: ", site, " is not a valid site code in KBA database."))
      next
    }
    #Site description English
    if(!is.na(siteonly$SiteDescriptionE)){
      xlsx::setCellValue(sitecells$`34.3`, specialCharacters(siteonly$SiteDescriptionE))
    }
    if(is.na(siteonly$SiteDescriptionE)){message(paste0("Warning: no value in English description for this site ", site, "."))}
    #Site description French
    if(!is.na(siteonly$SiteDescriptionF)){
      xlsx::setCellValue(sitecells$`34.4`, specialCharacters(siteonly$SiteDescriptionF))
    }
    if(is.na(siteonly$SiteDescriptionF)){message(paste0("Warning: no value in French description for this site ", site, "."))}
    #Rationale English
    if(!is.na(siteonly$RationaleE)){
      xlsx::setCellValue(sitecells$`35.3`, specialCharacters(siteonly$RationaleE))
    }
    if(is.na(siteonly$RationaleE)){message(paste0("Warning: no value in English rationale for this site ", site, "."))}
    #Rationale French
    if(!is.na(siteonly$RationaleF)){
      xlsx::setCellValue(sitecells$`35.4`, specialCharacters(siteonly$RationaleF))
    }
    if(is.na(siteonly$RationaleF)){message(paste0("Warning: no value in French rationale for this site ", site, "."))}
    #Conservation English
    if(!is.na(siteonly$ConservationE)){
      xlsx::setCellValue(sitecells$`39.3`, specialCharacters(siteonly$ConservationE))
    }
    if(is.na(siteonly$ConservationE)){message(paste0("Warning: no value in English conservation for this site ", site, "."))}
    #Conservation French
    if(!is.na(siteonly$ConservationF)){
      xlsx::setCellValue(sitecells$`39.4`, specialCharacters(siteonly$ConservationF))
    }
    if(is.na(siteonly$ConservationF)){message(paste0("Warning: no value in French conservation for this site ", site, "."))}

    #save workbook
    #can change this into an overwrite of the original file by giving it the same name

    xlsx::saveWorkbook(wb, file = paste0(dirname(files[i]),"/",gsub(".xlsm","", basename(files[i])),rename, ".xlsm"))

    #clear environment of worksheet objects and clear memory
    rm("wb", "sheets", "sitecells")
    gc()
  }

}

#' Convert sitecode for NF and LB sites
#'
#' Converts site codes for NF and LB to new NL code
#' @param SiteCode Site code to convert.
#' @return New sitecode.
#' @keywords internal
#' @examples convertNL("NF001")
convertNL <- Vectorize(function(SiteCode){
  if(substr(SiteCode,1,2) %in% c("LB","NF")){
    if(substr(SiteCode,1,2)=="NF") {
      return(paste0("NL",substr(SiteCode,3,5)))
    } else {

      return(paste0("NL",formatC((as.integer(substr(SiteCode,3,5))+46),width = 3, flag = "0")))
    }
  } else {
    return(SiteCode)
  }

})

#' Remove html characters
#'
#' @param x Character string.
#' @return Processed text.
#' @keywords internal
specialCharacters <- function(x){

  # é
  x <- gsub("&eacute;", "é", x)

  # É
  x <- gsub("&Eacute;", "É", x)

  # è
  x <- gsub("&egrave;", "è", x)

  # È
  x <- gsub("&Egrave;", "È", x)

  # ê
  x <- gsub("&ecirc;", "ê", x)

  # Ê
  x <- gsub("&Ecirc;", "Ê", x)

  # à
  x <- gsub("&agrave;", "à", x)

  # À
  x <- gsub("&Agrave;", "À", x)

  # â
  x <- gsub("&acirc;", "â", x)

  # Â
  x <- gsub("&Acirc;", "Â", x)

  # î
  x <- gsub("&icirc;", "î", x)

  # Î
  x <- gsub("&Icirc;", "Î", x)

  # ï
  x <- gsub("&iuml;", "ï", x)

  # Ï
  x <- gsub("&Iuml;", "Ï", x)

  # ù
  x <- gsub("&ugrave;", "ù", x)

  # Ù
  x <- gsub("&Ugrave;", "Ù", x)

  # û
  x <- gsub("&ucirc;", "û", x)

  # Û
  x <- gsub("&Ucirc;", "Û", x)

  # ô
  x <- gsub("&ocirc;", "ô", x)

  # Ô
  x <- gsub("&Ocirc;", "Ô", x)

  # œ
  x <- gsub("&oelig;", "œ", x)

  # Œ
  x <- gsub("&OElig;", "Œ", x)

  # ç
  x <- gsub("&ccedil;", "ç", x)

  # Ç
  x <- gsub("&Ccedil;", "Ç", x)

  # '
  x <- gsub("&rsquo;", "'", x)
  x <- gsub("&apos;", "'", x)

  # -
  x <- gsub("&ndash;", "-", x)

  # «
  x <- gsub("&laquo;", "«", x)

  # »
  x <- gsub("&raquo;", "»", x)

  #
  x <- gsub("&nbsp;", " ", x)

  # Return result
  return(x)
}



