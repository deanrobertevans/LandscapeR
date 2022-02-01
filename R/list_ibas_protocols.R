#' Get a list of sites with IBA Protocol Data
#' 
#' This function provides a useful way for querying and downloading Total and Raw IBA Protocol Data for a given site on one or multiple dates.
#' @param Province String or character vector containing province codes for which you want to check for IBAs with IBA Protocol data.
#' @return A dataframe of IBAs with IBA Protocol Data
#' @keywords IBA Protocol, download, NatureCounts
#' @export
#' @examples list_ibas_protocols(Procince=c("BC","ON"))
list_ibas_protocols <- function(Province=NULL){
  f <- system.file("data", paste0("IBAProtocol", ".rds"), package = "LandscapeR")
  if(!file.exists(f)) stop("Could not find IBA Protocol Metadata!")
  IBAProtocol <- readRDS(f)
  if(missing(Province)){
    print.data.frame(IBAProtocol,row.names = F)
  } else {
    IBAProtocolProvince <- IBAProtocol[which(IBAProtocol$Province %in% Province),]
    if(nrow(IBAProtocolProvince)==0){stop("Sorry there is no IBA Protocol Data for this province.")}
    print.data.frame(IBAProtocolProvince,row.names = F)
  }
  
}