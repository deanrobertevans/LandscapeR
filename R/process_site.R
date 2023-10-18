#' KBA Bird Threshold Tests on Single Polygon
#'
#' This function provides a method for querying NatureCounts data for a given polygon and testing against KBA Thresholds.
#' @param site a simple features shape for a single site or polygon. Comes from read_sf.
#' @param location A string location for the site. This is needed for geographically confined populations or subspecies. Can either be the provincial code e.g. 'BC' or a single IBA code e.g. 'BC017.' If ignored these thresholds will be ignored.
#' @param thresholds data.frame of new or current thresholds. If missing will use default values.
#' @param file.name String of file name and path to save your excel file. Must be in .xlsx format.
#' @param username String username for your NatureCounts account.
#' @param timeout integer value for the number of seconds to wait for naturecounts query. Default is 320 seconds so only increase in case of failure.
#' @return A excel spreadsheet with thresholds met for a given polygon. Returns nothing if no thresholds found.
#' @keywords KBA Bird Thresholds, download, NatureCounts, polygon
#' @export
#' @examples process_site(BCSite,location = "BC",thresholds = thresholdsfile, file.name = "BCSiteKBA.xlsx" ))

process_site <- function(site=NULL,location=NULL,thresholds=NULL,file.name=NULL, username=NULL, timeout=320){
  if(missing(site)) stop("Site is missing!")
  if(missing(thresholds)){
    thresholds <- as.data.frame(jsonlite::fromJSON("https://kba-maps.deanrobertevans.ca/api/thresholdstests"))
    colnames(thresholds) <- c('SpeciesID','BCSpeciesID','CommonName_EN','ScientificName','Population','Location','Criteria','ThresholdValue','Threshold','ReproductiveUnits','ThresholdType')
  }
  if(!all(names(thresholds) %in% c('SpeciesID','BCSpeciesID','CommonName_EN','ScientificName','Population','Location','Criteria','ThresholdValue','Threshold','ReproductiveUnits','ThresholdType'))){
    stop("Format of the thresholds table is incorrect. Must contain these columns: SpeciesID, BCSpeciesID, CommonName_EN, ScientificName, Population, Location, Criteria, ThresholdValue, Threshold, ReproductiveUnits, & ThresholdType.")
  }
  if(missing(file.name)) stop("Missing file path. Must be in .xlsx format!")
  if(!grepl("\\.xlsx$", file.name)) stop("file.name Must be in .xlsx format!")
  if (missing(username)) stop("Please provide a username for NatureCounts to download your data!")
  ### Create bounding box ###
  site <-  sf::st_transform(site,crs = 4326) ### Transform First
  bbox <- sf::st_bbox(site)

  ### Get data and process ###
  data <- naturecounts::nc_data_dl(region = list(bbox = bbox),fields_set = "core",
                                   username = username, info = "KBA Assessment",warn = FALSE, timeout = timeout)
  data <- subset(data, select=c("species_id","ObservationCount","CommonName","ScientificName","latitude","longitude","survey_year","survey_month","survey_day","CollectionCode"))
  data$ObservationCount <- as.numeric(data$ObservationCount)
  data <- data[!is.na(data$ObservationCount),]
  data <- data[!is.na(data$survey_year),]
  data <- data[!is.na(data$CommonName),]
  data <- data[!is.na(data$ScientificName),]
  data <- data[!is.na(data$latitude),]
  data <- data[!is.na(data$longitude),]
  data <- data[,1:10]

  ### Exclude data by shapefile
  data <- sf::st_as_sf(data, coords = c("longitude", "latitude"),   crs = 4326,remove =F)
  data <- sf::st_drop_geometry(sf::st_intersection(data,site))
  data <- subset(data, select=c("species_id","ObservationCount","CommonName","ScientificName","latitude","longitude","survey_year","survey_month","survey_day","CollectionCode"))
  ### Create empty datafame for species triggers ###
  triggerspecies <- data.frame(CommonName=character(),ScientificName=character(),Population=character(),Number=character(),Date=character(),NumberOfObservations=numeric(),
                               Criteria=character(),DataIssues=character())
  if(!missing(location)){
    for (k in 1:length(location)) {
      if(nchar(location[k])==5){
        if(file.exists(paste0("Z:\\Key Biodiversity Areas\\!KBA_Data\\IBA_Data\\Processed\\",location[k],".csv"))){
          IBAData <- read.csv(paste0("Z:\\Key Biodiversity Areas\\!KBA_Data\\IBA_Data\\Processed\\",location[k],".csv"))
          IBAData$latitude <- NA
          IBAData$longitude <- NA
          colnames(IBAData)[colnames(IBAData)=="Reference"] <- "CollectionCode"
          IBAData <- IBAData[!is.na(IBAData$ObservationCount),]
          IBAData <- IBAData[!is.na(IBAData$survey_year),]
          IBAData <- subset(IBAData,select = c("species_id","ObservationCount","CommonName","ScientificName","latitude","longitude","survey_year","survey_month","survey_day","CollectionCode"))
          data <- rbind(data,IBAData)
          data <- data[!duplicated(data[,c("CommonName","ObservationCount","survey_year","survey_month","survey_day")]),]
        }
      }
    }
  }

  for (i in 1:nrow(thresholds)) {
    if(!(is.na(thresholds$Location[i]) | thresholds$Location[i] == "")){
      if(!missing(location)){
        locationlist <- strsplit(thresholds$Location[i], ',')[[1]]
        if(nchar(locationlist[1])==2){
          if(any(substr(location, start = 1, stop = 2) %in% locationlist)){
            speciesdata <- data[which(data$species_id==thresholds$BCSpeciesID[i] | data$CommonName == thresholds$CommonName[i] | data$ScientificName == thresholds$ScientificName[i]),]
            if(nrow(speciesdata)>0){

              if(!is.na(thresholds$ReproductiveUnits[i])){
                speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
                speciesdata <- speciesdata[which(speciesdata$ObservationCount >= (thresholds$ReproductiveUnits[i]*2)),]
              } else {
                speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
              }
              if(nrow(speciesdata)>0){
                Deficient <- vector()
                if(max(speciesdata$survey_year)<2009){Deficient <- c(Deficient,"Out of Date")}
                if(length(unique(speciesdata$survey_year))==1){Deficient <- c(Deficient,"One Year of Data")}
                triggerspecies <- rbind(triggerspecies,data.frame(CommonName=thresholds$CommonName[i],ScientificName=thresholds$ScientificName[i],Population=thresholds$Population[i],
                                                                  Date=if(length(unique(speciesdata$survey_year))==1){paste0(unique(speciesdata$survey_year))}else{paste0(min(speciesdata$survey_year),"-",max(speciesdata$survey_year))},
                                                                  NumberOfObservations=nrow(speciesdata),
                                                                  Criteria=thresholds$Criteria[i],
                                                                  Number=if(length(unique(speciesdata$ObservationCount))==1){paste0(unique(speciesdata$ObservationCount))}else{paste0(min(speciesdata$ObservationCount),"-",max(speciesdata$ObservationCount))},
                                                                  DataIssues=paste0(Deficient,collapse = "; ") ))


              }

            }
          }
        } else {
          if(nchar(locationlist[1])==5){
            if(any(location %in% locationlist)){
              speciesdata <- data[which(data$species_id==thresholds$BCSpeciesID[i] | data$CommonName == thresholds$CommonName[i] | data$ScientificName == thresholds$ScientificName[i]),]
              if(nrow(speciesdata)>0){

                if(!is.na(thresholds$ReproductiveUnits[i])){
                  speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
                  speciesdata <- speciesdata[which(speciesdata$ObservationCount >= (thresholds$ReproductiveUnits[i]*2)),]
                } else {
                  speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
                }
                if(nrow(speciesdata)>0){
                  Deficient <- vector()
                  if(max(speciesdata$survey_year)<2009){Deficient <- c(Deficient,"Out of Date")}
                  if(length(unique(speciesdata$survey_year))==1){Deficient <- c(Deficient,"One Year of Data")}
                  triggerspecies <- rbind(triggerspecies,data.frame(CommonName=thresholds$CommonName[i],ScientificName=thresholds$ScientificName[i],Population=thresholds$Population[i],
                                                                    Date=if(length(unique(speciesdata$survey_year))==1){paste0(unique(speciesdata$survey_year))}else{paste0(min(speciesdata$survey_year),"-",max(speciesdata$survey_year))},
                                                                    NumberOfObservations=nrow(speciesdata),
                                                                    Criteria=thresholds$Criteria[i],
                                                                    Number=if(length(unique(speciesdata$ObservationCount))==1){paste0(unique(speciesdata$ObservationCount))}else{paste0(min(speciesdata$ObservationCount),"-",max(speciesdata$ObservationCount))},
                                                                    DataIssues=paste0(Deficient,collapse = "; ") ))


                }

              }
            }
          }
        }
      }

    } else {
      speciesdata <- data[which(data$species_id==thresholds$BCSpeciesID[i] | data$CommonName == thresholds$CommonName[i] | data$ScientificName == thresholds$ScientificName[i]),]
      if(nrow(speciesdata)>0){

        if(!is.na(thresholds$ReproductiveUnits[i])){
          speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
          speciesdata <- speciesdata[which(speciesdata$ObservationCount >= (thresholds$ReproductiveUnits[i]*2)),]
        } else {
          speciesdata <- speciesdata[which(speciesdata$ObservationCount >= thresholds$Threshold[i]),]
        }
        if(nrow(speciesdata)>0){
          Deficient <- vector()
          if(max(speciesdata$survey_year)<2009){Deficient <- c(Deficient,"Out of Date")}
          if(length(unique(speciesdata$survey_year))==1){Deficient <- c(Deficient,"One Year of Data")}
          triggerspecies <- rbind(triggerspecies,data.frame(CommonName=thresholds$CommonName[i],ScientificName=thresholds$ScientificName[i],Population=thresholds$Population[i],
                                                            Date=if(length(unique(speciesdata$survey_year))==1){paste0(unique(speciesdata$survey_year))}else{paste0(min(speciesdata$survey_year),"-",max(speciesdata$survey_year))},
                                                            NumberOfObservations=nrow(speciesdata),
                                                            Criteria=thresholds$Criteria[i],
                                                            Number=if(length(unique(speciesdata$ObservationCount))==1){paste0(unique(speciesdata$ObservationCount))}else{paste0(min(speciesdata$ObservationCount),"-",max(speciesdata$ObservationCount))},
                                                            DataIssues=paste0(Deficient,collapse = "; ") ))


        }

      }


    }

  }
  if(nrow(triggerspecies)>0){
    wb <- openxlsx::createWorkbook()

    ## Create the worksheets
    openxlsx::addWorksheet(wb, sheetName = "Triggers" )
    openxlsx::addWorksheet(wb, sheetName = "RawData" )
    openxlsx::writeData(wb, "Triggers", triggerspecies)
    openxlsx::writeData(wb, "RawData", data)
    openxlsx::saveWorkbook(wb, file = file.name, overwrite = TRUE)
    message("Triggers found at this site!")
    return(TRUE)
  } else {
    message("No triggers found at this site.")
    return(FALSE)}
}
