#' Download IBA Protocol Data
#' 
#' This function provides a useful way for querying and downloading Total and Raw IBA Protocol Data for a given site on one or multiple dates.
#' @param IBA String 5 letter alpha numeric site ID for the IBA in which you want to download data. This must be provided.
#' @param file.path String file path to the folder location where you want you data saved. Default will be to your working directory.
#' @param username String username for your NatureCounts account.
#' @return A saved xlsx file with raw and totaled IBA Protocol data.
#' @keywords IBA Protocol, download, NatureCounts
#' @export
#' @examples iba_protocol(IBA = "AB015",file.path = "D:/Key Biodiversity Areas/!KBA_Data/IBAProtocol",username = "sample")
iba_protocol <- function(IBA = NULL, file.path = NULL, username=NULL){
  if (missing(IBA)) stop("You need to provide an IBA code. If you unsure of what codes are available use meta_iba_codes() from the naturecounts package.")
  if (missing(file.path)) {message(paste0("No file.path provided so using current working directory: ",getwd()))
    file.path <- getwd()
  }
  if(!missing(file.path)){
    if(!dir.exists(file.path)) stop("The file path you provided is wrong or does not exist. Make sure file paths are displayed like this: D:\\\\Key Biodiversity Areas\\\\!KBA_Data\\\\IBAfProtocol")
  }
  if (missing(username)) stop("Please provide a username for NatureCounts to download your data!")
  
  data <- naturecounts::nc_data_dl(collections =  c("EBIRD-CA-AT","EBIRD-CA-BC","EBIRD-CA-NO",
                                                    "EBIRD-CA-ON","EBIRD-CA-PR","EBIRD-CA-QC","EBIRD-CA-SENS"),region =list(iba = IBA),fields_set = "core",
                                   username = username,info = "Download IBA Protocol Data.",warn = FALSE, timeout = 320)
  
  Protocol <- data[which(data$ProtocolType=="IBA Canada (protocol)"),]
  if(nrow(Protocol)==0) stop("Sorry there is no IBA Protocol data for this location!")
  datelist <- c("All",sort(unique(Protocol$ObservationDate)))
  datechoice <- utils::select.list(datelist,multiple = T,title = "Please select a date or dates to download or choose them all:")
  if(! "All" %in% datechoice){
    Protocol <- Protocol[which(Protocol$ObservationDate %in% datechoice),]
    if(length(datechoice) == 1){
      datestring <- datechoice
    } else {
      datestring <- "Multipledates"
    }
    
  } else {
    datestring <- "Alldates"
  }
  ### Select Columns for raw Data ###
  Protocol <- subset(Protocol,select = c(GlobalUniqueIdentifier,record_id,iba_site,Locality,SurveyAreaIdentifier,DecimalLatitude,DecimalLongitude,
                                         species_id,CommonName,ScientificName,survey_year,survey_month,survey_day,ObservationDate,TimeObservationsStarted,ObservationCount,
                                         DurationInHours,NumberOfObservers,EffortMeasurement1,EffortUnits1,Remarks,Remarks2,ProtocolType))
  
  ### Create dataframe to hold species sums ###
  speciestotals <- data.frame(iba_site=character(),species_id=integer(),ObservationCount=integer(),CommonName=character(),
                              ScientificName=character(),survey_year=integer(),survey_month=integer(),survey_day=integer(),
                              ObservationDate=character(),Reference=character())
  ### Loop through to get species sums ###
  dates <- unique(Protocol$ObservationDate)
  for (d in 1:length(dates)) {
    datedata <- Protocol[which(Protocol$ObservationDate==dates[d]),]
    species <- unique(datedata$CommonName)
    for (s in 1:length(species)) {
      speciesdata <- datedata[which(datedata$CommonName==species[s]),]
      speciesdata <- subset(speciesdata, select=c("iba_site","Locality","species_id","ObservationCount","CommonName","ScientificName","survey_year","survey_month","survey_day","ObservationDate"))
      speciesdata <- speciesdata[!duplicated(speciesdata),]
      sum <- sum(as.numeric(speciesdata$ObservationCount), na.rm = T)
      speciesdata$ObservationCount <- sum
      speciesdata <- subset(speciesdata, select=c("iba_site","CommonName","ScientificName","ObservationCount","survey_year","survey_month","survey_day","ObservationDate"))
      speciesdata <- speciesdata[!duplicated(speciesdata),]
      speciesdata$Reference <- paste0("Ebird IBA Protocol (",speciesdata$survey_year[1],")") 
      speciestotals <- rbind(speciestotals,speciesdata)
    }
    
  }
  ### Create workbook ###
  wb <- openxlsx::createWorkbook() 
  openxlsx::addWorksheet(wb, sheetName = "SpeciesTotals" )
  openxlsx::addWorksheet(wb, sheetName = "RawData" )
  openxlsx::writeData(wb, "SpeciesTotals", speciestotals)
  openxlsx::writeData(wb, "RawData", Protocol)
  file.name <- file.path(file.path,paste0(IBA,"_IBAProtocolData_",datestring,".xlsx"))
  openxlsx::saveWorkbook(wb, file = file.name, overwrite = TRUE)
  message(paste0("Your file has been saved here: ",file.name))
}