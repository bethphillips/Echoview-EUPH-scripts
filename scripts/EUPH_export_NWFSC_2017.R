### Created by Chelsea Stanley
### Last edited March 15, 2018

### Uses COM objects to run Echoview
### Integrates "EUPH mask" variable to export EUPH NASC

#########################################################

###############################################
#----------------  INPUT  --------------------#
###############################################

# acoustic variables(s) to integrate and their frequency
variables <- c("EUPH export 38kHz","EUPH export 120kHz")
frequency <- c("38","120")

# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)

# set the working directory
#setwd('..'); setwd('..')
BaseYearPath<-"N:/Survey.Acoustics/2017 Hake Sum SH_NP"
BaseJudgePath<-"N:/Survey.Acoustics/2017 Hake Sum SH_NP/Judging/Euphausiids"
BaseProjPath<-"N:/Survey.Acoustics/Projects & Analysis/Euphausiids/EUPH"
#BaseExportPath<-"N:/Survey.Acoustics/2017 Hake Sum SH_NP/Data_SH/Acoustics/Exports_euphausiids/"
setwd(BaseJudgePath)
gdepth=1 # depth for grid (e.g. 1 m)

# location of EV files

#EUPH_EV <- "EUPH_new_template/Ricker 2 frequency"
EUPH_EV <- "EUPH_new_template/Shimada 3 Frequency" # in BaseJudgePath

# Where to put exports
EUPH_exportbase <- "Data_SH/Acoustics/Exports_euphausiids/1_NASC_1m" # in BaseYearPath
date_exportdir<-format(Sys.time(),"%Y%m%d")
EUPH_export<-file.path(EUPH_exportbase,date_exportdir)
dir.create(file.path(BaseYearPath, EUPH_export))
for(v in 1:nrow(vars)){
  var <- vars$variables[v]
  dir.create(file.path(BaseYearPath, EUPH_export,var))
}


#list the EV files to integrate

EVfile.list <- list.files(file.path(BaseJudgePath, EUPH_EV), pattern = ".EV")

# bind variable and frequency together
vars <- data.frame(variables,frequency, stringsAsFactors = FALSE)

# create folder in Exports for each variable
for(f in variables){
  suppressWarnings(dir.create(file.path(BaseJudgePath, EUPH_export, f)))
}

# Loop through EV files 

#for (i in EVfile.list){
for (i in EVfile.list[1:length(EVfile.list)]){
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  EVApp$Minimize()  #keep window in background
  
  # EV filenames to open
  EVfileNames <- file.path(getwd(), EUPH_EV, i)
  EvName <- strsplit(i, split = '*.EV')[[1]]
  
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileNames)
  EVfileName <- file.path(getwd(),EVdir, i)
  print(EVfileName)
  print(i)
  # Variables object
  Obj <- EVfile[["Variables"]]
  
  # loop through variables for integration
  for(v in 1:nrow(vars)){
    var <- vars$variables[v]
    freq <- vars$frequency[v]
    varac <- Obj$FindByName(var)$AsVariableAcoustic()
    
    # Set analysis lines
    Obj_propA<-varac[['Properties']][['Analysis']]
    Obj_propA[['ExcludeAboveLine']]<-"15 m surface blank"
    Obj_propA[['ExcludeBelowLine']]<-"EV bottom pick to edit" 
    
    # Set analysis grid and exclude lines on Sv data
    Obj_propGrid <- varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, gdepth)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
 
    
    # export by cells
    exportcells <- file.path(BaseYearPath, EUPH_export, var, paste(EvName, freq, "cells.csv", sep="_"))
    varac$ExportIntegrationByCellsAll(exportcells)
    
    # Set analysis grid and exclude lines on Sv data back to original values
    Obj_propGrid<-varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, 50)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
    }

  
  
  # save EV file
  EVfile$Save()

  #close EV file
  EVApp$CloseFile(EVfile)
  
  
  #quit echoview
  EVApp$Quit()


## ------------- end loop

}

#####################################################
# Combine all Export .csv
#####################################################


