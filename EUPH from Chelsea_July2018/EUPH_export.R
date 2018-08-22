### Created by Chelsea Stanley
### Last edited March 15, 2018

### Uses COM objects to run Echoview
### Integrates "EUPH mask" variable to export EUPH NASC

#########################################################

###############################################
#----------------  INPUT  --------------------#
###############################################

# acoustic variables(s) to integrate and their frequency
variables <- c("CHU export hz38","CHU export hz120")
frequency <- c("38","120")

# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)

# set the working directory
setwd('..'); setwd('..')

# location of EV files

EUPH_EV <- "EUPH_new_template/Ricker 2 frequency"

# Where to put exports
EUPH_export <- "EUPH_new_template/Exports/CHU"
dir.create(file.path(getwd(), EUPH_export))

#list the EV files to integrate

EVfile.list <- list.files(file.path(getwd(), EUPH_EV), pattern = ".EV")

# bind variable and frequency together
vars <- data.frame(variables,frequency, stringsAsFactors = FALSE)

# create folder in Exports for each variable
for(f in variables){
  suppressWarnings(dir.create(file.path(getwd(), EUPH_export, f)))
}

# Loop through EV files 

for (i in EVfile.list){
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # EV filenames to open
  EVfileNames <- file.path(getwd(), EUPH_EV, i)
  EvName <- strsplit(i, split = '*.EV')[[1]]
  
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileNames)
 
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
    Obj_propGrid$SetDepthRangeGrid(1, 10)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
 
    
    # export by cells
    exportcells <- file.path(getwd(), EUPH_export, var, paste(EvName, freq, "cells.csv", sep="_"))
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


