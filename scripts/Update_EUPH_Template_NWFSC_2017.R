### Created by Jessica Nephin
### Modified by Chelsea Stanley
### Modified by Rebecca Thomas
### Last edited Aug 16, 2018

### Uses COM objects to run Echoview
### Exports regions and lines from EV files
### Imports regions and lines into new EV files with new template

### Script runs from location ~Survey/Vessel/Rscripts/EVfiles
#-----------------------------------------------------------------------------------------


# required packages 
require(RDCOMClient)

# set relative working directory
#setwd('..');setwd('..')
BaseYearPath<-"N:/Survey.Acoustics/2017 Hake Sum SH_NP"
BaseJudgePath<-"N:/Survey.Acoustics/2017 Hake Sum SH_NP/Judging/Euphausiids"
BaseProjPath<-"N:/Survey.Acoustics/Projects & Analysis/Euphausiids/EUPH"
setwd(BaseJudgePath)

###############################################
                # INPUT #
###############################################


# Location of original EV files
#EVdir <- "EV/Ricker 2 frequency"
EVdir<-"copy of FINAL for biomass"
#EVdir<-file.path(BaseJudgePath, "copy of FINAL for biomass")

# Location for files updated to EUPH template
EUPH_template <- "EUPH_new_template/Shimada 3 Frequency"
dir.create(file.path(getwd(), EUPH_template))

#tempate name and location
#template <- "RICKER_EUPH_template_2FREQ.EV"
template <- "SHI_EUPH_template_3FREQ.EV"
#Tempdir <- "Scripts/EUPH/Templates/Ricker"
Tempdir <- "Templates/Shimada"

# Name of the bottom line in original EV files
EVbottom <- TRUE
EVbottomname <- "1.0 m bottom offset"

# Does the new template include a bottom line? What is it's name?
bottomline <- TRUE
bottomname <- "EV bottom pick to edit"


###############################################




###################################################
                 # Locate #

#location of calibration file (.ecs)
#CALdir <- "Cal"
CALdir<-"Judging"  #From BaseYearPath

#location of .raw files
#RAWdir <- "RAW"
RAWdir<-"Data_SH/Acoustics/EK60_raw"  #From BaseYearPath

#location for Exports
Exports <- "EUPH_new_template/Exports"
dir.create(file.path(getwd(), Exports))

#location for region exports
Reg <- "EUPH_new_template/Exports/Regions"
dir.create(file.path(getwd(), Reg))

#location for line exports
Line <- "EUPH_new_template/Exports/Lines"
dir.create(file.path(getwd(), Line))

#location for marker region exports
#Marks<-"EUPH_new_template/Exports/Markers"

#########################
# list the EV files to run
EVfile.list <- list.files(file.path(getwd(),EVdir), pattern=".EV")

### move old ev files to old template directory
# file.copy(file.path(getwd(), EVdir, EVfile.list), file.path(getwd(), EVOlddir))
# file.remove(file.path(getwd(), EVdir, EVfile.list))


###############################################
#           Open EV file to update            #
###############################################

for (i in EVfile.list){
  
  # EV filename
  name <- sub(".EV","",i)
  EVfileName <- file.path(getwd(),EVdir, i)
  print(EVfileName)
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  EVApp$Minimize()  #Minimize EV file to run in background
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileName)
  
  # Set fileset object
  filesetObj <- EVfile[["Filesets"]]$Item(0)
  
  # list raw files
  num <- filesetObj[["DataFiles"]]$Count()
  raws <- NULL
  for (l in 0:(num-1)){
    dataObj <- filesetObj[["DataFiles"]]$Item(l)
    dataPath <- dataObj$FileName()
    dataName <- sub(".*\\\\|.*/","",dataPath)
    raws <- c(raws,dataName) 
  }
  
  # get .ecs filename
  calPath <- filesetObj$GetCalibrationFileName()
  calName <- sub(".*\\\\|.*/","",calPath)
  
  # export .evr file
  # filename
  regionfilename <-  file.path(getwd(),Reg, paste(name, "evr", sep="."))
  # export
  EVfile[["Regions"]]$ExportDefinitionsAll(regionfilename)
  
  # export bottom line
  if(EVbottom == TRUE){
  linesObj <- EVfile[["Lines"]]
  bottom <- linesObj$FindbyName(EVbottomname)
  bottomfilename <- file.path(getwd(),Line, paste(name, "bottom", "evl", sep="."))
  bottom$Export(bottomfilename)
  }

  

  #quit echoview
  
  EVApp$Quit()
  

  
  
  #####################################
  #          Make EV file             #
  #####################################
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Open template EV file
  EVfile <- EVApp$OpenFile(file.path(BaseProjPath, Tempdir, template))
  
  # Set fileset object
  filesetObj <- EVfile[["Filesets"]]$Item(0)
  
  # Set calibration file
  if(!calPath == ""){
  add.calibration <- filesetObj$SetCalibrationFile(file.path(BaseYearPath,CALdir, calName))
  }
  
  # Add raw files
  for (r in raws){
    filesetObj[["DataFiles"]]$Add(file.path(BaseYearPath,RAWdir,r))
  }
  
  # Add regions
  EVfile$Import(regionfilename)
  
  # number of editable lines in template
  ls <- NULL
  linesObj <- EVfile[["Lines"]]
  for(k in 0:(linesObj$Count()-1)){
    tmp <- linesObj$Item(k)
    linedit <- tmp$AsLineEditable()
    ls <- c(ls,linedit)
  }
  linenum <- length(ls)
  
  # Add bottom line and overwrite template bottom line if it exists
  EVfile$Import(bottomfilename)
  bottom <- linesObj$FindbyName(paste0("Line",linenum+1))
  linenum <- linenum + 1
  if(bottomline == TRUE){
    oldbottom <- linesObj$FindbyName(bottomname)
    oldbottom$OverwriteWith(bottom)
    linesObj$Delete(bottom)
  } else if(bottomline == FALSE){
    bottom[["Name"]] <- "Bottom"
  }
  
  
  
  
  
  # Save EV file
  EVfile$SaveAS(file.path(getwd(),EUPH_template,i))
  
  # Close EV file
  EVApp$CloseFile(EVfile)
  
  # Quit echoview
  EVApp$Quit()
  
}
