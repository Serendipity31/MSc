## USE R 3.3.2 64 bit version (important!)
##Notes for future...
###R starts with column 0, but row 1 for data frame with headers

#Insructions are found in the 'Information to Use R script.txt' file

# DO NOT CHANGE THIS SCRIPT EXCEPT FOR THE VALUE OF THE YEAR IN STEP 1B!!!
#_______________________________________________________________________________________________________________________________

###PART 1: SET UP

# Step 1a: Set working directory to match the folder from Step 0, changing the code below as required, and load package xlsx

setwd("C:/Users/Owner/Desktop/SRUC MOU Calculation/2016_2017")

#at work: setwd("G:/CKD BAULCOMB/MSc Involvement/SRUC MOU Calculation/2016_2017")
library(xlsx)

# Step 1b: Initialise the following objects
Programmes <- c("EE", "EPM", "FS", "SS", "SPH")	
Research_Groups <- c("LEES", "CropsSoils")

###PART 2: Calculations related to teaching

# Step 1: Import data reference data (e.g. year, course names, tuition fee schedule, and fee status data)

ImportTuitionData <- function() {
	#Imports file containing year in which MSc commencses (e.g. 2016 for the 2016-2017 acadmeic year)
	yr <<- read.xlsx("Inputs/ReferenceInfo/Year_for_Calculation.xlsx", sheetIndex=1, rowIndex=1, colIndex=1, header=FALSE)

	#Trim trailing whitespace in case this appears
		## Source of this approach is: http://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace-in-r
		### Look for sub-comment by Thieme Hennis Sep 19 '14 
	yr <<- as.data.frame(apply(yr,2,function (x) sub("\\s+$", "", x)))
	yr <<- yr[1,1]
	
	# Import the data table showing all SRUC courses, programmes that own them, RGs, and credits/weightings
	SRUC_Courses <<- read.xlsx("Inputs/ReferenceInfo/SRUC_Courses.xlsx", sheetIndex=1, header=TRUE, as.data.frame=TRUE)
	#Trim trailing whitespace in case this appears
		## Source of this approach is: http://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace-in-r
		### Look for sub-comment by Thieme Hennis Sep 19 '14 
	#SRUC_Courses <<- as.data.frame(apply(SRUC_Courses,2,function (x) sub("\\s+$", "", x)))
	
	Courses <<- SRUC_Courses[,2]
	Programme_Ownership <<- SRUC_Courses[,3]
	Credit_Weighting <<- SRUC_Courses[,6]
	
	# Import file as data frame showing all programmes, School, Home Fees (FT), Overseas Fees (FT)
	# csv version: TuitionFees <<- as.data.frame(read.csv("Inputs/Fees_2016.csv", header=TRUE, sep=","))
	TuitionFees <<- read.xlsx("Inputs/ReferenceInfo/Fees_2016.xlsx", sheetIndex=1, header=TRUE, as.data.frame=TRUE)
	
	#Trim trailing whitespace that appear to exist in the "Programme" columns (as this inhibits merging later)
		## Source of this approach is: http://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace-in-r
		### Look for sub-comment by Thieme Hennis Sep 19 '14 
		
		
		#Keep only 1st 5 columns to remove ODL and APC and any other fee info that's not useful
	TuitionFees <<-	TuitionFees[1:5]
		
	# Rename column showing programme name
	names(TuitionFees)[names(TuitionFees)=="Name.of.Programme"] <<- "Programme"
	# Put all of the fee related information within one column (this is necessary for later)
	textcols <<- TuitionFees[,1:3]

	textcolsws <<- as.data.frame(apply(textcols,c(1,2),function (x) sub("\\s+$", "", x)))
	allcols <<- cbind(textcolsws, TuitionFees[,4:5])

	TuitionFees_stacked <<- cbind(allcols[gl(nrow(TuitionFees), 1, 2*nrow(allcols)), 1:3], stack(allcols[,4:5]))
	#Delete Programme Code column to ensure the stacking function works below
	TuitionFees_stacked <<-	TuitionFees_stacked[-2]

	##Rename the columns from the defaults to what they are to allow merging later
	names(TuitionFees_stacked)[3] <<- "Tuition"
	names(TuitionFees_stacked)[4] <<- "Fee_Status"

}
ImportTuitionData()
TuitionFees_stacked[1:10,] #checks function has worked
as.character(TuitionFees_stacked[1, 2])

					     
ImportFeeStatusData <- function() {

	# Import the datafile showing the fee status determined by admissions for all students in 5 schools (CFUF/UF)
	FeeStatus <<- read.xlsx("Inputs/ReferenceInfo/FeeStatus_2016.xlsx", sheetIndex=1, header=TRUE, as.data.frame=TRUE)
	#Trim trailing whitespace in case this appears
		## Source of this approach is: http://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace-in-r
		### Look for sub-comment by Thieme Hennis Sep 19 '14 
	#FeeStatus <<- as.data.frame(apply(FeeStatus,2,function (x) sub("\\s+$", "", x)))
	FeeStatus <<- FeeStatus[,1:15]
	#At this point, all the students are in the list, so need to select the subset consisting of all part time students in
	# these schools, and export them to a file that can be used as the template for next year to ensure no one is missed out.
	## In 2016, will have to add Sydney Chandler in by hand (as the only one I know who stiched status from FT to PT
	ptstudents <<- subset(FeeStatus, grepl("/*P$", FeeStatus$Prog, ignore.case=TRUE))
	#Export this file so that it's ready to go for next year
	write.xlsx(ptstudents, paste("Outputs/FutureInputs/PTStudent_from_FeeStatus_", yr, ".xlsx", sep=""))
	# Search for and remove dupliate rows if students 'hang' around in the admissions system for years
		   #this function should look at all columns to determine uniqueness, and then remove full rows
		   #http://stats.stackexchange.com/questions/6759/removing-duplicated-rows-data-frame-in-r
	FeeStatus <<- FeeStatus[!duplicated(FeeStatus),]	   
	# convert UUN column within FeeStatus to character to match with CourseData column
	FeeStatus$UUN <<- as.character(FeeStatus$UUN)
}

ImportFeeStatusData()
FeeStatus[1:10,] #checks function has worked
	

# Step 2: Iteratively import attendance lists, rearrange them, and merge with fee status,  tuition fee, and credit information to 
# enable the calculation of course fee owed to each programme per student on the courses (irrespective of PT status or school of
# origin)
		
ImportClassData <- function () {
	CourseData <<- NULL
	i = 1
	Lost_Student_Check <<- data.frame(Courses=character(), Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
	CourseData <<- vector('list', length(Courses))
		
	## Imports attendance list
	while (i <= length(Courses)) {
		## Imports attendance list
		fn <- paste("Inputs/Classes/", Courses[i], "_CLASS_LIST_", yr, ".xlsx", sep="")
		## Creates dataframe associated with course that holds position i in Courses
		CourseData[[i]] <<-read.xlsx(fn, sheetIndex=1, header=TRUE, as.data.frame=TRUE)		
		
		## renames columns...not sure if need as [,1] or [1]...test at work was just [1] on R.3.3.2
		names(CourseData[[i]])[1]<<-"UUN"
		names(CourseData[[i]])[2]<<-"Surname"
		names(CourseData[[i]])[3]<<-"Forename"
		names(CourseData[[i]])[4]<<-"Programme"
		names(CourseData[[i]])[5]<<-"Matriculation"
		names(CourseData[[i]])[6]<<-"Enrollment"
		names(CourseData[[i]])[7]<<-"School"
		names(CourseData[[i]])[8]<<-"Signature"
		## removes all columns after the 'School' column		
		CourseData[[i]] <<-CourseData[[i]][1:7]
		# Need to remove last 2 characters from UUN column
		CourseData[[i]]$UUN <<- as.character(CourseData[[i]]$UUN)
		CourseData[[i]]$UUN <<- substr(CourseData[[i]]$UUN, 1, nchar(CourseData[[i]]$UUN)-2)
		CourseData[[i]]$UUN <<- as.character(CourseData[[i]]$UUN)
		# replace school names used by default in attendance list with ones matching the what I want as table headers 
			### Note: if need to match Fee Status sheet spellings, the only one to change is SPSS to "Social & Political Science"
		# used this tip: http://stackoverflow.com/questions/22418864/r-replace-entire-strings-based-on-partial-match
		# Need to specify column and/or change it to character from factor to get to work...
		CourseData[[i]]$School <<- as.character(CourseData[[i]]$School)
		CourseData[[i]]$School[grepl("School Of Geosciences", CourseData[[i]]$School, ignore.case=FALSE)] <<- "GeoSciences"
		CourseData[[i]]$School[grepl("School Of Social And Political Science", CourseData[[i]]$School, ignore.case=FALSE)] <<- "SPSS"
		CourseData[[i]]$School[grepl("School Of Engineering", CourseData[[i]]$School, ignore.case=FALSE)] <<- "Engineering"
		CourseData[[i]]$School[grepl("School Of Law", CourseData[[i]]$School, ignore.case=FALSE)] <<- "Law"
		CourseData[[i]]$School[grepl("Business School", CourseData[[i]]$School, ignore.case=FALSE)] <<- "Business"
		CourseData[[i]]$School[grepl("Edinburgh College Of Art", CourseData[[i]]$School, ignore.case=FALSE)] <<- "Art"
		# remove any rows where there is a PhD student enrolled, as they should not be enrolled
		CourseData[[i]] <<- CourseData[[i]][!grepl("PhD", CourseData[[i]]$Programme),]
		# remove any rows with an auditing student, as they shouldn't be counted
		CourseData[[i]] <<- CourseData[[i]][!grepl("Audit", CourseData[[i]]$Enrollment),]
		
		i=i+1
	}
	
	names(CourseData) <<- Courses
}

ImportClassData()
CourseData[[19]]
CourseData["FEE"]
CourseData[20]

MergeClassData_FeeStatus <- function () {

	i = 1
	CourseDataFS <<- vector('list', length(Courses))
			
	## Sets up merger with fee status data
	while (i <= length(Courses)) {

		# Need to document the number of MSc students within the course before merging, in case the FeeStatus data
			#is incomplete
			Pre_Merge_Length <- length(CourseData[[i]]$UUN)
			Pre_Merge_UUN <- as.vector(CourseData[[i]]$UUN)
		
		## Change case from 'S' to 's' in Fee Status data frame to match with Course Data
		FeeStatus$UUN <<- sapply(FeeStatus$UUN, tolower)
		## Merges attendance list with fee status information
		CourseDataFS[[i]] <<-merge(CourseData[[i]], FeeStatus[ , c("UUN", "FSG")], by=c("UUN"))
		CourseDataFS[[i]] <<- CourseDataFS[[i]][!duplicated(CourseDataFS[[i]]),]
		# Need to now check if any students no longer appear in the data frame. If they don't appear it is because
			# they weren't in the Fee Status sheet...and the most likely explanation for that is they are pursuing their
			# studies for longer than initially anticipated (e.g. have had an interruption or concession to change status)
			# If this happens, this should print a warning to prompt us to go back and find the missing student data and
			# add it into the FeeStatus sheet if appropriate (i.e. unless they have already paid all their tuition AND we
			# have already been paid for it, and it's just the delay in them actually participating in whichever course
			Post_Merge_Length <- length(CourseDataFS[[i]]$UUN)
			Post_Merge_UUN <- as.vector(CourseDataFS[[i]]$UUN)
			Diff <- Post_Merge_Length - Pre_Merge_Length
			
			if (Diff <0) {
				Highlights <- paste("Warning:", abs(Diff) , "Student(s) LOST during merge")
			}
			else {
				Highlights <- ""
			}
			
			# Code from: http://stackoverflow.com/questions/17598134/compare-two-lists-in-r
			# Look for Teemu Daniel Laajala
			Inboth <- Pre_Merge_UUN[Pre_Merge_UUN %in% Post_Merge_UUN] # in both, same as call: intersect(first, second)
			OnlyInPreMerge <- Pre_Merge_UUN[!Pre_Merge_UUN %in% Post_Merge_UUN] # only in 'first', same as: setdiff(first, second)
			OnlyInPostMerge <- Post_Merge_UUN[!Post_Merge_UUN %in% Pre_Merge_UUN] # only in 'second', same as: setdiff(second, first)
			
			#For reference on listing lost UUNs in final column: 
			# http://stackoverflow.com/questions/13973116/convert-r-vector-to-string-vector-of-1-element
			Lost_Student_Check[i,] <<- c(Courses[i], Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
					 
		# Rename FSG column to be "Fee_Status"
		names(CourseDataFS[[i]])[8]<<-"Fee_Status"
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		CourseDataFS[[i]]$Fee_Status[grepl("RUK|SEU", CourseDataFS[[i]]$Fee_Status, ignore.case=FALSE)] <<- "H"
		i=i+1
	}
	names(CourseDataFS) <<- Courses

}	

MergeClassData_FeeStatus()
CourseDataFS
CourseDataFS[19]
					  
MergeClassFeeStatus_TuitionInfo <- function () {
	
	i = 1
	CourseDataFSTI <<- vector('list', length(Courses))

			
	## Sets up merger with tuition information
	while (i <= length(Courses)) {
		
		# Then complete merger	
		CourseDataFSTI[[i]] <<-merge(CourseDataFS[[i]], TuitionFees_stacked[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		CourseDataFSTI[[i]] <<- CourseDataFSTI[[i]][!duplicated(CourseDataFSTI[[i]]),]

		## Inputs credit weighting for course 
		CourseDataFSTI[[i]][,10]<<-(Credit_Weighting[[i]])
		## Names this column
		names(CourseDataFSTI[[i]])[names(CourseDataFSTI[[i]])=="V10"]<<-"Credit_Weighting"
		## Re-orders attendance list with fee information so it's easier to read
		CourseDataFSTI[[i]] <<-CourseDataFSTI[[i]][c("UUN", "Surname", "Forename", "Programme", "School", "Matriculation", "Enrollment", "Fee_Status", "Tuition", "Credit_Weighting")]
				
		## Calculates portion of total fee associated with each student on the course
		## Builds the Course_Fee column and FTE Tuition column to ensure PT students pay same for course as FT students
		CourseDataFSTI[[i]]$Course_Fee <<- ifelse(grepl("3 Years", CourseDataFSTI[[i]]$Programme, ignore.case=TRUE), 
											0.125 * 3 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting)), 
											ifelse(grepl("2 Years", CourseDataFSTI[[i]]$Programme, ignore.case=TRUE), 
											0.125 * 2 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting)), 
											0.125 * 1 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting))))
		
		CourseDataFSTI[[i]]$Net_TopSlice <<- 0.80 * as.numeric(as.character(CourseDataFSTI[[i]]$Course_Fee))
		CourseDataFSTI[[i]]$Net_GSAdmin <<- 0.85 * CourseDataFSTI[[i]]$Net_TopSlice
		CourseDataFSTI[[i]]$Net_PGO <<- 0.90 * CourseDataFSTI[[i]]$Net_GSAdmin
		#Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	
	names(CourseDataFSTI) <<- Courses
	write.xlsx(Lost_Student_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep="" ), sheetName="Courses", append=TRUE)			 
}

MergeClassFeeStatus_TuitionInfo()
CourseDataFSTI[[19]]
	  
					  
#######################Start HERE###################					  
					  
#Step 2: Income associated with teaching individual courses
Course_Level_Finances <- function() {
	
	i = 1
	
	CourseFinances <<- data.frame(All_Schools=numeric(), GeoSciences=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), Art=numeric(), stringsAsFactors=FALSE)
	
	while (i <= length(Courses)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		gs <- subset(CourseDataFSTI[[i]], School == "GeoSciences")
		spss <- subset(CourseDataFSTI[[i]], School == "SPSS")
		law <- subset(CourseDataFSTI[[i]], School == "Law")			
		eng <- subset(CourseDataFSTI[[i]], School == "Engineering")
		bus <- subset(CourseDataFSTI[[i]], School == "Business")
		art <- subset(CourseDataFSTI[[i]], School == "Art")
			
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All_Courses <- sum(CourseDataFSTI[[i]]$Course_Fee)
		Total_gs <- sum(gs$Course_Fee)
		Total_spss <- sum(spss$Course_Fee)
		Total_law <- sum(law$Course_Fee)
		Total_eng <- sum(eng$Course_Fee)
		Total_bus <- sum(bus$Course_Fee)
		Total_art <- sum(art$Course_Fee)
		
		CourseFinances[i,] <<- list(Total_All_Courses, Total_gs,Total_spss, Total_law, Total_eng, Total_bus, Total_art)
					
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	CourseFinances <<- data.frame(Courses, Programme_Ownership, CourseFinances)
	CourseFinances_Totals <<- data.frame("All_Courses", "All_Programmes", sum(CourseFinances$All_Schools), sum(CourseFinances$GeoSciences),sum(CourseFinances$SPSS),sum(CourseFinances$Law),sum(CourseFinances$Engineering),sum(CourseFinances$Business),sum(CourseFinances$Art))
	names(CourseFinances_Totals) <<- names(CourseFinances)
	CourseFinances <<- rbind(CourseFinances, CourseFinances_Totals)
}

#To Check inputs worked
Course_Level_Finances()
CourseFinances
CourseFinances[7,]

CourseFinances_Totals
	
#Step 3: Income associated with teaching for individual programmes
Programme_Level_Finances_Teaching <- function() {
	
	i = 1
	
	#Creates subsets of CourseFinances to match the courses owned by individual programmes
	ProgrammeData_TC <<- vector('list', length(Programmes))
	
	while(i <= length(Programmes)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		ProgrammeData_TC[[i]] <<- subset(CourseFinances, Programme_Ownership == as.character(Programmes[[i]]))
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	#Allows one to call summary finance table by programme name
	names(ProgrammeData_TC) <<- Programmes

	#Creates a summary table of programme finances
	j=1
	ProgrammeFinances_TC <<- data.frame(All_Schools=numeric(), GeoSciences=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), Art=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		#Within the dataframe showing courses owned by programmes, sum the relevant tuition fee components by column
		Total_All_Schools <- sum(ProgrammeData_TC[[j]]$All_Schools)
		Total_gs <- sum(ProgrammeData_TC[[j]]$GeoSciences)
		Total_spss <- sum(ProgrammeData_TC[[j]]$SPSS)
		Total_law <- sum(ProgrammeData_TC[[j]]$Law)
		Total_eng <- sum(ProgrammeData_TC[[j]]$Engineering)
		Total_bus <- sum(ProgrammeData_TC[[j]]$Business)
		Total_art <- sum(ProgrammeData_TC[[j]]$Art)
		
		ProgrammeFinances_TC[j,] <<- list(Total_All_Schools, Total_gs,Total_spss, Total_law, Total_eng, Total_bus, Total_art)
		row.names(ProgrammeFinances_TC)[j] <<- Programmes[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
	#Append summary row to table
	ProgrammeFinances_TC_Totals <<- data.frame(sum(ProgrammeFinances_TC$All_Schools), sum(ProgrammeFinances_TC$GeoSciences),sum(ProgrammeFinances_TC$SPSS),sum(ProgrammeFinances_TC$Law),sum(ProgrammeFinances_TC$Engineering),sum(ProgrammeFinances_TC$Business),sum(ProgrammeFinances_TC$Art))
	names(ProgrammeFinances_TC_Totals) <<- names(ProgrammeFinances_TC)
	row.names(ProgrammeFinances_TC_Totals) <<- "All_Programmes"
	ProgrammeFinances_TC <<- rbind(ProgrammeFinances_TC, ProgrammeFinances_TC_Totals)
		
}
		
#To Check inputs worked		
Programme_Level_Finances_Teaching()
ProgrammeData_TC
ProgrammeData_TC["EE"]
ProgrammeFinances_TC		
				   		
###PART 3: Calculations related to dissertation supervision
	
#Step 1: Calculate and apportion income for SRUC students within a particular programme associated with dissertation supervision. 
# The data files must include PT students for each year they are paying fees, as an equal % is taken each year

SRUC_Prog_DS <- function() {

	i = 1
	Lost_SRUC_DSStudent_Check <<- data.frame(Programme=character(), Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
	SRUC_Student_DS <<- vector('list', length(Programmes))
	SRUC_Student_DSFS <<- vector('list', length(Programmes))
	SRUC_Student_DSFSTI <<- vector('list', length(Programmes))
		
	while (i <= length(Programmes)) {
		## Imports xlsx file for each SRUC programme showing supervision details
		fn <- paste("Inputs/Dissertations/", Programmes[i], "_Dissertations", yr, ".xlsx", sep="")
		## Creates dataframe associated with course that holds position i in Courses
		SRUC_Student_DS[[i]]<<- read.xlsx(fn, header=TRUE, sheetIndex=1,as.data.frame=TRUE)
		
		# Remove trailing numbers from UUNs
		SRUC_Student_DS[[i]]$UUN <<- as.character(SRUC_Student_DS[[i]]$UUN)
		SRUC_Student_DS[[i]]$UUN <<- substr(SRUC_Student_DS[[i]]$UUN, 1, nchar(SRUC_Student_DS[[i]]$UUN)-2)
		SRUC_Student_DS[[i]]$UUN <<- as.character(SRUC_Student_DS[[i]]$UUN)

		# Need to document the number of MSc students within the course before merging, in case the FeeStatus data
			#is incomplete
			Pre_Merge_Length <- length(SRUC_Student_DS[[i]]$UUN)
			Pre_Merge_UUN <- as.vector(SRUC_Student_DS[[i]]$UUN)
		
		## Merges attendance list with fee status information
		SRUC_Student_DSFS[[i]] <<-merge(SRUC_Student_DS[[i]], FeeStatus[ , c("UUN", "FSG")], by=c("UUN"))
		# Need to now check if any students no longer appear in the data frame. If they don't appear it is because
			# they weren't in the Fee Status sheet...and the most likely explanation for that is they are pursuing their
			# studies for longer than initially anticipated (e.g. have had an interruption or concession to change status)
			# If this happens, this should print a warning to prompt us to go back and find the missing student data and
			# add it into the FeeStatus sheet if appropriate (i.e. unless they have already paid all their tuition AND we
			# have already been paid for it, and it's just the delay in them actually participating in whichever course
			Post_Merge_Length <- length(SRUC_Student_DSFS[[i]]$UUN)
			Post_Merge_UUN <- as.vector(SRUC_Student_DSFS[[i]]$UUN)
			Diff <- Post_Merge_Length - Pre_Merge_Length
			
			if (Diff <0) {
				Highlights <- paste("Warning:", abs(Diff) , "Student(s) LOST during merge")
			}
			else {
				Highlights <- ""
			}
			
			# Code from: http://stackoverflow.com/questions/17598134/compare-two-lists-in-r
			# Look for Teemu Daniel Laajala
			Inboth <- Pre_Merge_UUN[Pre_Merge_UUN %in% Post_Merge_UUN] # in both, same as call: intersect(first, second)
			OnlyInPreMerge <- Pre_Merge_UUN[!Pre_Merge_UUN %in% Post_Merge_UUN] # only in 'first', same as: setdiff(first, second)
			OnlyInPostMerge <- Post_Merge_UUN[!Post_Merge_UUN %in% Pre_Merge_UUN] # only in 'second', same as: setdiff(second, first)
			
			#For reference on listing lost UUNs in final column: 
			# http://stackoverflow.com/questions/13973116/convert-r-vector-to-string-vector-of-1-element
			Lost_SRUC_DSStudent_Check[i,] <<- c(Programmes[i], Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
					   
		# Rename FSG column to be "Fee_Status"
		names(SRUC_Student_DSFS[[i]])[names(SRUC_Student_DSFS[[i]])=="FSG"] <<-"Fee_Status"
		
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		SRUC_Student_DSFS[[i]]$Fee_Status[grepl("RUK|SEU", SRUC_Student_DSFS[[i]]$Fee_Status, ignore.case=FALSE)] <<- "H"
		
		
		## Merges supervision list with fee information
		SRUC_Student_DSFSTI[[i]]<<-merge(SRUC_Student_DSFS[[i]], TuitionFees_stacked[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Re-orders supervision list with fee information so it's easier to read
		SRUC_Student_DSFSTI[[i]]<<-SRUC_Student_DSFSTI[[i]][c("UUN", "Surname", "Forename", "Programme", "Matriculation", "Enrollment", "School", "Supervisor", "Organisation", "Detail", "Fee_Status", "Tuition")]
		## Calculates portion of total fee associated with each student's supervision
		SRUC_Student_DSFSTI[[i]][,13]<<-(0.25 * SRUC_Student_DSFSTI[[i]][,12])
		## Names this column to highlight the fee portion due to each student on the programme for supervision
		names(SRUC_Student_DSFSTI[[i]])[names(SRUC_Student_DSFSTI[[i]])=="V13"]<<-"Supervision_Fee"
		
		## Advances to the next course and repeats above steps until the list of programmes is exhausted
		i = i+1
	}
	
	names(SRUC_Student_DSFSTI) <<- Programmes
	names(SRUC_Student_DS) <<- Programmes
	names(SRUC_Student_DSFS) <<- Programmes
								     
	
	# Calculates allocation of disertation supervision funds for SRUC students
	j = 1
	
	ProgrammeFinances_SRUCstudent_DS <<- data.frame(All=numeric(), GBP_to_SRUC=numeric(), GBP_to_GeoSciences=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		sruc <- subset(SRUC_Student_DSFSTI[[j]], Organisation == "SRUC")
		gs <- subset(SRUC_Student_DSFSTI[[j]], Organisation == "University")
					
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All <- sum(SRUC_Student_DSFSTI[[j]]$Supervision_Fee)
		Total_sruc <- sum(sruc$Supervision_Fee)
		Total_gs <- sum(gs$Supervision_Fee)
				
		ProgrammeFinances_SRUCstudent_DS[j,] <<- list(Total_All, Total_sruc, Total_gs)
		row.names(ProgrammeFinances_SRUCstudent_DS)[j] <<- Programmes[j]
		
		CourseDataFSTI[[i]]$Net_TopSlice <<- 0.80 * as.numeric(as.character(CourseDataFSTI[[i]]$Course_Fee))
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
	
	#Append summary row to table
	ProgrammeFinances_SRUCstudent_DS_Totals <<- data.frame(sum(ProgrammeFinances_SRUCstudent_DS$All), sum(ProgrammeFinances_SRUCstudent_DS$GBP_to_SRUC),sum(ProgrammeFinances_SRUCstudent_DS$GBP_to_GeoSciences))
	names(ProgrammeFinances_SRUCstudent_DS_Totals) <<- names(ProgrammeFinances_SRUCstudent_DS)
	row.names(ProgrammeFinances_SRUCstudent_DS_Totals) <<- "All_Programmes"
	ProgrammeFinances_SRUCstudent_DS <<- rbind(ProgrammeFinances_SRUCstudent_DS, ProgrammeFinances_SRUCstudent_DS_Totals)
	
	#output lost student check
	write.xlsx(Lost_SRUC_DSStudent_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep=""), sheetName="IntStudDS", append=TRUE)
}
	
#To Check inputs worked		
SRUC_Prog_DS()
SRUC_Student_DS[1]
SRUC_Student_DSFS[1]
SRUC_Student_DSFSTI[1]

SRUC_Student_DS
SRUC_Student_DSFS
SRUC_Student_DSFSTI

ProgrammeFinances_SRUCstudent_DS
ProgrammeFinances_SRUCstudent_DS["EE",]
	
#Step 2: Calculates supervision fees associated with SRUC staff supervising non-SRUC students
# The data files must ask over how many summers a PT student is completing their dissertation to ensure full supervision fee is obtained from PT students

SRUC_ExternalStudent_DS <- function() {
	
	Lost_nonSRUC_DSStudent_Check <<- data.frame(Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
	
	## Imports xlsx for showing external student supervision details
	fn <- paste("Inputs/Dissertations/", "SRUC_ExternalDissertations", yr, ".xlsx", sep="")
	SRUC_ExternalStudent_DSS <<- read.xlsx(fn, header=TRUE, sheetIndex=1, as.data.frame=TRUE)
	
		# Remove trailing numbers from UUNs
		SRUC_ExternalStudent_DSS$UUN <<- as.character(SRUC_ExternalStudent_DSS$UUN)
		SRUC_ExternalStudent_DSS$UUN <<- substr(SRUC_ExternalStudent_DSS$UUN, 1, nchar(SRUC_ExternalStudent_DSS$UUN)-2)
		SRUC_ExternalStudent_DSS$UUN <<- as.character(SRUC_ExternalStudent_DSS$UUN)

		
		# Need to document the number of MSc students within the course before merging, in case the FeeStatus data
			#is incomplete
			Pre_Merge_Length <- length(SRUC_ExternalStudent_DSS$UUN)
			Pre_Merge_UUN <- as.vector(SRUC_ExternalStudent_DSS$UUN)
		
		## Merges attendance list with fee status information
		SRUC_ExternalStudent_DSFS <<- merge(SRUC_ExternalStudent_DSS, FeeStatus[ , c("UUN", "FSG")], by=c("UUN"))
		
		# Rename FSG column to be "Fee_Status"
		names(SRUC_ExternalStudent_DSFS)[names(SRUC_ExternalStudent_DSFS)=="FSG"] <<-"Fee_Status"
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		SRUC_ExternalStudent_DSFS$Fee_Status[grepl("RUK|SEU", SRUC_ExternalStudent_DSFS$Fee_Status, ignore.case=FALSE)] <<- "H"
		
		
		# Need to now check if any students no longer appear in the data frame. If they don't appear it is because
			# they weren't in the Fee Status sheet...and the most likely explanation for that is they are pursuing their
			# studies for longer than initially anticipated (e.g. have had an interruption or concession to change status)
			# If this happens, this should print a warning to prompt us to go back and find the missing student data and
			# add it into the FeeStatus sheet if appropriate (i.e. unless they have already paid all their tuition AND we
			# have already been paid for it, and it's just the delay in them actually participating in whichever course
			Post_Merge_Length <- length(SRUC_ExternalStudent_DSFS$UUN)
			Post_Merge_UUN <- as.vector(SRUC_ExternalStudent_DSFS$UUN)
			Diff <- Post_Merge_Length - Pre_Merge_Length
			
			if (Diff <0) {
				Highlights <- paste("Warning:", abs(Diff) , "Student(s) LOST during merge")
			}
			else {
				Highlights <- ""
			}
			
			# Code from: http://stackoverflow.com/questions/17598134/compare-two-lists-in-r
			# Look for Teemu Daniel Laajala
			Inboth <- Pre_Merge_UUN[Pre_Merge_UUN %in% Post_Merge_UUN] # in both, same as call: intersect(first, second)
			OnlyInPreMerge <- Pre_Merge_UUN[!Pre_Merge_UUN %in% Post_Merge_UUN] # only in 'first', same as: setdiff(first, second)
			OnlyInPostMerge <- Post_Merge_UUN[!Post_Merge_UUN %in% Pre_Merge_UUN] # only in 'second', same as: setdiff(second, first)
			
			#For reference on listing lost UUNs in final column: 
			# http://stackoverflow.com/questions/13973116/convert-r-vector-to-string-vector-of-1-element
			Lost_nonSRUC_DSStudent_Check[1,] <<- c(Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
			write.xlsx(Lost_nonSRUC_DSStudent_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep=""), sheetName="ExtStudDS", append=TRUE)									
		
	## Merges supervision list with fee information
	SRUC_ExternalStudent_DSFSTI <<- merge(SRUC_ExternalStudent_DSFS, TuitionFees_stacked[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
	
	## Re-orders supervision list with fee information so it's easier to read
	SRUC_ExternalStudent_DSFSTI<<- SRUC_ExternalStudent_DSFSTI[c("UUN", "Surname", "Forename", "Programme", "Matriculation", "Enrollment", "School", "Supervisor", "Research_Group", "Fee_Status", "Tuition", "Num_Summers_Supervision")]
	SRUC_ExternalStudent_DSFSTI
	## Calculates portion of total fee associated with each student's supervision
	SRUC_ExternalStudent_DSFSTI[,13] <<- ifelse((grepl("3 Years", SRUC_ExternalStudent_DSFSTI$Programme, ignore.case=TRUE) & SRUC_ExternalStudent_DSFSTI$Num_Summers_Supervision==1), 
											0.25 * (3 / 1) * SRUC_ExternalStudent_DSFSTI$Tuition, 
											ifelse((grepl("3 Years", SRUC_ExternalStudent_DSFSTI$Programme, ignore.case=TRUE) & SRUC_ExternalStudent_DSFSTI$Num_Summers_Supervision==2), 
											0.25 * (3 / 2 ) * SRUC_ExternalStudent_DSFSTI$Tuition,
											ifelse((grepl("3 Years", SRUC_ExternalStudent_DSFSTI$Programme, ignore.case=TRUE) & SRUC_ExternalStudent_DSFSTI$Num_Summers_Supervision==3), 
											0.25 * (3 / 3 ) * SRUC_ExternalStudent_DSFSTI$Tuition,
											ifelse((grepl("2 Years", SRUC_ExternalStudent_DSFSTI$Programme, ignore.case=TRUE) & SRUC_ExternalStudent_DSFSTI$Num_Summers_Supervision==1), 
											0.25 * (2 / 1) * SRUC_ExternalStudent_DSFSTI$Tuition, 
											ifelse((grepl("2 Years", SRUC_ExternalStudent_DSFSTI$Programme, ignore.case=TRUE) & SRUC_ExternalStudent_DSFSTI$Num_Summers_Supervision==2), 
											0.25 * (2 / 2 ) * SRUC_ExternalStudent_DSFSTI$Tuition,
											0.25 * SRUC_ExternalStudent_DSFSTI$Tuition)))))
	
	
	## Names this column to highlight the fee portion due to each student on the programme for supervision
	names(SRUC_ExternalStudent_DSFSTI)[13]<<-"Supervision_Fee"
	SRUC_ExternalStudent_DSFSTI
	
	# Need to specify column and/or change it to character from factor to get to work...
		SRUC_ExternalStudent_DSFSTI$School <<- as.character(SRUC_ExternalStudent_DSFSTI$School)
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Geosciences", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "GeoSciences"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Social And Political Science", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "SPSS"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Engineering", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "Engineering"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Law", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "Law"
		SRUC_ExternalStudent_DSFSTI$School[grepl("Business School", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "Business"
		SRUC_ExternalStudent_DSFSTI$School[grepl("Edinburgh College Of Art", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <<- "Art"	

#Step 3: Groups supervision fees associated with SRUC staff supervising non-SRUC students by Research Group

	i = 1
	RGData_ExtDS <<- vector('list', length(Research_Groups))
	
	while (i <= length(Research_Groups)) {
		## Separates out the subsets associated with each Research Group
		RGData_ExtDS[[i]] <<- subset(SRUC_ExternalStudent_DSFSTI, Research_Group == as.character(Research_Groups[[i]]))
		
		## Advances to the next Research group and repeats above steps until the list of is exhausted
		i = i+1
	}
	
	RGData_ExtDS 
	names(RGData_ExtDS) <<- Research_Groups

	# Creates summary financial picture by research group 
	j = 1
	RGFinances_ExtDS <<- data.frame(All=numeric(), GeoSciences=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), Art=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(RGData_ExtDS)) {
		#Step 1: Pull out the subsets of students from each school being supervised by SRUC staff in each research group
		gs <- subset(RGData_ExtDS[[j]], School == "GeoSciences")
		spss <- subset(RGData_ExtDS[[j]], School == "SPSS")
		law <- subset(RGData_ExtDS[[j]], School == "Law")			
		eng <- subset(RGData_ExtDS[[j]], School == "Engineering")
		bus <- subset(RGData_ExtDS[[j]], School == "Business")
		art <- subset(RGData_ExtDS[[j]], School == "Art")
			
		#Step 2: Determine the total tuition associated with supervising by student school 
		Total_All <- sum(RGData_ExtDS[[j]]$Supervision_Fee)
		Total_gs <- sum(gs$Supervision_Fee)
		Total_spss <- sum(spss$Supervision_Fee)
		Total_law <- sum(law$Supervision_Fee)
		Total_eng <- sum(eng$Supervision_Fee)
		Total_bus <- sum(bus$Supervision_Fee)
		Total_art <- sum(art$Supervision_Fee)
		
		RGFinances_ExtDS[j,] <<- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus, Total_art)
		row.names(RGFinances_ExtDS)[j] <<- Research_Groups[j]
		
		## Advances to the next research group and repeats above steps until the list of research groups is exhausted
		j = j+1
	}	
	
	#Append summary row to table
	RGFinances_ExtDS_Totals <<- data.frame(sum(RGFinances_ExtDS$All), sum(RGFinances_ExtDS$GeoSciences), sum(RGFinances_ExtDS$SPSS), sum(RGFinances_ExtDS$Law), sum(RGFinances_ExtDS$Engineering), sum(RGFinances_ExtDS$Business), sum(RGFinances_ExtDS$Art))
	names(RGFinances_ExtDS_Totals) <<- names(RGFinances_ExtDS)
	row.names(RGFinances_ExtDS_Totals) <<- "All_ResearchGroups"
	RGFinances_ExtDS <<- rbind(RGFinances_ExtDS, RGFinances_ExtDS_Totals)
	
}

#To Check inputs worked
SRUC_ExternalStudent_DS()

SRUC_ExternalStudent_DSS
SRUC_ExternalStudent_DSFS
SRUC_ExternalStudent_DSFSTI

RGData_ExtDS
RGFinances_ExtDS
RGFinances_ExtDS["LEES",]

 	
###PART 4: Calculations related to administration

# In order to avoid charging the administration fee for PT students more than once, and to simplify calculations, charge the admin fee for them only in the year when they 
# work on their dissertations (rather than annually based on a fee fraction)

SRUC_Admin <- function() {

#Calculate 20% top-slice for university

##Taught component top-slice
TC_Top_Slice <<- 0.20 * ProgrammeFinances_TC
TC_SRUC_Share <<- 0.80 * ProgrammeFinances_TC

##Internal (own programme) dissertation supervision component top-slice
InternalDS_Top_Slice <<- 0.20 * ProgrammeFinances_SRUCstudent_DS
InternalDS_SRUC_Share <<- 0.80 * ProgrammeFinances_SRUCstudent_DS

##External (non-SRUC) dissertation supervision component top-slice
ExternalDS_Top_Slice <<- 0.20 * RGFinances_ExtDS 
ExternalDS_SRUC_Share <<- 0.80 * RGFinances_ExtDS

#Calculate Division between SRUC and GeoSciences for rest

##Taught component admin division
TC_GS_Admin <<- 0.15 * TC_SRUC_Share
SRUC_FinalTC_Share <<- 0.85 * TC_SRUC_Share

##Internal (own programme) dissertation supervsion admin division
IntDS_GS_Admin <<- 0.15 * InternalDS_SRUC_Share
SRUC_FinalIntDS_Share <<- 0.85 * InternalDS_SRUC_Share
SRUC_FinalIntDS_Share <<- SRUC_FinalIntDS_Share[2]

##External (non-SRUC) dissertation supervision admin division
ExtDS_GS_Admin <<- 0.15 * ExternalDS_SRUC_Share
SRUC_FinalExtDS_Share <<- 0.85 * ExternalDS_SRUC_Share

#Key Totals

##Total Top-Slice
Total_Top_Slice <<- TC_Top_Slice[6,1] + InternalDS_Top_Slice[6,1] + ExternalDS_Top_Slice[3,1]

##Total SRUC Fee Allocation
Total_SRUC_Income <<- SRUC_FinalTC_Share[6,1] + SRUC_FinalIntDS_Share[6,1] + SRUC_FinalExtDS_Share[3,1]
Total_SRUC_PGOffice <<- 0.10 * Total_SRUC_Income

##PGR Office Income Sources
Gross_TC_PGO_Allocation <<- 0.10 * SRUC_FinalTC_Share
Gross_IntDS_PGO_Allocation <<- 0.10 * SRUC_FinalIntDS_Share
Gross_ExtDS_PGO_Allocation <<- 0.10 * SRUC_FinalExtDS_Share

##Total Programme & Research Group Income
Gross_TC_Allocation <<- 0.90 * SRUC_FinalTC_Share
Gross_IntDS_Allocation <<- 0.90 * SRUC_FinalIntDS_Share
Gross_ExtDS_Allocation <<- 0.90 * SRUC_FinalExtDS_Share

Gross_TC_IntDS_Per_Programme <<- data.frame(matrix(0, nrow=6, ncol = 0)) 
Gross_TC_IntDS_Per_Programme$Taught_Component <<- Gross_TC_Allocation[,1]
Gross_TC_IntDS_Per_Programme$Int_DS <<- Gross_IntDS_Allocation[,1]
Gross_TC_IntDS_Per_Programme$Total <<- Gross_TC_Allocation[,1] + Gross_IntDS_Allocation[,1]
row.names(Gross_TC_IntDS_Per_Programme) <<- c("EE", "EPM", "FS", "SS", "SPH", "All_Programmes")

##Total GeoSciences Admin Allocation
Total_GS_Admin <<- TC_GS_Admin[6,1] + IntDS_GS_Admin[6,1] + ExtDS_GS_Admin[3,1]

##Summary Table
GRAND_TOTALS <<- data.frame(Total_Top_Slice, Total_SRUC_Income, Total_GS_Admin, (0.85 * InternalDS_SRUC_Share[6,3]), (Total_GS_Admin + (0.85 * InternalDS_SRUC_Share[6,3])))
colnames(GRAND_TOTALS)[4] <<- "GBP_to_GS_for_DS"
colnames(GRAND_TOTALS)[5] <<- "GeoSciences_Total"

}

SRUC_Admin()

#To Check inputs worked
GRAND_TOTALS

Total_SRUC_PGOffice
Gross_TC_PGO_Allocation 
Gross_IntDS_PGO_Allocation 
Gross_ExtDS_PGO_Allocation

Gross_TC_Allocation 
Gross_IntDS_Allocation 
Gross_ExtDS_Allocation 
Gross_TC_IntDS_Per_Programme

###PART 5: Export all relevant results to Excel workbook

#This Excel file should be a workbook with the following features:
# 1. The first page should show the total owed to SRUC across all programmes, showing various splits
# 2. The total owed to each programme showing the same split as above, and for each course 'owned' by that programme
# Subsequent worksheets should show the individual course pages for all SRUC courses

	## Step 1: Prep the worksheets that are desired in the Excel document
		#Grand Totals
		GRAND_TOTALS
		
		#Top-Slice
		TC_Top_Slice
		InternalDS_Top_Slice
		ExternalDS_Top_Slice
		
		#GeoSciences Administration Fee Allocation
		TC_GS_Admin
		IntDS_GS_Admin
		ExtDS_GS_Admin
		
		#SRUC Totals after admin costs
		SRUC_FinalTC_Share
		SRUC_FinalIntDS_Share
		SRUC_FinalExtDS_Share
		
		#SRUC PGO Summary Information (PGO allocation before time costs put against it)
		Total_SRUC_PGOffice
		Gross_TC_PGO_Allocation 
		Gross_IntDS_PGO_Allocation 
		Gross_ExtDS_PGO_Allocation 
		
		#SRUC Programmes & Research Groups (Income before time costs put against it)
		Gross_TC_Allocation
		Gross_IntDS_Allocation
		Gross_ExtDS_Allocation
		
		#Ecological Economics Worksheets
		EE_Diss <- SRUC_Student_DSFSTI[1]
		EE_Courses <- ProgrammeData_TC[1]
		FEE_Summary <- CourseDataFSTI[1] 
		EV_Summary <- CourseDataFSTI[2]
		AEE_Summary <- CourseDataFSTI[3]
		PPP_Summary <- CourseDataFSTI[4]
		EIA_Summary <- CourseDataFSTI[5]
		
		#EPM Worksheets 
		EPM_Diss <- SRUC_Student_DSFSTI[2]
		EPM_Courses <- ProgrammeData_TC[2]
		AQCG_Summary <- CourseDataFSTI[6]
		LUEI_Summary <- CourseDataFSTI[7]						    
		WRM_Summary <- CourseDataFSTI[8]						    
		EVM_Summary <- CourseDataFSTI[9]
		AEST_Summary <- CourseDataFSTI[10]						    
								    
		#FS Worksheets
		FS_Diss <- SRUC_Student_DSFSTI[3]
		FS_Courses <- ProgrammeData_TC[3]
		FAFS_Summary <- CourseDataFSTI[11]
		IFS_Summary <- CourseDataFSTI[12]						    
		SFP_Summary <- CourseDataFSTI[13]						    
								    
		#SS Worksheets 
		SS_Diss <- SRUC_Student_DSFSTI[4]
		SS_Courses <- ProgrammeData_TC[4]
		SPM_Summary <- CourseDataFSTI[14]
		SET_Summary <- CourseDataFSTI[15]						    
		SSCA_Summary <- CourseDataFSTI[16]						    
								    
		#SPH Worksheets
		SPH_Diss <- SRUC_Student_DSFSTI[5]
		SPH_Courses <- ProgrammeData_TC[5]
		FPH_Summary <- CourseDataFSTI[17]
		FOPH_Summary <- CourseDataFSTI[18]
		PHGC_Summary <- CourseDataFSTI[19]						    

#Basing it on approach shown here: https://statmethods.wordpress.com/2014/06/19/quickly-export-multiple-r-objects-to-an-excel-workbook/

SRUC.PGT.AnnualFinancialSummary <- function (file, GRAND_TOTALS, 
									TC_Top_Slice, InternalDS_Top_Slice, ExternalDS_Top_Slice,
									TC_GS_Admin, IntDS_GS_Admin, ExtDS_GS_Admin,
									SRUC_FinalTC_Share, SRUC_FinalIntDS_Share, SRUC_FinalExtDS_Share,
									Total_SRUC_PGOffice, Gross_TC_PGO_Allocation, Gross_IntDS_PGO_Allocation, Gross_ExtDS_PGO_Allocation,
									Gross_TC_Allocation, Gross_IntDS_Allocation, Gross_ExtDS_Allocation,
									EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
					    			EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					    			FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					    			SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					    			SPH_Diss, SPH_Courses, FPH_Summary, FOPH_Summary, PHGC_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(GRAND_TOTALS, 
					TC_Top_Slice, InternalDS_Top_Slice, ExternalDS_Top_Slice,
					TC_GS_Admin, IntDS_GS_Admin, ExtDS_GS_Admin,
					SRUC_FinalTC_Share, SRUC_FinalIntDS_Share, SRUC_FinalExtDS_Share, 
					Total_SRUC_PGOffice, Gross_TC_PGO_Allocation, Gross_IntDS_PGO_Allocation, Gross_ExtDS_PGO_Allocation,
					Gross_TC_Allocation, Gross_IntDS_Allocation, Gross_ExtDS_Allocation,
					EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
					EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					SPH_Diss, SPH_Courses, FPH_Summary, FOPH_Summary, PHGC_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

#To generate the Excel workbook, run this code
SRUC.PGT.AnnualFinancialSummary(paste("Outputs/SRUC_PGT_FinancialSummary_", yr, ".xlsx", sep=""), 
									GRAND_TOTALS, 
									TC_Top_Slice, InternalDS_Top_Slice, ExternalDS_Top_Slice,
									TC_GS_Admin, IntDS_GS_Admin, ExtDS_GS_Admin,
									SRUC_FinalTC_Share, SRUC_FinalIntDS_Share, SRUC_FinalExtDS_Share,
									Total_SRUC_PGOffice, Gross_TC_PGO_Allocation, Gross_IntDS_PGO_Allocation, Gross_ExtDS_PGO_Allocation,
									Gross_TC_Allocation, Gross_IntDS_Allocation, Gross_ExtDS_Allocation,
									EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
					    			EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					    			FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					    			SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					    			SPH_Diss, SPH_Courses, FPH_Summary, FOPH_Summary, PHGC_Summary)

#Generates excel file with just EE information
EE.PGT.AnnualFinancialSummary <- function (file, EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

EE.PGT.AnnualFinancialSummary(paste("Outputs/EE_PGT_FinancialSummary_", yr, ".xlsx", sep=""), EE_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)

#Generates excel file with just EPM information
EPM.PGT.AnnualFinancialSummary <- function (file, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

EPM.PGT.AnnualFinancialSummary(paste("Outputs/EPM_PGT_FinancialSummary_", yr, ".xlsx", sep=""), EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary)

#Generates excel file with just FS information
FS.PGT.AnnualFinancialSummary <- function (file, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

FS.PGT.AnnualFinancialSummary(paste("Outputs/FS_PGT_FinancialSummary_", yr, ".xlsx", sep=""), FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary)

#Generates excel file with just SS information 
SS.PGT.AnnualFinancialSummary <- function (file, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

FS.PGT.AnnualFinancialSummary(paste("Outputs/SS_PGT_FinancialSummary_", yr, ".xlsx", sep=""), SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary)

#Generates excel file with just SPH information 
SPH.PGT.AnnualFinancialSummary <- function (file, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SPH_Diss, SPH_Courses, FPH_Summary, FOPH_Summary, PHGC_Summary)
	fargs <- as.list(match.call(expand.dots = TRUE))
	objnames <- as.character(fargs)[-c(1,2)]
	nobjects <- length(objects)
	for (i in 1:nobjects) {
		if (i ==1) {
			write.xlsx(objects[[i]], file, sheetName = objnames[i])
		}
		else {
			write.xlsx(objects[[i]], file, sheetName = objnames[i], append = TRUE)
		}
	}
	print(paste("Workbook", file, "has", nobjects, "worksheets."))
}

SPH.PGT.AnnualFinancialSummary(paste("Outputs/SPH_PGT_FinancialSummary_", yr, ".xlsx", sep=""), SPH_Diss, SPH_Courses, FPH_Summary, FOPH_Summary, PHGC_Summary)




















