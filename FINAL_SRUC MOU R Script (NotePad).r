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
		# remove any rows where there is a PhD student enrolled, as they should not be enrolled
		CourseData[[i]] <<- CourseData[[i]][!grepl("PhD", CourseData[[i]]$Programme),]
		
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
		## Builds the Course_Fee column based on whether a student is one of the 2 part time options
		##  or everyone else (i.e. all the ft options). The % charged are increased to ensure the absolute 
		## quantity paid for each course is the same as a ft student
		CourseDataFSTI[[i]]$Course_Fee <<- ifelse(grepl("3 Years", CourseDataFSTI[[i]]$Programme, ignore.case=TRUE), 
							0.15 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting)), 
						ifelse(grepl("2 Years", CourseDataFSTI[[i]]$Programme, ignore.case=TRUE), 
						       0.10 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting)), 
						       0.05 * CourseDataFSTI[[i]]$Tuition * as.numeric(as.character(CourseDataFSTI[[i]]$Credit_Weighting))))
		
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
	
	CourseFinances <<- data.frame(All=numeric(), GeoSciences=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
	while (i <= length(Courses)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		gs <- subset(CourseDataFSTI[[i]], School == "GeoSciences")
		spss <- subset(CourseDataFSTI[[i]], School == "SPSS")
		law <- subset(CourseDataFSTI[[i]], School == "Law")			
		eng <- subset(CourseDataFSTI[[i]], School == "Engineering")
		bus <- subset(CourseDataFSTI[[i]], School == "Business")
			
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All <- sum(CourseDataFSTI[[i]]$Course_Fee)
		Total_gs <- sum(gs$Course_Fee)
		Total_spss <- sum(spss$Course_Fee)
		Total_law <- sum(law$Course_Fee)
		Total_eng <- sum(eng$Course_Fee)
		Total_bus <- sum(bus$Course_Fee)
		
		CourseFinances[i,] <<- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus)
			
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	CourseFinances <<- data.frame( Courses, Programme_Ownership, CourseFinances)
}

#To Check inputs worked
Course_Level_Finances()
CourseFinances["FEE",]
	
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
	ProgrammeFinances_TC <<- data.frame(All=numeric(), GeoSciences=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		#Within the dataframe showing courses owned by programmes, sum the relevant tuition fee components by column
		Total_All <- sum(ProgrammeData_TC[[j]]$All)
		Total_gs <- sum(ProgrammeData_TC[[j]]$GeoSciences)
		Total_spss <- sum(ProgrammeData_TC[[j]]$SPSS)
		Total_law <- sum(ProgrammeData_TC[[j]]$Law)
		Total_eng <- sum(ProgrammeData_TC[[j]]$Engineering)
		Total_bus <- sum(ProgrammeData_TC[[j]]$Business)
		
		ProgrammeFinances_TC[j,] <<- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus)
		row.names(ProgrammeFinances_TC)[j] <<- Programmes[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
}
		
#To Check inputs worked		
Programme_Level_Finances_Teaching()
ProgrammeData_TC
ProgrammeFinances_TC		
				   		
###PART 3: Calculations related to dissertation supervision
	
#Step 1: Calculate and apportion income for SRUC students associated with dissertation supervision. 
# The data files must include PT students for each year they are paying fees, as an equal % is taken each year

SRUC_Prog_DS <- function() {

	i = 1
	Lost_Student_Check <- data.frame(Programme=character(), Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
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
			Lost_Student_Check[i,] <- c(Programmes[i], Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
					   
		# Rename FSG column to be "Fee_Status"
		names(SRUC_Student_DSFS[[i]])[names(SRUC_Student_DSFS[[i]])=="FSG"] <<-"Fee_Status"
		
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		SRUC_Student_DSFS[[i]]$Fee_Status[grepl("RUK|SEU", SRUC_Student_DSFS[[i]]$Fee_Status, ignore.case=FALSE)] <<- "H"
		
		
		## Merges supervision list with fee information
		SRUC_Student_DSFSTI[[i]]<<-merge(SRUC_Student_DSFS[[i]], TuitionFees_stacked[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Re-orders supervision list with fee information so it's easier to read
		SRUC_Student_DSFSTI[[i]]<<-SRUC_Student_DSFSTI[[i]][c("UUN", "Surname", "Forename", "Programme", "Matriculation", "Enrollment", "School", "Supervisor", "Organisation", "Detail", "Fee_Status", "Tuition")]
		## Calculates portion of total fee associated with each student's supervision
		SRUC_Student_DSFSTI[[i]][,13]<<-(0.10 * SRUC_Student_DSFSTI[[i]][,12])
		## Names this column to highlight the fee portion due to each student on the programme for supervision
		names(SRUC_Student_DSFSTI[[i]])[names(SRUC_Student_DSFSTI[[i]])=="V13"]<<-"Supervision_Fee"
		
		## Advances to the next course and repeats above steps until the list of programmes is exhausted
		i = i+1
	}
	
	names(SRUC_Student_DSFSTI) <<- Programmes
								     
	
	# Calculates allocation of disertation supervision funds for SRUC students
	j = 1
	
	ProgrammeFinances_SRUCstudent_DS <<- data.frame(All=numeric(), SRUC=numeric(), GeoSciences=numeric(), stringsAsFactors=FALSE)
	
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
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
	
	write.xlsx(Lost_Student_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep=""), sheetName="IntStudDS", append=TRUE)
}
	
#To Check inputs worked		
SRUC_Prog_DS()
SRUC_Student_DS
SRUC_Student_DSFS
SRUC_Student_DSFSTI


ProgrammeFinances_SRUCstudent_DS
ProgrammeFinances_SRUCstudent_DS["EE",]
	
#Step 2: Calculates supervision fees associated with SRUC staff supervising non-SRUC students
SRUC_ExternalStudent_DS <- function() {
	
	Lost_Student_Check <- data.frame(Research_Group=character(), Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
	
	## Imports xlsx for showing external student supervision details
	fn <- paste("Inputs/Dissertations/", "SRUC_ExternalDissertations", yr, ".xlsx", sep="")
	SRUC_ExternalStudent_DSS <<- read.xlsx(fn, header=TRUE, sheetIndex=1, as.data.frame=TRUE)
	
		# Remove trailing numbers from UUNs
		SRUC_ExternalStudent_DSS$UUN <- as.character(SRUC_ExternalStudent_DSS$UUN)
		SRUC_ExternalStudent_DSS$UUN <- substr(SRUC_ExternalStudent_DSS$UUN, 1, nchar(SRUC_ExternalStudent_DSS$UUN)-2)
		SRUC_ExternalStudent_DSS$UUN <- as.character(SRUC_ExternalStudent_DSS$UUN)

		
		# Need to document the number of MSc students within the course before merging, in case the FeeStatus data
			#is incomplete
			Pre_Merge_Length <- length(SRUC_ExternalStudent_DSS$UUN)
			Pre_Merge_UUN <- as.vector(SRUC_ExternalStudent_DSS$UUN)
		
		## Merges attendance list with fee status information
		SRUC_ExternalStudent_DSFS <<-merge(SRUC_ExternalStudent_DSS, FeeStatus[ , c("UUN", "FSG")], by=c("UUN"))
		
		# Rename FSG column to be "Fee_Status"
		names(SRUC_ExternalStudent_DSFS)[names(SRUC_ExternalStudent_DSFS)=="FSG"] <-"Fee_Status"
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		SRUC_ExternalStudent_DSFS$Fee_Status[grepl("RUK|SEU", SRUC_ExternalStudent_DSFS$Fee_Status, ignore.case=FALSE)] <- "H"
		
		
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
			Lost_Student_Check[i,] <- c(Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
					   
		
	## Merges supervision list with fee information
	SRUC_ExternalStudent_DSFSTI <<-merge(SRUC_ExternalStudent_DSFS, TuitionFees_stacked[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
	
	## Re-orders supervision list with fee information so it's easier to read
	SRUC_ExternalStudent_DSFSTI<<- SRUC_ExternalStudent_DSFSTI[c("UUN", "Surname", "Forename", "Programme", "Matriculation", "Enrollment", "School", "Supervisor", "Research_Group", "Fee_Status", "Tuition")]
	SRUC_ExternalStudent_DSFSTI
	## Calculates portion of total fee associated with each student's supervision
	SRUC_ExternalStudent_DSFSTI[12] <-(0.10 * SRUC_ExternalStudent_DSFSTI[11])
	SRUC_ExternalStudent_DSFSTI
	## Names this column to highlight the fee portion due to each student on the programme for supervision
	names(SRUC_ExternalStudent_DSFSTI)[12]<-"Supervision_Fee"
	SRUC_ExternalStudent_DSFSTI

	# Need to specify column and/or change it to character from factor to get to work...
		SRUC_ExternalStudent_DSFSTI$School <- as.character(SRUC_ExternalStudent_DSFSTI$School)
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Geosciences", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <- "GeoSciences"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Social And Political Science", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <- "SPSS"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Engineering", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <- "Engineering"
		SRUC_ExternalStudent_DSFSTI$School[grepl("School Of Law", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <- "Law"
		SRUC_ExternalStudent_DSFSTI$School[grepl("Business School", SRUC_ExternalStudent_DSFSTI$School, ignore.case=FALSE)] <- "Business"		


	#Provides data on students by supervisor research group
	i = 1
	RGData <- vector('list', length(Research_Groups))
	
	while (i <= length(Research_Groups)) {
		## Separates out the subsets associated with each Research Group
		RGData[[i]] <- subset(SRUC_ExternalStudent_DSFSTI, Research_Group == as.character(Research_Groups[[i]]))
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	
	RGData 
	names(RGData) <- Research_Groups



	write.xlsx(Lost_Student_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep="", ), sheetName="ExtStudDS", append=TRUE)
				
	# Creates summary financial picture by research group 
	j = 1
	RGFinances <- data.frame(All=numeric(), GS=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(RGData)) {
		#Step 1: Pull out the subsets of students from each school being supervised by SRUC staff in each research group
		gs <- subset(RGData[[j]], School == "GeoSciences")
		spss <- subset(RGData[[j]], School == "SPSS")
		law <- subset(RGData[[j]], School == "Law")			
		eng <- subset(RGData[[j]], School == "Engineering")
		bus <- subset(RGData[[j]], School == "Business")
			
		#Step 2: Determine the total tuition associated with supervising by student school 
		Total_All <- sum(RGData[[j]]$Supervision_Fee)
		Total_gs <- sum(gs$Supervision_Fee)
		Total_spss <- sum(spss$Supervision_Fee)
		Total_law <- sum(law$Supervision_Fee)
		Total_eng <- sum(eng$Supervision_Fee)
		Total_bus <- sum(bus$Supervision_Fee)
		
		RGFinances[j,] <- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus)
		row.names(RGFinances)[j] <- Research_Groups[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}	
}

#To Check inputs worked
SRUC_ExternalStudent_DS()
SRUC_ExternalStudent_DSS
SRUC_ExternalStudent_DSFS
SRUC_ExternalStudent_DSFSTI
RGData
RGFinances
RGFinances["LEES",]

 	
###PART 4: Calculations related to administration

# In order to avoid charging the administration fee for PT students more than once, and to simplify calculations, charge the admin fee for them only in the year when they 
# work on their dissertations (rather than annually based on a fee fraction)

SRUC_Admin <- function() {

	i = 1
	Lost_Student_Check <- data.frame(Programme, Pre_Merge=numeric(), Post_Merge=numeric(), Difference=numeric(), Highlights=character(), Lost_UUNs=character(), stringsAsFactors=FALSE)
	SRUC_AdminData <<- vector('list', length(Programmes))
		
	while (i <= length(Programmes)) {
		## Imports csv for each SRUC programme showing supervision details
		fn <- paste("Inputs/",Programmes[i], "_Dissertations", yr, ".xlsx", sep="")
		## Creates dataframe associated with course that holds position i in Courses
		SRUC_AdminData[[i]]<<- read.xlsx(fn, header=TRUE, as.data.frame=TRUE)
		
		#Trim trailing whitespace in case this appears
			## Source of this approach is: http://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace-in-r
			### Look for sub-comment by Thieme Hennis Sep 19 '14 
			SRUC_AdminData[[i]] <<- as.data.frame(apply(SRUC_AdminData[[i]],2,function (x) sub("\\s+$", "", x)))
		
		
		# Need to document the number of MSc students within the course before merging, in case the FeeStatus data
			#is incomplete
			Pre_Merge_Length <- length(SRUC_AdminData[[i]]$UUN)
			Pre_Merge_UUN <- as.vector(SRUC_AdminData[[i]]$UUN)
		## Merges attendance list with fee status information
		SRUC_AdminData[[i]] <<-merge(SRUC_AdminData[[i]], FeeStatus[ , c("UUN", "FSG")], by=c("UUN"))
		# Need to now check if any students no longer appear in the data frame. If they don't appear it is because
			# they weren't in the Fee Status sheet...and the most likely explanation for that is they are pursuing their
			# studies for longer than initially anticipated (e.g. have had an interruption or concession to change status)
			# If this happens, this should print a warning to prompt us to go back and find the missing student data and
			# add it into the FeeStatus sheet if appropriate (i.e. unless they have already paid all their tuition AND we
			# have already been paid for it, and it's just the delay in them actually participating in whichever course
			Post_Merge_Length <- length(SRUC_AdminData[[i]]$UUN)
			Post_Merge_UUN <- as.vector(SRUC_AdminData[[i]]$UUN)
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
			Lost_Student_Check[i,] <- c(Programmes[i], Pre_Merge_Length, Post_Merge_Length, abs(Diff), Highlights, paste(OnlyInPreMerge, collapse=", "))			   
					   
		# Rename FSG column to be "Fee_Status"
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="FSG"] <<-"Fee_Status"
		# Change any entry with RUK or SEU as the Fee Status Group to H (thus everything is O or H) 
		SRUC_AdminData[[i]]$Fee_Status[grepl("RUK|SEU", SRUC_AdminData[[i]]$Fee_Status, ignore.case=FALSE)] <<- "H"
		
		## Merges supervision list with fee information
		SRUC_AdminData[[i]]<<-merge(SRUC_AdminData[[i]], TuitionFees[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Re-orders supervision list with fee information so it's easier to read
		SRUC_AdminData[[i]]<<-SRUC_AdminData[[i]][c("UUN", "Surname", "Forename", "Programme", "Matriculation", "Enrollment", "School", "Supervisor", "Organisation", "Detail", "Fee_Status", "Tuition")]
		## Calculates portion of total fee associated with admin for each student
		SRUC_AdminData[[i]][,13]<<-(0.40 * SRUC_AdminData[[i]][,12])
		## Names this column to highlight the fee portion due for admin
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V13"]<<-"Admin_Fee_Total"
		## Calculates portion of total admin fee that belongs to SRUC and GeoSciences
		SRUC_AdminData[[i]][,14]<<-(0.75 * SRUC_AdminData[[i]][,13])
		SRUC_AdminData[[i]][,15]<<-(0.25 * SRUC_AdminData[[i]][,13])
		## Re-Names these columns 
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V14"]<<-"Admin_Fee_SRUC"
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V15"]<<-"Admin_Fee_GeoSciences"
		
		## Advances to the next course and repeats above steps until the list of programmes is exhausted
		i = i+1
	}
							
	names(SRUC_AdminData) <<- Programmes
	write.xlsx(Lost_Student_Check, paste("Outputs/Tests/LostStudentCheck_", yr, ".xlsx", sep="", ), sheetName="Admin", append=TRUE)							    

	# Colates summary information on admin fee by programme and SRUC vs. GeoSciences
	j = 1
	
	ProgrammeFinances_Admin <<- data.frame(All=numeric(), SRUC=numeric(), GeoSciences=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		# #Step 1: Define subsets of dataframes to group students from different schools on each course
		# sruc <- subset(SRUC_AdminData[[j]], Organisation == "SRUC")
		# gs <- subset(SRUC_AdminData[[j]], Organisation == "University")
					
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All <- sum(SRUC_AdminData[[j]]$Admin_Fee_Total)
		Total_sruc <- sum(SRUC_AdminData[[j]]$Admin_Fee_SRUC)
		Total_gs <- sum(SRUC_AdminData[[j]]$Admin_Fee_GeoSciences)
				
		ProgrammeFinances_Admin[j,] <<- list(Total_All, Total_sruc, Total_gs)
		row.names(ProgrammeFinances_Admin)[j] <<- Programmes[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}	
}
	
#To Check inputs worked		
SRUC_Admin()
SRUC_AdminData
ProgrammeFinances_Admin
ProgrammeFinances_Admin["EE",]

###PART 5: Institution-Level summary calculations

Institutional_Summary <- function() {

	i = 1
		
	#Shows the amount of money due to SRUC *from* each of the 5 schools we engage with for PGT delivery
	SRUC_InstitutionalSummary <<- data.frame(Total=numeric(),Subtotal_GeoSciences=numeric(), GeoSciences_Teaching=numeric(), GeoSciences_Dissertations=numeric(), GeoScienes_Administration=numeric(),
																Subtotal_SPSS=numeric(),SPSS_Teaching=numeric(), SPSS_Dissertations=numeric(),
																Subtotal_Law=numeric(), Law_Teaching=numeric(), Law_Dissertations=numeric(),
																Subtotal_Engineering=numeric(),Engineering_Teaching=numeric(), Engineering_Dissertations=numeric(),
																Subtotal_Business=numeric(), Business_Teaching=numeric(), Business_Dissertations=numeric(),
																stringsAsFactors=FALSE)
	
	#Populate GeoSciences portion of summary table
	while (i <=length(Programmes)) {
							
		## Step 1: Determine total moneys owed from GeoSciences (vs others) to each programme across all elements of the fee
		Total_teaching_gs <- ProgrammeFinances_TC[[i,"GeoSciences"]]
		Total_diss_gs <- ProgrammeFinances_SRUCstudent_DS[[i,"SRUC"]]
		Total_admin_gs <- ProgrammeFinances_Admin[[i, "SRUC"]]
		Total_gs <- Total_teaching_gs + Total_diss_gs + Total_admin_gs

		Total_teaching_spss <- ProgrammeFinances_TC[[i,"SPSS"]]
		Total_diss_spss <- 0
		Total_spss <- Total_teaching_spss + Total_diss_spss 

		Total_teaching_law <- ProgrammeFinances_TC[[i,"Law"]]
		Total_diss_law <- 0
		Total_law <- Total_teaching_law + Total_diss_law 
		
		Total_teaching_eng <- ProgrammeFinances_TC[[i,"Engineering"]]
		Total_diss_eng <- 0
		Total_eng <- Total_teaching_eng + Total_diss_eng 
		
		Total_teaching_bus <- ProgrammeFinances_TC[[i,"Business"]]
		Total_diss_bus <- 0
		Total_bus <- Total_teaching_bus + Total_diss_bus
		
		Total_All <- Total_gs + Total_spss + Total_law + Total_eng + Total_bus
				
		SRUC_InstitutionalSummary[i,] <<- list(Total_All, Total_gs, Total_teaching_gs, Total_diss_gs, Total_admin_gs, 
															Total_spss, Total_teaching_spss, Total_diss_spss,
															Total_law, Total_teaching_law, Total_diss_law,
															Total_eng, Total_teaching_eng, Total_diss_eng,
															Total_bus, Total_teaching_bus, Total_diss_bus)
															
		row.names(SRUC_InstitutionalSummary)[i] <<- Programmes[i]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
		##Step 2: Determine total moneys owed to each research group across all relevant elements of the fee
		
		j = 1
		
		while (j <= length(Research_Groups)) {
		## Step 1: Determine total moneys owed from GeoSciences to each programme across all elements of the fee
		Total_teaching_gs <- 0
		Total_diss_gs <- RGFinances[[j,"GeoSciences"]]
		Total_admin_gs <- 0
		Total_gs <- Total_teaching_gs + Total_diss_gs + Total_admin_gs

		Total_teaching_spss <- 0
		Total_diss_spss <- RGFinances[[j, "SPSS"]]
		Total_spss <- Total_teaching_spss + Total_diss_spss 

		Total_teaching_law <- 0
		Total_diss_law <- RGFinances[[j, "Law"]]
		Total_law <- Total_teaching_law + Total_diss_law 
		
		Total_teaching_eng <- 0
		Total_diss_eng <- RGFinances[[j, "Engineering"]]
		Total_eng <- Total_teaching_eng + Total_diss_eng 
		
		Total_teaching_bus <- 0
		Total_diss_bus <- RGFinances[[j, "Business"]]
		Total_bus <- Total_teaching_bus + Total_diss_bus
		
		Total_All <- Total_gs + Total_spss + Total_law + Total_eng + Total_bus
				
		SRUC_InstitutionalSummary[j+5,] <<- list(Total_All, Total_gs, Total_teaching_gs, Total_diss_gs, Total_admin_gs, 
															Total_spss, Total_teaching_spss, Total_diss_spss,
															Total_law, Total_teaching_law, Total_diss_law,
															Total_eng, Total_teaching_eng, Total_diss_eng,
															Total_bus, Total_teaching_bus, Total_diss_bus)
															
		row.names(SRUC_InstitutionalSummary)[j+5] <<- Research_Groups[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
		#Step 3: Determine sum totals across the whole organisation for all interactions with all schools
		SumTotal_All <-sum(SRUC_InstitutionalSummary$Total)
		SumTotal_gs <- sum(SRUC_InstitutionalSummary$Subtotal_GeoSciences)
		SumTotal_gs_t <- sum(SRUC_InstitutionalSummary$GeoSciences_Teaching)
		SumTotal_gs_d <- sum(SRUC_InstitutionalSummary$GeoSciences_Dissertations)
		SumTotal_gs_a <- sum(SRUC_InstitutionalSummary$GeoSciences_Administration)
		
		SumTotal_spss <- sum(SRUC_InstitutionalSummary$Subtotal_SPSS)
		SumTotal_spss_t <- sum(SRUC_InstitutionalSummary$SPSS_Teaching)
		SumTotal_spss_d <- sum(SRUC_InstitutionalSummary$SPSS_Dissertations)
		
		SumTotal_law <- sum(SRUC_InstitutionalSummary$Subtotal_Law)
		SumTotal_law_t <- sum(SRUC_InstitutionalSummary$Law_Teaching)
		SumTotal_law_d <- sum(SRUC_InstitutionalSummary$Law_Dissertations)
		
		SumTotal_eng <- sum(SRUC_InstitutionalSummary$Subtotal_Engineering)
		SumTotal_eng_t <- sum(SRUC_InstitutionalSummary$Engineering_Teaching)
		SumTotal_eng_d <- sum(SRUC_InstitutionalSummary$Engineering_Dissertations)
		
		SumTotal_bus <- sum(SRUC_InstitutionalSummary$Subtotal_Business)
		SumTotal_bus_t <- sum(SRUC_InstitutionalSummary$Business_Teaching)
		SumTotal_bus_d <- sum(SRUC_InstitutionalSummary$Business_Dissertations)
		
		SRUC_InstitutionalSummary[8,]<<-list(SumTotal_All, SumTotal_gs, SumTotal_gs_t, SumTotal_gs_d, SumTotal_gs_a, 
															SumTotal_spss, SumTotal_spss_t, SumTotal_spss_d,
															SumTotal_law, SumTotal_law_t, SumTotal_law_d,
															SumTotal_eng, SumTotal_eng_t, SumTotal_eng_d,
															SumTotal_bus, SumTotal_bus_t,SumTotal_bus_d)
		row.names(SRUC_InstitutionalSummary)[8] <<- "Sum Totals"
		
		SRUC_InstitutionalSummary <<-SRUC_InstitutionalSummary[c("Total", "Subtotal_GeoSciences", "GeoSciences_Teaching", "GeoSciences_Dissertations", "GeoScienes_Administration",
																	"Subtotal_SPSS", "SPSS_Teaching", "SPSS_Dissertations",
																	"Subtotal_Law", "Subtotal_Engineering", "Subtotal_Business",
																	"Law_Teaching", "Law_Dissertations",
																	"Engineering_Teaching", "Engineering_Dissertations",
																	"Business_Teaching", "Business_Dissertations")]
}
	
#To Check inputs worked		
Institutional_Summary()
SRUC_InstitutionalSummary
SRUC_InstitutionalSummary["EE",]


###PART 6: Export all relevant results to Excel workbook

#This Excel file should be a workbook with the following features:
# 1. The first page should show the total owed to SRUC across all programmes, showing split by School (GS, Engineering, SPSS, Law, Business)
# 2. The total owed to each programme showing the same split as above, and for each course 'owned' by that programme
# Subsequent worksheets should show the individual course pages for all SRUC courses

	## Step 1: Prep the worksheets that are desired in the Excel document
		#SRUC Summary Information
		SRUC_Summary <- SRUC_InstitutionalSummary
		Admin_Summary <- ProgrammeFinances_Admin
		LEES_Diss <- RGData[1]
		CS_Diss <- RGData[2]
		
		#Ecological Economics Worksheets
		EE_Admin <- SRUC_AdminData[1]
		EE_Diss <- SRUC_Student_DS[1]
		EE_Courses <- ProgrammeData_TC[1]
		FEE_Summary <- CourseData[1] 
		EV_Summary <- CourseData[2]
		AEE_Summary <- CourseData[3]
		PPP_Summary <- CourseData[4]
		EIA_Summary <- CourseData[5]
		
		#EPM Worksheets 
		EPM_Admin <- SRUC_AdminData[2]
		EPM_Diss <- SRUC_Student_DS[2]
		EPM_Courses <- ProgrammeData_TC[2]
		AQCG_Summary <- CourseData[6]
		LUEI_Summary <- CourseData[7]						    
		WRM_Summary <- CourseData[8]						    
		EVM_Summary <- CourseData[9]
		AEST_Summary <- CourseData[10]						    
								    
		#FS Worksheets
		FS_Admin <- SRUC_AdminData[3]
		FS_Diss <- SRUC_Student_DS[3]
		FS_Courses <- ProgrammeData_TC[3]
		FAFS_Summary <- CourseData[11]
		IFS_Summary <- CourseData[12]						    
		SFP_Summary <- CourseData[13]						    
								    
		#SS Worksheets 
		SS_Admin <- SRUC_AdminData[4]
		SS_Diss <- SRUC_Student_DS[4]
		SS_Courses <- ProgrammeData_TC[4]
		SPM_Summary <- CourseData[14]
		SET_Summary <- CourseData[15]						    
		SSCA_Summary <- CourseData[16]						    
								    
		#SPH Worksheets
		SPH_Admin <- SRUC_AdminData[5]
		SPH_Diss <- SRUC_Student_DS[5]
		SPH_Courses <- ProgrammeData_TC[5]
		FPH_Summary <- CourseData[17]
		FOPH_Summary <- CourseData[18]
		PHGC_Summary <- CourseData[19]						    

#Basing it on approach shown here: https://statmethods.wordpress.com/2014/06/19/quickly-export-multiple-r-objects-to-an-excel-workbook/

SRUC.PGT.AnnualFinancialSummary <- function (file, SRUC_Summary, Admin_Summary,
								EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
					    			EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					    			FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					    			SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					    			SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SRUC_Summary, Admin_Summary,
					EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
		       			EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary)
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
SRUC.PGT.AnnualFinancialSummary(paste("Outputs/SRUC_PGT_FinancialSummary_", yr, ".xlsx", sep=""), SRUC_Summary, Admin_Summary,
					EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary,
		       			EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary, 
					FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary,
					SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary,
					SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary)

#Generates excel file with just EE information
EE.PGT.AnnualFinancialSummary <- function (file, EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)
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

EE.PGT.AnnualFinancialSummary(paste("Outputs/EE_PGT_FinancialSummary_", yr, ".xlsx", sep=""), EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)

#Generates excel file with just EPM information
EPM.PGT.AnnualFinancialSummary <- function (file, EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary)
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

EPM.PGT.AnnualFinancialSummary(paste("Outputs/EPM_PGT_FinancialSummary_", yr, ".xlsx", sep=""), EPM_Admin, EPM_Diss, EPM_Courses, AQCG_Summary, LUEI_Summary, WRM_Summary, EVM_Summary, AEST_Summary)

#Generates excel file with just FS information
FS.PGT.AnnualFinancialSummary <- function (file, FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary)
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

FS.PGT.AnnualFinancialSummary(paste("Outputs/FS_PGT_FinancialSummary_", yr, ".xlsx", sep=""), FS_Admin, FS_Diss, FS_Courses, FAFS_Summary, IFS_Summary, SFP_Summary)

#Generates excel file with just SS information 
SS.PGT.AnnualFinancialSummary <- function (file, SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary)
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

FS.PGT.AnnualFinancialSummary(paste("Outputs/SS_PGT_FinancialSummary_", yr, ".xlsx", sep=""), SS_Admin, SS_Diss, SS_Courses, SPM_Summary, SET_Summary, SSCA_Summary)

#Generates excel file with just SPH information 
SPH.PGT.AnnualFinancialSummary <- function (file, SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary)
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

SPH.PGT.AnnualFinancialSummary(paste("Outputs/SPH_PGT_FinancialSummary_", yr, ".xlsx", sep=""), SPH_Admin, SPH_Diss, SPH_Courses, FPH_Summary, SOPH_Summary, PHGC_Summary)




















