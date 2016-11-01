##Notes for future...
###R starts with column 0, but row 1 for data frame with headers

#Insructions are found in the 'Information to Use R script.txt' file

# DO NOT CHANGE THIS SCRIPT EXCEPT FOR THE VALUE OF THE YEAR IN STEP 1B!!!
#_______________________________________________________________________________________________________________________________

###PART 1: SET UP

# Step 1a: Set working directory to match the folder from Step 0, changing the code below as required, and load package xlsx

setwd("C:/Users/Owner/Desktop/SRUC MOU R Test/")

#at work: "G:/CKD BAULCOMB/MSc Involvement/MOU Test documents/"
library(xlsx)

# Step 1b: Initialise the following objects

yr <- read.xlsx("Year for Calculation.xlsx", sheetName="Sheet1", rowIndex=1, colIndex=1, header=FALSE)
yr <- yr[1,1]
SRUC_Courses <- read.xlsx("SRUC Courses.xlsx", sheetName="Sheet1", header=TRUE, as.data.frame=TRUE)
Courses <- SRUC_Courses[,2]
Programme_Ownership <- SRUC_Courses[,3]
Credit_Weighting <- SRUC_Courses[,6]

Programmes <- c("EE", "EPM", "FS", "SS", "SPH")	
Research_Groups <- c("LEES", "CropsSoils")

###PART 2: Calculations related to teaching

# Step 1: Import data on tuition fees and students attending courses, mergin and re-arranging to see students, program details, fee status, and fees for each course
#This function will import the fee schedule and attendance list as long as the working directory is specified
	
ImportData <- function () {
	#Step 1: Import csv file as data frame showing all programmes, School, Home Fees (FT), Overseas Fees (FT)
	# csv version: TuitionFees <<- as.data.frame(read.csv("Inputs/Fees_2016.csv", header=TRUE, sep=","))
	TuitionFees <<- as.data.frame(read.xlsx("Inputs/Fees_2016.xlsx", header=TRUE))
		
	#Step 2: Put all of the fee related information within one column (this is necessary for later)
	TuitionFees <<- cbind(TuitionFees[gl(nrow(TuitionFees), 1, 2*nrow(TuitionFees)), 1:2], stack(TuitionFees[,3:4]))
		
	##Rename the columns from the defaults to what they are to allow merging later
	names(TuitionFees)[names(TuitionFees)=="ind"] <<- "Fee_Status"
	names(TuitionFees)[names(TuitionFees)=="values"] <<- "Tuition"
		
	# Step 3: Import attendance lists for all courses, merge with fee info, and calculate fee fraction for each student on each course
	i = 1
	
	CourseData <<- vector('list', length(Courses))
		
	## Imports attendance list
	while (i <= length(Courses)) {
		## Imports attendance list
		fn <- paste("Inputs/", Courses[i], "CLASS LIST", yr, ".xlsx", sep=" ")
		## Creates dataframe associated with course that holds position i in Courses
		CourseData[[i]] <<-read.xlsx(fn, header=TRUE, as.data.frame=TRUE)
		## removes last column (signature column)
		CourseData[[i]] <<-(CourseData[[i]])[,-8]
		# removes unnecessary 5th column (matriculation)
		CourseData[[i]] <<-(CourseData[[i]])[,-5]
		# removes the 'new' 5th column (formly 6th column on CE vs. CA)
		CourseData[[i]] <<-(CourseData[[i]])[,-5]
		## renames remaining columns
		names(CourseData[[i]])[,1]<<-"UUN"
		names(CourseData[[i]])[,2]<<-"Surname"
		names(CourseData[[i]])[,3]<<-"Forename"
		names(CourseData[[i]])[,4]<<-"Programme"
		names(CourseData[[i]])[,5]<<-"School"
		# replace school names used by default in attendance list with ones matching the fees sheet
		# used this tip: http://stackoverflow.com/questions/22418864/r-replace-entire-strings-based-on-partial-match
		(CourseData[[i]])[grepl("School Of Geosciences", (CourseData[[i]]), ignore.case=FALSE)] <<- "GeoSciences"
		(CourseData[[i]])[grepl("School Of Social And Political Science", (CourseData[[i]]), ignore.case=FALSE)] <<- "Social & Political Science"
		(CourseData[[i]])[grepl("School Of Engineering", (CourseData[[i]]), ignore.case=FALSE)] <<- "Engineering"
		
		#####need to do the following here:
		# school of law doesn't show up in the attendance list, so need to find way to ID them
		# Need to figure out what CA is before delete it
		# Need to see how auditing students show up and then remove them from the list (as we don't get paid for them)
		# need to see how the Business school would show up in classl ist to see if need to change school name
		
		## Merges attendance list with fee information
		CourseData[[i]] <<-merge(CourseData[[i]], TuitionFees[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Inputs credit weighting for course 
		CourseData[[i]][,8]<<-(Credit_Weighting[[i]])
		## Names this column
		names(CourseData[[i]])[names(CourseData[[i]])=="V8"]<<-"Credit_Weighting"
		## Re-orders attendance list with fee information so it's easier to read
		CourseData[[i]] <<-CourseData[[i]][c("UUN", "Surname", "Forename", "Programme", "School", "Fee_Status", "Tuition", "Credit_Weighting")]
		## Calculates portion of total fee associated with each student on the course
		CourseData[[i]][,9]<<-(0.05 * CourseData[[i]][,7] * CourseData[[i]][,8])
		## Names this column to highlight the fee portion due to each student on the course
		names(CourseData[[i]])[names(CourseData[[i]])=="V9"]<<-"Course_Fee"
			
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	
	names(CourseData) <<- Courses
	
}

#To execute:
ImportData()

#To Check inputs worked
TuitionFees
CourseData["FEE"]
FEE <- CourseData[[1]]
CourseData[[2]]
CourseData[[3]]
CourseData[[4]]
CourseData[[5]]

#Step 2: Income associated with teaching individual courses
Course_Level_Finances <- function() {
	
	i = 1
	
	CourseFinances <<- data.frame(All=numeric(), GS=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
	while (i <= length(Courses)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		gs <- subset(CourseData[[i]], School == "GeoSciences")
		spss <- subset(CourseData[[i]], School == "SPSS")
		law <- subset(CourseData[[i]], School == "Law")			
		eng <- subset(CourseData[[i]], School == "Engineering")
		bus <- subset(CourseData[[i]], School == "Business")
			
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All <- sum(CourseData[[i]]$Course_Fee)
		Total_gs <- sum(gs$Course_Fee)
		Total_spss <- sum(spss$Course_Fee)
		Total_law <- sum(law$Course_Fee)
		Total_eng <- sum(eng$Course_Fee)
		Total_bus <- sum(bus$Course_Fee)
		
		CourseFinances[i,] <<- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus)
		row.names(CourseFinances)[i] <<- Courses[i]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	CourseFinances <<- data.frame( Programme_Ownership, CourseFinances)
}

#To Check inputs worked
Course_Level_Finances()
CourseFinances["FEE",]
	
#Step 3: Income associated with teaching for individual programmes
Programme_Level_Finances_Teaching <- function() {
	
	i = 1
	
	#Creates subsets of CourseFinances to match the courses owned by individual programmes
	ProgrammeData_TC <<- vector('list', length(Courses))
	while(i <= length(Programmes)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		ProgrammeData_TC[[i]] <<- subset(CourseFinances, Programme_Ownership == Programmes[[i]])
			
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	#Allows one to call summary finance table by programme name
	names(ProgrammeData_TC) <<- Programmes
	
	#Creates a summary table of programme finances
	j=1
	ProgrammeFinances_TC <<- data.frame(All=numeric(), GS=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		#Within the dataframe showing courses owned by programmes, sum the relevant tuition fee components by column
		Total_All <- sum(ProgrammeData_TC[[j]]$All)
		Total_gs <- sum(ProgrammeData_TC[[j]]$GS)
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
EE_Dissertations2016

		
###PART 3: Calculations related to dissertation supervision
	
#Step 1: Calculate and apportion income for SRUC students associated with dissertation supervision
SRUC_Prog_DS <- function() {

	i = 1
	
	SRUC_Student_DS <<- vector('list', length(Programmes))
		
	while (i <= length(Programmes)) {
		## Imports csv for each SRUC programme showing supervision details
		fn <- paste("Inputs/", Programmes[i], "_Dissertations", yr, ".csv", sep="")
		## Creates dataframe associated with course that holds position i in Courses
		SRUC_Student_DS[[i]]<<- as.data.frame(read.csv(fn, header=TRUE, sep=","))
		## Merges supervision list with fee information
		SRUC_Student_DS[[i]]<<-merge(SRUC_Student_DS[[i]], TuitionFees[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Re-orders supervision list with fee information so it's easier to read
		SRUC_Student_DS[[i]]<<-SRUC_Student_DS[[i]][c("UUN", "Surname", "Forename", "Programme", "School", "Supervisor", "Organisation", "Detail", "Fee_Status", "Tuition")]
		## Calculates portion of total fee associated with each student's supervision
		SRUC_Student_DS[[i]][,11]<<-(0.10 * SRUC_Student_DS[[i]][,10])
		## Names this column to highlight the fee portion due to each student on the programme for supervision
		names(SRUC_Student_DS[[i]])[names(SRUC_Student_DS[[i]])=="V11"]<<-"Supervision_Fee"
		
		## Advances to the next course and repeats above steps until the list of programmes is exhausted
		i = i+1
	}
	
	names(SRUC_Student_DS) <<- Programmes
	
	# Calculates allocation of disertation supervision funds for SRUC students
	j = 1
	
	ProgrammeFinances_SRUCstudent_DS <<- data.frame(All=numeric(), SRUC=numeric(), GeoSciences=numeric(), stringsAsFactors=FALSE)
	
	while (j <= length(Programmes)) {
		#Step 1: Define subsets of dataframes to group students from different schools on each course
		sruc <- subset(SRUC_Student_DS[[j]], Organisation == "SRUC")
		gs <- subset(SRUC_Student_DS[[j]], Organisation == "University")
					
		#Step 2: Determine the tuition associated with each course (in total and by school)
		Total_All <- sum(SRUC_Student_DS[[j]]$Supervision_Fee)
		Total_sruc <- sum(sruc$Supervision_Fee)
		Total_gs <- sum(gs$Supervision_Fee)
				
		ProgrammeFinances_SRUCstudent_DS[j,] <<- list(Total_All, Total_sruc, Total_gs)
		row.names(ProgrammeFinances_SRUCstudent_DS)[j] <<- Programmes[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
}
	
#To Check inputs worked		
SRUC_Prog_DS()
SRUC_Student_DS
ProgrammeFinances_SRUCstudent_DS
ProgrammeFinances_SRUCstudent_DS["EE",]
	
#Step 2: Calculates supervision fees associated with SRUC staff supervising non-SRUC students
SRUC_ExternalStudent_DS <- function() {
	
	## Imports csv for showing external student supervision details
	fn <- paste("Inputs", "SRUC_ExternalDissertations", yr, ".csv", sep="")
	SRUC_ExternalStudent_DS <<- as.data.frame(read.csv(fn, header=TRUE, sep=","))
	## Merges supervision list with fee information
	SRUC_ExternalStudent_DS <<-merge(SRUC_ExternalStudent_DS, TuitionFees[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
	## Re-orders supervision list with fee information so it's easier to read
	SRUC_ExternalStudent_DS<<-SRUC_ExternalStudent_DS[c("UUN", "Surname", "Forename", "Programme", "School", "Supervisor", "Research_Group", "Fee_Status", "Tuition")]
	## Calculates portion of total fee associated with each student's supervision
	SRUC_ExternalStudent_DS[,10]<<-(0.10 * SRUC_ExternalStudent_DS[,9])
	## Names this column to highlight the fee portion due to each student on the programme for supervision
	names(SRUC_ExternalStudent_DS)[names(SRUC_ExternalStudent_DS)=="V10"]<<-"Supervision_Fee"
	
	#Provides data on students by supervisor research group
	i = 1
	RGData <<- vector('list', length(Research_Groups))
	
	while (i <= length(Research_Groups)) {
		## Separates out the subsets associated with each Research Group
		RGData[[i]] <<- subset(SRUC_ExternalStudent_DS, Research_Group == Research_Groups[[i]])
				
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		i = i+1
	}
	
	names(RGData) <<- Research_Groups
	
	# Creates summary financial picture by research group 
	j = 1
	RGFinances <<- data.frame(All=numeric(), GS=numeric(), SPSS=numeric(), Law=numeric(), Engineering=numeric(),Business=numeric(), stringsAsFactors=FALSE)
	
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
		
		RGFinances[j,] <<- list(Total_All, Total_gs,Total_spss, Total_law, Total_eng, Total_bus)
		row.names(RGFinances)[j] <<- Research_Groups[j]
		
		## Advances to the next course and repeats above steps until the list of courses is exhausted
		j = j+1
	}
		
}

#To Check inputs worked
SRUC_ExternalStudent_DS()
SRUC_ExternalStudent_DS
RGData
RGFinances
RGFinances["LEES",]

 	
###PART 4: Calculations related to administration

# In order to avoid charging the administration fee for PT students more than once, and to simplify calculations, charge the admin fee for them only in the year when they 
# work on their dissertations (rather than annually based on a fee fraction)

SRUC_Admin <- function() {

	i = 1
	
	SRUC_AdminData <<- vector('list', length(Programmes))
		
	while (i <= length(Programmes)) {
		## Imports csv for each SRUC programme showing supervision details
		fn <- paste("Inputs/",Programmes[i], "_Dissertations", yr, ".csv", sep="")
		## Creates dataframe associated with course that holds position i in Courses
		SRUC_AdminData[[i]]<<- as.data.frame(read.csv(fn, header=TRUE, sep=","))
		## Merges supervision list with fee information
		SRUC_AdminData[[i]]<<-merge(SRUC_AdminData[[i]], TuitionFees[ , c("Tuition", "Programme", "Fee_Status")], by=c("Programme", "Fee_Status"))
		## Re-orders supervision list with fee information so it's easier to read
		SRUC_AdminData[[i]]<<-SRUC_AdminData[[i]][c("UUN", "Surname", "Forename", "Programme", "School", "Supervisor", "Organisation", "Detail", "Fee_Status", "Tuition")]
		## Calculates portion of total fee associated with admin for each student
		SRUC_AdminData[[i]][,11]<<-(0.40 * SRUC_AdminData[[i]][,10])
		## Names this column to highlight the fee portion due for admin
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V11"]<<-"Admin_Fee_Total"
		## Calculates portion of total admin fee that belongs to SRUC and GeoSciences
		SRUC_AdminData[[i]][,12]<<-(0.75 * SRUC_AdminData[[i]][,11])
		SRUC_AdminData[[i]][,13]<<-(0.25 * SRUC_AdminData[[i]][,11])
		## Re-Names these columns 
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V12"]<<-"Admin_Fee_SRUC"
		names(SRUC_AdminData[[i]])[names(SRUC_AdminData[[i]])=="V13"]<<-"Admin_Fee_GeoSciences"
		
		## Advances to the next course and repeats above steps until the list of programmes is exhausted
		i = i+1
	}
	
	names(SRUC_AdminData) <<- Programmes
	
	# COlates summary information on admin fee by programme and SRUC vs. GeoSciences
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
							
		## Step 1: Determine total moneys owed from GeoSciences to each programme across all elements of the fee
		Total_teaching_gs <- ProgrammeFinances_TC[[i,"GS"]]
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
		Total_diss_gs <- RGFinances[[j,"GS"]]
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
		
		#EPM Worksheets (complete when know courses)
		EPM_Admin <- SRUC_AdminData[2]
		EPM_Diss <- SRUC_Student_DS[2]
		EPM_Courses <- ProgrammeData_TC[2]
		
		#FS Worksheets(complete when know courses)
		FS_Admin <- SRUC_AdminData[3]
		FS_Diss <- SRUC_Student_DS[3]
		FS_Courses <- ProgrammeData_TC[3]
		
		#SS Worksheets (complete when know courses)
		SS_Admin <- SRUC_AdminData[4]
		SS_Diss <- SRUC_Student_DS[4]
		SS_Courses <- ProgrammeData_TC[4]
		
		#SPH Worksheets (complete when know courses)
		SPH_Admin <- SRUC_AdminData[5]
		SPH_Diss <- SRUC_Student_DS[5]
		SPH_Courses <- ProgrammeData_TC[5]

#Basing it on approach shown here: https://statmethods.wordpress.com/2014/06/19/quickly-export-multiple-r-objects-to-an-excel-workbook/

SRUC.PGT.AnnualFinancialSummary <- function (file, SRUC_Summary, Admin_Summary,
													EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary) {
	require(xlsx, quietly=TRUE)
	objects <- list(SRUC_Summary, Admin_Summary,
					EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)
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
SRUC.PGT.AnnualFinancialSummary("Outputs/SRUC_PGT_FinancialSummary_2016.xlsx", SRUC_Summary, Admin_Summary,
								EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)

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

EE.PGT.AnnualFinancialSummary("Outputs/EE_PGT_FinancialSummary_2016.xlsx", EE_Admin, EE_Diss, LEES_Diss, EE_Courses, FEE_Summary, EV_Summary, AEE_Summary, PPP_Summary, EIA_Summary)

#Generates excel file with just EPM information (complete when know courses)
EPM.PGT.AnnualFinancialSummary <- function (file, EPM_Admin, EPM_Diss, EPM_Courses, ) {
	require(xlsx, quietly=TRUE)
	objects <- list(EPM_Admin, EPM_Diss, EPM_Courses, )
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

EPM.PGT.AnnualFinancialSummary("Outputs/EPM_PGT_FinancialSummary_2016.xlsx", EPM_Admin, EPM_Diss, EPM_Courses,)

#Generates excel file with just FS information (complete when know courses)
FS.PGT.AnnualFinancialSummary <- function (file, FS_Admin, FS_Diss, FS_Courses, ) {
	require(xlsx, quietly=TRUE)
	objects <- list(FS_Admin, FS_Diss, FS_Courses, )
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

FS.PGT.AnnualFinancialSummary("Outputs/FS_PGT_FinancialSummary_2016.xlsx", FS_Admin, FS_Diss, FS_Courses,)

#Generates excel file with just SS information (complete when know courses)
SS.PGT.AnnualFinancialSummary <- function (file, SS_Admin, SS_Diss, SS_Courses, ) {
	require(xlsx, quietly=TRUE)
	objects <- list(SS_Admin, SS_Diss, SS_Courses, )
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

FS.PGT.AnnualFinancialSummary("Outputs/SS_PGT_FinancialSummary_2016.xlsx", SS_Admin, SS_Diss, SS_Courses,)

#Generates excel file with just SPH information (complete when know courses)
SPH.PGT.AnnualFinancialSummary <- function (file, SPH_Admin, SPH_Diss, SPH_Courses, ) {
	require(xlsx, quietly=TRUE)
	objects <- list(SPH_Admin, SPH_Diss, SPH_Courses, )
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

SPH.PGT.AnnualFinancialSummary("Outputs/SPH_PGT_FinancialSummary_2016.xlsx", SPH_Admin, SPH_Diss, SPH_Courses,)




















