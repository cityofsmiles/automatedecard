#--------------------------------------------------------------------
# This R script produces two LaTeX files to be used in printing report cards and card slips containing student infos and grades in grading periods chosen by the user.
#
#--------------------------------------------------------------------


#--------------------------------------------------------------------
# To run this script in R console, copy this:
# setwd("/root/storage/emulated/0/Documents/documents/excel/shs/19-20/for-card"); source("cards-n-slip-termux-2020.R")
#
#--------------------------------------------------------------------


#--------------------------------------------------------------------
# To run this script in bash and print only the cards, copy this:
# go 'cd /root/storage/emulated/0/Documents/documents/excel/shs/19-20/for-card && Rscript cards-n-slip-termux-2020.R && pdflatex cards.tex && termux-open cards.pdf'
#
#--------------------------------------------------------------------


#--------------------------------------------------------------------
# To run this script in bash and print only the card slips, copy this:
# go 'cd /root/storage/emulated/0/Documents/documents/excel/shs/19-20/for-card/ && Rscript cards-n-slip-termux-2020.R && pdflatex card-slips.tex && termux-open card-slips.pdf'
#
#--------------------------------------------------------------------


#--------------------------------------------------------------------
# To run this script in bash and print both the cards and card slips, copy this:
# cd /root/storage/emulated/0/Documents/documents/excel/shs/19-20/for-card; Rscript cards-n-slip-termux-2019.R; pdflatex cards.tex; pdflatex card-slips.tex
#
#--------------------------------------------------------------------


# Choose the grading periods to print and whether the cards and card slips will be printed. Type 'y' for yes or 'n' for no. 
cards <- 'n'
infos <- 'y'
firstgrading <- 'y' 
secondgrading <- 'y' 
thirdgrading <- 'y' 
fourthgrading <- 'y' 
finalgrading <- 'y' 
cardslip <- 'y' 
printprincipalsname <- 'y'

# Set the range of students whose cards or slips will be printed.
firstpage <- 1
lastpage <- 47


# Assign the excel file to read. 
ExcelFile <- '8-Hubble-summary-grades.xlsx'


##-------------------------------------------------------------##
##       Don't change anything beyond        ##
##                        this point!                           ##
##-------------------------------------------------------------##



# Load readxl package.
library(readxl) 


# Assign the filenames and range of data in excel file. 
texcardformat <- 'cards-format-center.tex' #'cards-format-left.tex' 'cards-format-center-2020.tex'
texcardoutput <- 'cards.tex'
texslipformat <- 'card-slip-format.tex'
texslipoutput <- 'card-slips.tex'
ExcelRange <- "A2:AE72"


# Read and write intro part for latex file of cards.
if (cards == "y") {

start_line_intro_card = 0

end_line_intro_card = 18

line_count_intro_card = end_line_intro_card - start_line_intro_card

texcardintro <- scan(texcardformat, '', skip = start_line_intro_card, nlines = line_count_intro_card, sep = '\n')

strIntroCard <- paste(texcardintro, collapse="\n")

cat(strIntroCard, file=texcardoutput, sep="\n", append=FALSE)
}

# Read and write intro part for latex file of card slips.
if (cardslip == "y") {

start_line_intro_slip = 0

end_line_intro_slip = 19

line_count_intro_slip = end_line_intro_slip - start_line_intro_slip

texslipintro <- scan(texslipformat, '', skip = start_line_intro_slip, nlines = line_count_intro_slip, sep = '\n')

strIntroSlip <- paste(texslipintro, collapse="\n")

cat(strIntroSlip, file=texslipoutput, sep="\n", append=FALSE)
}


# Print in the LaTeX files the infos and grading periods chosen by the user. 

if (infos == "y") {

	if (cards == "y") {
	
cat("\\def \\infos {y}
", file=texcardoutput, sep="\n", append=TRUE) }

	if (cardslip == "y") {
	
cat("\\def \\infos {y}
", file=texslipoutput, sep="\n", append=TRUE) }

} else {

	if (cards == "y") {
	
cat("\\def \\infos {n}
", file=texcardoutput, sep="\n", append=TRUE) }

	if (cardslip == "y") {
	
cat("\\def \\infos {n}
", file=texslipoutput, sep="\n", append=TRUE) }
}

if (firstgrading == "y") {
	if (cards == "y") {
cat("\\def \\firstg {y}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\firstg {y}", file=texslipoutput, sep="\n", append=TRUE) }
} else {
	if (cards == "y") {
cat("\\def \\firstg {n}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\firstg {n}", file=texslipoutput, sep="\n", append=TRUE) }
}

if (secondgrading == "y") {
	if (cards == "y") {
cat("\\def \\secondg {y}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\secondg {y}", file=texslipoutput, sep="\n", append=TRUE) }
} else {
	if (cards == "y") {
cat("\\def \\secondg {n}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\secondg {n}", file=texslipoutput, sep="\n", append=TRUE) }
}

if (thirdgrading == "y") {
	if (cards == "y") {
cat("\\def \\thirdg {y}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\thirdg {y}", file=texslipoutput, sep="\n", append=TRUE) }
} else {
	if (cards == "y") {
cat("\\def \\thirdg {n}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\thirdg {n}", file=texslipoutput, sep="\n", append=TRUE) }
}

if (fourthgrading == "y") {
	if (cards == "y") {
cat("\\def \\fourthg {y}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\fourthg {y}", file=texslipoutput, sep="\n", append=TRUE) }
} else {
	if (cards == "y") {
cat("\\def \\fourthg {n}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\fourthg {n}", file=texslipoutput, sep="\n", append=TRUE) }
}

if (finalgrading == "y") {
	if (cards == "y") {
cat("\\def \\finalg {y}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\finalg {y}", file=texslipoutput, sep="\n", append=TRUE) }
} else {
	if (cards == "y") {
cat("\\def \\finalg {n}", file=texcardoutput, sep="\n", append=TRUE) }
	if (cardslip == "y") {
cat("\\def \\finalg {n}", file=texslipoutput, sep="\n", append=TRUE) }
}

if (cards == "y") {
	if (printprincipalsname == "y") {
		cat("\\def \\printprincipalsname {y}
", file=texcardoutput, sep="\n", append=TRUE) }
	else {
		cat("\\def \\printprincipalsname {n}
", file=texcardoutput, sep="\n", append=TRUE)}
}


# Begin the LaTeX file for cards. 
if (cards == "y") {
start_line_begin_card = 19
end_line_begin_card = 22
line_count_begin_card = end_line_begin_card - start_line_begin_card
texcardbegin <- scan(texcardformat, '', skip = start_line_begin_card, nlines = line_count_begin_card, sep = '\n')
strbegincard <- paste(texcardbegin, collapse="\n")
cat(strbegincard, file=texcardoutput, sep="\n", append=TRUE)
}


# Begin the LaTeX file for card slips. 
if (cardslip == "y") {
start_line_begin_slip = 20
end_line_begin_slip = 23
line_count_begin_slip = end_line_begin_slip - start_line_begin_slip
texslipbegin <- scan(texslipformat, '', skip = start_line_begin_slip, nlines = line_count_begin_slip, sep = '\n')
strbeginslip <- paste(texslipbegin, collapse="\n")
cat(strbeginslip, file=texslipoutput, sep="\n", append=TRUE)
}


# Read body of latex file for cards.
if (cards == "y") {
start_line_body_card = 22
end_line_body_card = 276
line_count_body_card = end_line_body_card - start_line_body_card
texcardbody <- scan(texcardformat, '',  skip = start_line_body_card, nlines = line_count_body_card, sep = '\n')
string_card_body <- paste(texcardbody, collapse="\n")
}


# Read body of latex file for card slips.
if (cardslip == "y") {
start_line_body_slip = 23
end_line_body_slip = 147
line_count_body_slip = end_line_body_slip - start_line_body_slip
texslipbody <- scan(texslipformat, '',  skip = start_line_body_slip, nlines = line_count_body_slip, sep = '\n')
string_slip_body <- paste(texslipbody, collapse="\n")
}


# Set variable names.
CodeName <- c("stcode", "stname") 

main_infos_A <- c("stlrn", "stgender", "stage")

main_infos_B <- c("stgrade", "stsection", "stschoolyear", "stadviser") 

grading_periods <- c(if (firstgrading=="y") {"1"}, if (secondgrading=="y") {"2"}, if (thirdgrading=="y") {"3"}, if (fourthgrading=="y") {"4"}, if (finalgrading=="y") {"5"}, if (finalgrading=="y") {"r"}) 

subjects  <- c("fil", "eng", "math", "sci", "ap", "esp", "tle", "mapeh", "music", "arts", "pe", "health") 

subs <- c() 
index <- 0
for (i in grading_periods) {
	for (j in subjects) {
		index <- index + 1
		subs[[index]] <- paste(j, i, sep='') 
	}
}

finals_info_A <- c("avegrade", "action", "honor", "acthonup", "acthondown")

finals_info_B <- c("checker", "stprincipal")

namesofvals <- c(CodeName, if (infos=="y") {main_infos_A}, subs, if (finalgrading=="y") {finals_info_A}, if (infos=="y") {main_infos_B}, if (finalgrading=="y") {finals_info_B})  


# Read the data from the excel files. 
periods_card <- c(infos, firstgrading, secondgrading, thirdgrading, fourthgrading, finalgrading) 

periods_excel <- c("forcard", "1st grading", "2nd grading", "3rd grading", "4th grading", "finals") 

indx <- 0
count <- 0
DF <- c() 
for (i in periods_card) {
	indx <- indx + 1
	if (i == "y") {
	count <- count + 1
	index <- indx 
	DF[[count]] <- read_excel(ExcelFile, sheet = periods_excel[[index]], range=ExcelRange, col_names = TRUE)
	DF[[count]] <- DF[[count]][, colSums(is.na(DF[[count]])) != nrow(DF[[count]]), ]
	DF[[count]] <- DF[[count]][rowSums(is.na(DF[[count]])) != ncol(DF[[count]]), ]
	}
}

infosA <- read_excel(ExcelFile, sheet = "info", range="B1:B6", col_names = F)

infosA <- c(infosA) 

names(infosA[[1]])  <- c(main_infos_B, finals_info_B) 


# Merge the dataframes. 
MergedDF <- DF[[1]]
if (length(DF) > 1) {
for (j in 2:length(DF)) {
	MergedDF <- merge(MergedDF,DF[[j]], by=c("code","name"), sort=FALSE)
	}
}


# Assign the values in the dataframe into the variables.
for (i in firstpage:lastpage){
	if (cards == "y") {
 	str_card <- string_card_body
 	}
 	if (cardslip == "y") {
	str_slip <- string_slip_body
	}
 	for (j in 1:ncol(MergedDF)){
 		if (cards == "y") {
 		str_card <- gsub(namesofvals[[j]], toString(MergedDF[i,j]), str_card)
 		}
 		if (cardslip == "y") {
 		str_slip <- gsub(namesofvals[[j]], toString(MergedDF[i,j]), str_slip)
 		}
 	}

if (infos=="y") {
for (k in 1:4) {
	if (cards == "y") {
	str_card <- gsub(names(infosA[[1]][k]), toString(infosA[[1]][k]), str_card)
	}
	if (cardslip == "y") {
	str_slip <- gsub(names(infosA[[1]][k]), toString(infosA[[1]][k]), str_slip)
	}
}
}

if (finalgrading=="y") {
	if (cards == "y") {
		if (printprincipalsname=="y") {
			str_card <- gsub(names(infosA[[1]][5]), toString(infosA[[1]][5]), str_card)
			str_card <- gsub(names(infosA[[1]][6]), toString(infosA[[1]][6]), str_card)
		} else {
			str_card <- gsub(names(infosA[[1]][5]), toString(infosA[[1]][5]), str_card)
		}
	}
	if (cardslip == "y") {
	str_slip <- gsub(names(infosA[[1]][5]), toString(infosA[[1]][5]), str_slip)
	}
}



# Print the values in the latex file of cards.
if (cards == "y") {
cat(str_card, file=texcardoutput, sep="\n", append=TRUE)
}

# Print the values in the latex file of card slips.
if (cardslip == "y") {
cat(str_slip, file=texslipoutput, sep="\n", append=TRUE)
}
}


# End the LaTeX files.
if (cards == "y") {
cat("\\end{document}", file=texcardoutput, sep="\n", append=TRUE)
}

if (cardslip == "y") {
cat("\\end{document}", file=texslipoutput, sep="\n", append=TRUE)
}

