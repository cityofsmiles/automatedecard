# setwd("/storage/emulated/0/Documents/documents/latex/1920-Bheng/"); source("card-bheng-info.R")

#--
# This R script produces two LaTeX files to be used in printing report cards containing student infos. 
#--

#--
# To run this script, copy this:
# Rscript card-info.R && pdflatex card-info-to-print.tex && okular card-info-to-print.pdf


# Load readxl package.
library(readxl) 

# Set the range of students whose cards will be printed.
firstpage=1
lastpage=27

# Assign the filenames. 
ExcelFile <- 'student-data.xlsx'
texcardformat <- 'card-info-format.tex'  
texcardoutput <- 'card-info-to-print.tex'

# Set variable names.
infos <- c("stlrn", "stschoolyear", "stname", "stage", "stgender", "stgrade", "stsection",  "stadviser", "stprincipal")

# Read excel file. 
DF <- read_excel(ExcelFile, sheet = "for-card", range="A8:I41", col_names = TRUE)

# Read and write intro part for latex file of cards.

start_line_intro_card = 0
end_line_intro_card = 21
line_count_intro_card = end_line_intro_card - start_line_intro_card
texcardintro <- scan(texcardformat, '', skip = start_line_intro_card, nlines = line_count_intro_card, sep = '\n')
strIntroCard <- paste(texcardintro, collapse="\n")
cat(strIntroCard, file=texcardoutput, sep="\n", append=FALSE)

# Read body of latex file for cards.
start_line_body_card = 21
end_line_body_card = 52
line_count_body_card = end_line_body_card - start_line_body_card
texcardbody <- scan(texcardformat, '',  skip = start_line_body_card, nlines = line_count_body_card, sep = '\n')
string_card_body <- paste(texcardbody, collapse="\n")

# Assign the values in the dataframe into the variables.

for (i in firstpage:lastpage){
str_card <- string_card_body
	for (j in 1:ncol(DF)){
		str_card <- gsub(infos[[j]], toString(DF[i,j]), str_card)
	}


# Print the values in the latex file of cards.

cat(str_card, file=texcardoutput, sep="\n", append=TRUE)
}

# End the LaTeX file.
cat("}}
\\end{document}", file=texcardoutput, sep="\n", append=TRUE)


