library(dplyr)
library(ReporteRs)
library(ggplot2)
library(ggthemes)
library(broom)


setwd("C:/Users/Greg/Documents/ReportRs 9-15-2016")


# Get Data from the site: 
# http://www.amstat.org/publications/jse/jse_data_archive.htm
Salary_Data = read.delim("baseball.dat.txt", sep = " ", header = FALSE, quote = "\"")


#Data is a fixed with file, with widths specified by the numbers in that argument. Negative numbers indicates columns
#to be skipped
Salary_Data = read.fwf("baseball.dat.txt",
    widths = c(4, -1, 5, -1,5,-1,3,-1,3,-1,2,-1,2,-1,2,-1,3,-1,3,-1,3,-1,2,-1,2,-1,1,-1,1,-1,1,-1,1,-2,17,-1))


#give variables names to meaningful 
colnames(Salary_Data) = c("Salary", "BA", "OBP", "R", "H", "Doubles", "Triples", "HR", "RBI", "BB", "SO", "SB", "E",
                          "FA", "FA91_92", "Arbitration_El", "Arb91_92", "Name")



#General Comments:
#THe point of this exercise is to give a general example of how to step through the analytic process and show how ReportRs can aid in this process as well as 
#helping to ease the work involved in making reptitive slides

#Step 1: Before we make a model, lets look @ scatter plots of all the columns vs Salary
Test_Scatter = ggplot(Salary_Data)+
  geom_point(aes_string(y = "Salary", x = "BA"))+
  theme_minimal()+
  xlab("BA")+
  ylab("Salary")
Test_Scatter

#But what if I want to see all of them...and put them in Slides so I can look @ them with colleagues?

#REPORT R's!!!!

# CHeck out http://davidgohel.github.io/ReporteRs/powerpoint.html#complete_example
#  for some good examples

#Step 1 create a power point Document:
mydoc = pptx()

# check my layout names:
slide.layouts(mydoc)

#See how a layout actually looks:
slide.layouts(mydoc,"Title and Content" )
slide.layouts(mydoc,"Content with Caption")

#Step 2 Create a new slide
#I'll use the plot above as well as incorporating basic stats with in text:

Avg = mean(Salary_Data[,"BA"])
Med = median(Salary_Data[,"BA"])
Min = min(Salary_Data[,"BA"])
max = max(Salary_Data[,"BA"])
Correlation = cor(Salary_Data[,"Salary"], Salary_Data[,"Arbitration_El"])

#Here I create a text vector with the various stats
Stats_Text = c(paste("The mean", "BA", "is", round(Avg,3)),
               paste("The median", "BA", "is", round(Med,3)),
               paste("The minimum", "BA", "is", round(Min,3)),
               paste("The maximum", "BA", "is", round(max,3)),
               paste("The the correlation between Salary and", "BA", "is", round(Correlation,3)))
               
#Here is a one slide example of how to create and write a power point document.  The function
#names are very straight forward.  
mydoc <- mydoc %>% 
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "Salary vs BA" ) %>% 
  addPlot( function( ) print( Test_Scatter ) ) %>% 
  addPageNumber() %>% 
  addDate( )%>%
  addParagraph(value = Stats_Text)

writeDoc(mydoc, file = "./PPT_Files/Test1.pptx")


#Now to it's pretty easy to convert that to a loop go go through all the examples
Loop_Doc = pptx()

for(i in colnames(Salary_Data[,c(-1,-18)])){
  Stats_Text = c(paste("The mean", i, "is", round(Avg,3)),
                 paste("The median", i, "is", round(Med,3)),
                 paste("The minimum", i, "is", round(Min,3)),
                 paste("The maximum", i, "is", round(max,3)),
                 paste("The the correlation between Salary and", i, "is", round(Correlation,3)))
  Loop_Scatter = ggplot(Salary_Data)+
    geom_point(aes_string(y = "Salary", x = i))+
    theme_minimal()+
    xlab(i)+
    ylab("Salary")

  Loop_Doc <- Loop_Doc %>% 
    addSlide( slide.layout = "Content with Caption" ) %>% 
    addTitle( paste("Salary vs", i) ) %>% 
    addPlot( function( ) print( Loop_Scatter ) ) %>% 
    addPageNumber() %>% 
    addDate( )%>%
    addParagraph(value = Stats_Text)
  
}
  
writeDoc(Loop_Doc, file = "./PPT_Files/loop_Example.pptx")
  
#Ok now that we have looked at the variables, we can start buijlding a simple model.

#Lets start with an all inclusive model
#This model uses every variable except name to predict Salary, it also includes 2nd order interactions
Sal_Model = lm(data = Salary_Data, formula = Salary ~ (.-Name)^2)

summary(Sal_Model)

#The Broom Package produces nice a nice data frame to clean up model interpretation. You can do this with base R, but its not as concise.
Tidy_Model_Output = tidy(Sal_Model)

# Let's create a new power point with a list of all the term, and a few diagnostic plots:
Fit_PPT = pptx()
#Table is too long to fit on one slide, so I've set an arbitrary length.
Page_Length = 30

#Loop through all the different terms in the model and add a new page with a table on it
for(i in 1:ceiling(nrow(Tidy_Model_Output)/Page_Length)){
options( "ReporteRs-fontsize" = 12 )
if(i*Page_Length > nrow(Tidy_Model_Output)){ Table_Data = Tidy_Model_Output[(i*Page_Length-(Page_Length-1)):nrow(Tidy_Model_Output),]}
  else{
Table_Data = Tidy_Model_Output[(i*Page_Length-(Page_Length-1)):(i*Page_Length),]
}
# Create a FlexTable with data.frame mtcars, display rownames
# use different formatting properties for header and body cells
MyFTable <- FlexTable( data = Table_Data, add.rownames = FALSE, 
                       body.cell.props = cellProperties( border.color = "#EDBD3E"), 
                       header.cell.props = cellProperties( background.color = "#5B7778" )
) %>% 
  setZebraStyle( odd = "#DDDDDD", even = "#FFFFFF" ) %>% # zebra stripes - alternate colored backgrounds on table rows
  setFlexTableWidths( widths = c(1, rep(1.5, 4)) ) %>% 
  setFlexTableBorders( inner.vertical = borderProperties( color="#EDBD3E", style="dotted" ), 
                       inner.horizontal = borderProperties( color = "#EDBD3E", style = "none" ), 
                       outer.vertical = borderProperties( color = "#EDBD3E", style = "solid" ), 
                       outer.horizontal = borderProperties( color = "#EDBD3E", style = "solid" )
  ) # applies a border grid on table

MyFTable[ Table_Data$p.value < .05, 1:5] = chprop( myCellProps, background.color = '#e4ed29')
MyFTable

Fit_PPT <- Fit_PPT%>%
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "FlexTable example" ) %>% 
  addFlexTable( MyFTable )%>%
  addParagraph(value = paste("Model Terms",(i*Page_Length-(Page_Length-1)),"-", i*20))

}


#Add Diagnostic Plots for the Model

Fit_PPT <- Fit_PPT %>% 
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "Diagnostic Plots" ) %>% 
  addPlot( function( ) print( plot(Sal_Model,which = 1) ) ) %>% 
  addPageNumber() %>% 
  addDate( )
 

Fit_PPT <- Fit_PPT %>% 
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "Diagnostic Plots" ) %>% 
  addPlot( function( ) print( plot(Sal_Model,which = 2) ) ) %>% 
  addPageNumber() %>% 
  addDate( )

Fit_PPT <- Fit_PPT %>% 
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "Diagnostic Plots" ) %>% 
  addPlot( function( ) print( plot(Sal_Model,which = 3) ) ) %>% 
  addPageNumber() %>% 
  addDate( )

Fit_PPT <- Fit_PPT %>% 
  addSlide( slide.layout = "Content with Caption" ) %>% 
  addTitle( "Diagnostic Plots" ) %>% 
  addPlot( function( ) print( plot(Sal_Model,which = 4) ) ) %>% 
  addPageNumber() %>% 
  addDate( )

writeDoc(Fit_PPT,file = "./PPT_Files/Interactions_2.pptx" )



