library(officer)
#library(dplyr)

bd <- read.csv("BdTestGagnon - BD.csv", header=TRUE, stringsAsFactors = TRUE ); 

feqDepartement <- bd[rep(row.names(bd), bd$DEPARTEMENT),2]
#print(feqDepartement)

freqDep <-table(bd$DEPARTEMENT)

id1 <-(((bd$Q1-1)/6*100) + ((bd$Q2-1)/6*100) )/2

id2 <-(((bd$Q4-1)/6*100) + ((bd$Q5-1)/6*100) )/2

#print(id2)
myData <- data.frame(bd$DEPARTEMENT,as.data.frame(id1), as.data.frame(id2) )

print(myData)



my_pres <- read_pptx()

my_pres <- add_slide(my_pres, layout = "Title Slide", master = "Office Theme")
my_pres <- ph_with(my_pres, value = "Présentation X", location = ph_location_type(type = "ctrTitle"))

my_pres <- add_slide(my_pres, layout = "Title and Content", master = "Office Theme")


my_pres <- ph_with(my_pres, value = "Slide 1", location = ph_location_type(type = "title"))
my_pres <- ph_with(my_pres, value = "A footer", location = ph_location_type(type = "ftr"))
my_pres <- ph_with(my_pres, value = format(Sys.Date()), location = ph_location_type(type = "dt"))
my_pres <- ph_with(my_pres, value = "slide 2", location = ph_location_type(type = "sldNum"))
my_pres <- ph_with(my_pres, value = as.data.frame(freqDep) , location = ph_location_type(type = "body")) 

print(my_pres, target = "first_example.pptx") 
