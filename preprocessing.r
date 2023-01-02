#installation of package to read excel files
install.packages("readxl")

library("readxl")
library(dplyr)

#function to calculate Mode
Mode<-function(x){
  ux<-na.omit(unique(x))
  tab<-tabulate(match(x,ux)); ux[tab==max(tab)]
}

table1<- read_excel("E:\\Semester 5\\Data Science Project\\walmart table1.xlsx")

p<-data.frame(table1)

View(p)

# preprocessing of table1
p$Discount[is.na(p$Discount)]<-mean(p$Discount,na.rm=TRUE)

p<-p[,-3]



table2<- read_excel("E:\\Semester 5\\Data Science Project\\walmart table 2.xlsx")

q<-data.frame(table2)

View(q)

#preprocessing of table2

k<-Mode(q$Order.Priority)

q$Order.Priority[q$Order.Priority=="Not Specified"]=k


q$Product.Base.Margin[is.na(q$Product.Base.Margin)]<-mean(q$Product.Base.Margin,na.rm=TRUE)

q$Sales[is.na(q$Sales)]<-mean(q$Sales,na.rm=TRUE)


table3<- read_excel("E:\\Semester 5\\Data Science Project\\walmart table 3.xlsx")

r<-data.frame(table3) 

View(r)

#preprocessing of table3
r$Ship.Mode[is.na(r$Ship.Mode)]<-Mode(r$Ship.Mode)

r$Shipping.Cost[is.na(r$Shipping.Cost)]<-mean(r$Shipping.Cost,na.rm=TRUE)


#merging the tables
Final1<-merge(p,q,by="ID")

Final.Walmart<-merge(Final1,r,by="ID")

#preprocessing final data frame
df_seg<- group_by(Final.Walmart,Customer.Segment)

df_mean<-summarise(df_seg,Qty=mean(Order.Quantity, na.rm = TRUE))

for(x in df_mean$Customer.Segment)
{
  a<-is.na(Final.Walmart$Order.Quantity)
  b<-Final.Walmart$Customer.Segment==x 
  Final.Walmart$Order.Quantity[a&b]=ceiling(df_mean$Qty[df_mean$Customer.Segment==x])
}

Final.Walmart<-Final.Walmart[,-6]

#Fixing Date Format

Final.Walmart$Order.Date<-format(as.POSIXct(Final.Walmart$Order.Date), format="%d-%m-%Y")
Final.Walmart$Ship.Date<-format(as.POSIXct(Final.Walmart$Ship.Date), format="%d-%m-%Y")


#Converting to Excel File
install.packages("writexl")
library("writexl")

write_xlsx(Final.Walmart,"E:\\Semester 5\\Data Science Project\\Final_Walmart_Data.xlsx")












