

setwd('//gtc-docsrv01/GTC_Share/Commercial/Sales report and target tracking/Sales Performance Report/')
library(xlsx)
library(dplyr)
library(ggplot2)
library(reshape2)
#filename <- paste('//gtc-docsrv01/GTC_Share/Data Analytics/Analysis/AccountsMonthly/SalesWeekly/data/', as.Date(Sys.time()), '.xlsx', sep = '')
filename <- paste('//gtc-docsrv01/GTC_Share/Commercial/Sales report and target tracking/Sales Performance Report/Sales_WeeklyUpdate_V9', '.xlsm', sep = '')
# load package for sql
#library(DBI)
library(RODBC)
# connect to database
odbcChannel <- odbcConnect('echo_core', uid='Daria Alekseeva', pwd='Welcome01')
#odbcChannel <- odbcConnect('dr sql', uid='Daria Alekseeva', pwd='Welcome01')

salesdata <- sqlQuery( odbcChannel, 
                         "
                         Declare @ReportDateStart date
                         Declare @ReportDateEnd date
                         Declare @1 date
 
                         set @ReportDateEnd=getdate()
                         Set @ReportDateStart=dateadd(mm,-12,getdate())
                         Set @1=@ReportDateStart;
                        
                         Select
                         @ReportDateStart as 'ReportDateStart',
                         @ReportDateEnd as 'ReportDateEnd',
                         ca.dateOpened as 'OpeningDate', 
                         ca.name 'AccountName',
                         isnull((select ca1.name from echo_core_prod..customer_accounts ca1 where ca.parent_id is not null and ca.parent_id=ca1.id), ca.name) as 'Parent/AccountName', 
                         isnull((select ca1.number from echo_core_prod..customer_accounts ca1 where ca.parent_id is not null and ca.parent_id=ca1.id), ca.number) as 'Parent/AccountNumber', 
                         i.fullName as 'SalesPerson',
                         ca.phone,
                         ca.contact,
                         (Select sum(totalNetPrice) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate between @ReportDateStart and @ReportDateEnd and jobStatus in (7,10)) as 'TotalNetPrice',
                         (Select sum(totalCharge) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate between @ReportDateStart and @ReportDateEnd and jobStatus in (7,10)) as 'TotalCharge',
                         (select count(id) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate between @ReportDateStart and @ReportDateEnd and jobStatus in (7,10)) as 'TotalJobsReportPeriod',
                         (Select sum(totalPrice) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate between @ReportDateStart and @ReportDateEnd and jobStatus in (7,10)) as 'TotalSpendReportPeriod', 
                         (select count(id) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate > @ReportDateEnd and jobStatus not in (7,10)) as 'FutureJobs',
                         (select max(jobdate) from echo_core_prod..jobs where customer_account_id=ca.id and jobDate between @ReportDateStart and @ReportDateEnd and jobStatus =7) as 'LastJob'
                         from echo_core_prod..customer_accounts ca, echo_core_prod..customer_grades cg, echo_core_prod..individuals i, echo_core_prod..individuals i1, echo_core_prod..depots d, echo_core_prod..invoicing_settings invs, echo_core_prod..pricing_groups pg
                         where 
                         cg.id=ca.grade_id and
                         i.id=ca.salesman_id and
                         i1.id=ca.manager_id and 
                         d.id=ca.depot_id and
                         invs.id=ca.invoicing_settings_id and
                         ca.pricing_group_id=pg.id and
                         ca.dateOpened > @1
                         order by LastJob desc
                         ")



data <- sqlQuery( odbcChannel, 
                  "
                  Declare @ReportDateStart date
                  Declare @ReportDateEnd date
                  
                  set @ReportDateEnd=getdate()
                  Set @ReportDateStart=dateadd(mm,-12,getdate())
                  
                  
                  select 
                   @ReportDateStart 'ReportDateStart', @ReportDateEnd 'ReportDateEnd', j.id, j.totalnetprice, j.totalCharge, i.fullName,convert(date,j.jobdate) 'jobDate', datepart(ISO_week,j.jobdate) +2 'JobWeek', datepart(month,j.jobdate) 'JobMonth', datepart(year,j.jobdate) 'JobYear',ca.name,ca.number,convert(date,ca.dateOpened) 'dateOpen'
                  from Echo_Core_Prod..jobs j
                  left join Echo_Core_Prod..customer_accounts ca on ca.id = j.customer_account_id
                  left join Echo_Core_Prod..individuals i  on i.id = ca.salesman_id
                  where ca.dateOpened >@ReportDateStart
                  and jobStatus in (7,10)
                  and j.jobDate between ca.dateOpened and @ReportDateEnd
                  order by dateOpen, jobDate desc
                  ")


odbcClose(odbcChannel)



odbcChannel <- odbcConnect('Zeacom', uid ='snapshot', pwd='Z3ac0m1234')
calls  <- sqlQuery( odbcChannel, 
                    "
                    declare @yesterdayFROM datetime
                    declare @yesterdayTO datetime
                    set @yesterdayTO =getdate()
                    set @yesterdayFROM = DATEADD(DAY, -90, CONVERT(CHAR(10), getdate(), 111))
                    
                    select ac.CLID,ac.Exno,ac.Type,ac.Date,n.FirstName,n.LastName,ac.TalkTime,datepart(ISO_WEEK,ac.Date) + 2 'CallWeek',datepart(YY,ac.Date)  'CallYear'
                    from ZeacomConfig..pn_audit_calls ac
                    left join ZeacomConfig..pn_numbers pn on pn.number = ac.Exno
                    left join ZeacomConfig..names n on n.UniqueID = pn.NameID
                    
                    where ac.Date is not null
                    and ac.Date between @yesterdayFROM and @yesterdayTO
                    --  and ac.Resolution in ('Q','A')
                    ")

odbcClose(odbcChannel)

data$count<-1

dataBack<-data

data<-dataBack
#data[data$JobWeek <=27,"JobYear"]<-2017

match <- c('2200', '6443', 'G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9.1', 'G9.5','G10', 'GTC888', 'G50', 'G51', 'G5555', '7002', 'LONGTC1387','LHR Cash','LHR','LONGTCLHR-CREDIT')

data <- data[!(data$number %in% match),]



data[(data$JobWeek>=53 & data$JobMonth ==12 & data$JobYear ==2016),"JobYear"]<-2017
data[(data$JobWeek==54 & data$JobYear ==2017),"JobWeek"]<-2
data[(data$JobWeek==53 & data$JobYear ==2017),"JobWeek"]<-1


data[(data$JobWeek==55 & data$JobMonth ==1 & data$JobYear ==2016),"JobWeek"]<-2
data[(data$JobWeek>=54 & data$JobMonth ==12 & data$JobYear ==2015),"JobYear"]<-2016
data[(data$JobWeek==54 & data$JobMonth ==12 & data$JobYear ==2016),"JobWeek"]<-1
data[(data$JobWeek==55 & data$JobMonth ==12 & data$JobYear ==2016),"JobWeek"]<-2





PersonWeekly<-group_by(data,JobYear,JobWeek,fullName) %>% summarise(totalnetprice=sum(totalnetprice,na.rm=TRUE),
                                                                    totalcharge= sum(totalCharge,na.rm=TRUE),
                                                                    jobs=sum(count,na.rm=TRUE))



AccountWeekly<-group_by(data,JobYear,JobWeek,fullName,name,number) %>% summarise(totalnetprice=sum(totalnetprice,na.rm=TRUE),
                                                                                 totalcharge= sum(totalCharge,na.rm=TRUE),
                                                                                 jobs=sum(count,na.rm=TRUE))
PureWeekly<-group_by(data,JobYear,JobWeek) %>% summarise(totalnetprice=sum(totalnetprice,na.rm=TRUE),
                                                                    totalcharge= sum(totalCharge,na.rm=TRUE),
                                                                    jobs=sum(count,na.rm=TRUE))







#calls work
callsback<-calls
calls<-callsback
calls$CLID  <- as.character(calls$CLID)
calls$Exno  <- as.character(calls$Exno)
calls$Type  <- as.character(calls$Type)

calls[is.na(calls$CLID),"CLID"]<-"0"
calls[is.na(calls$Exno),"Exno"]<-"0"
calls[is.na(calls$Type),"Type"]<-0
typeof(calls$Type)


cc<-c("354","417","209")
cc<-c("354","417","419","448","9012","386","209", "385")
calls <- calls[calls$Exno %in% cc,]

calls <- calls[calls$Type == "O",]
calls$Count<-1



#Put a cap on talk time at 20 mins
calls$TalkTime[calls$TalkTime>=1200]<-1200

calls$fullname <-paste(calls$FirstName,calls$LastName,sep=" ")

WeeklyCalls<-as.data.frame(group_by(calls,fullname,CallWeek,CallYear) %>%  summarise(
  Calls=sum(Count), TalkTime = sum(TalkTime,na.rm=TRUE)))





# sales data section


#remove NA's AC
salesdata[is.na(salesdata$LastJob),"LastJob"]<-"1997-01-01"
salesdata[is.na(salesdata)]<-0
#---------------------------------
#set date for column names
#x<-as.POSIXlt(Sys.Date())
#w<-x
#w$mon<-w$mon-2

#month1<-format(w,"%b %y")
#month1<-paste("Month starting",month1,sep=" ")
#w$mon<-w$mon+1
#month2<-format(w,"%b %y")
#month2<-paste("Month starting",month2,sep=" ")
#w$mon<-w$mon+1
#month3<-format(w,"%b %y")


#---------------------------------



#final report and rename columns with dynamic month names
final<-salesdata[c("ReportDateStart",
                     "OpeningDate",
                     "SalesPerson",
                     "Parent/AccountName",
                     "Parent/AccountNumber",
                     "AccountName",
                     "contact",
                     "TotalJobsReportPeriod",
                     "TotalSpendReportPeriod",
                     "FutureJobs",
                     "LastJob")
                   ]

SalesPeople<-c("Niche Sullivan",
               "Claire Thackeray",
               "Kennifer Patric",
               "Mark Taylor",
               "Pierre Netty",
               "Ronak Nayee",
               "Joshua King")

Niche<-final[final$SalesPerson == "Niche Sullivan",]
Claire<-final[final$SalesPerson == "Claire Thackeray",]
Mark<-final[final$SalesPerson == "Mark Taylor",]
Ronak<-final[final$SalesPerson == "Ronak Nayee",]
Joshua<-final[final$SalesPerson == "Joshua King",]


#append calls
CallsOld<-read.xlsx(filename, sheetName = "Calls",header = TRUE, rowIndex=NULL)
CallsOld<-CallsOld[,2:6]
WeeklyCalls<-unique(rbind(CallsOld,WeeklyCalls))

# fix variable type in weeklyCalls
WeeklyCalls$Calls = as.numeric(WeeklyCalls$Calls)
WeeklyCalls$TalkTime = as.numeric(WeeklyCalls$TalkTime)


#load into workbook
wb<-loadWorkbook(filename)
sheets<-getSheets(wb)


removeSheet(wb,sheetName  = "PersonWeekly")
removeSheet(wb,sheetName ="AccountWeekly")
removeSheet(wb,sheetName ="NS_All")
removeSheet(wb,sheetName ="Calls")

sheets<-getSheets(wb)

AccountSheet<-createSheet(wb,sheetName = "AccountWeekly")
PersonSheet<-createSheet(wb,sheetName = "PersonWeekly")
CallsSheet<-createSheet(wb,sheetName = "Calls")
#NicheSheet<-createSheet(wb,sheetName ="NS_Niche")
#ClaireSheet<-createSheet(wb,sheetName ="NS_Claire")
#MarkSheet<-createSheet(wb,sheetName ="NS_Mark")
#RonakSheet<-createSheet(wb,sheetName ="NS_Ronak")
AllSheet<-createSheet(wb,sheetName ="NS_All")


addDataFrame(as.data.frame(PersonWeekly),PersonSheet)
addDataFrame(as.data.frame(AccountWeekly),AccountSheet)
addDataFrame(WeeklyCalls,CallsSheet)
addDataFrame(as.data.frame(final),AllSheet)

saveWorkbook(wb, filename)
string1<-paste("The sales weekly report is at: 
               ",filename,sep="")



#Now send it to Lee
library(RDCOMClient)
# Send mail for 3D
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["subject"]] = 'Weekly Sales Report'
outMail[["To"]] = "daria.alekseeva@greentomatocars.com;sean.sauter@greentomatocars.com;Yinka.Ogunniyi@greentomatocars.com;muaaz.sarfaraz@greentomatocars.com"
#outMail[["To"]] = "antony.carolan@greentomatocars.com"
outMail[["body"]] =string1 

outMail$Send()
rm(list = c("OutApp","outMail"))
