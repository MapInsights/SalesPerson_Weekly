taskkill /IM Outlook.exe /f
start "" "%ProgramFiles(x86)%\Microsoft Office\Office14\outlook.exe"

@echo on@echo on
"C:\Program Files\R\R-3.2.3\bin\R.exe" CMD BATCH C:\Programs\gtc_tasks\SalesPerson_Weekly\SalesPerson_Weekly.R