# Excel Sales Dashboard
This project is to demonstrate my Data Visualization skills in Excel.

![](/assets/images/ExcelDashboard.gif)

### Step-by-step process of implementing excel dashboard
1. Review and explore  raw data in excel sheets.
   ![](/assets/images/RawData.png)
   
2. Gather customer data into ‘Orders’ worksheet using XLOOKUP.
   - Formula used to get Customer name column data from customers sheet =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)
   - Formula used to get Email address column data from customers sheet =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001))
   - Formula used to get Country column data from customers sheet =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)
4.	Use INDEX MATCH to gather columns data – ‘Coffee Type’, ’Roast Type’, ’Size’, ’Unit Price’ into ‘Orders’ worksheet from Products sheets =INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
5.	Calculate Sales column data using formula =L2*E2
6.	Create new column ‘Coffee Type Name’ and generate full name using IF function =IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))
7.	Create new column ‘Roast Type Name’ and generate full name using IF function =IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))
8.	Add Loyalty Card column data into ‘Orders Sheet’ using formula =XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,"",0)
9.	Format data column to ‘dd-mmm-yyyy’.
10.	Format Size column to ‘0.0 kg’
11.	Add currency signs to Unit Price and Sales column.
12.	Create pivot table from Orders Table
13.	Insert line chart, timeline and slicers
14.	Format all visual components.
15.	Move all charts and slicers to new sheet to create dashboard.
16.	Connect all slicers to remaining charts.
