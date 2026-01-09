# Excel Sales Dashboard
This project is to demonstrate my Data Visualization skills in Excel.

![](/assets/images/ExcelDashboard.gif)

### Step-by-step process of implementing excel dashboard
1. Review and explore  raw data in excel sheets.
   ![](/assets/images/RawData.png)
   
2. Gather customer data into ‘Orders’ worksheet using XLOOKUP.
   - Formula used to get Customer name column data from customers sheet =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)
     ![](/assets/images/getCustomerName.png)
     
   - Formula used to get Email address column data from customers sheet =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001))
     ![](/assets/images/getCustomerEmail.png)
     
   - Formula used to get Country column data from customers sheet =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)
     ![](/assets/images/getCustomerCountry.png)
     
3.	Use INDEX MATCH to gather columns data – ‘Coffee Type’, ’Roast Type’, ’Size’, ’Unit Price’ into ‘Orders’ worksheet from Products sheets =INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
   ![](/assets/images/IndexMatch.png)
     
4.	Calculate Sales column data using formula =L2*E2
   ![](/assets/images/getSales.png)
     
5.	Create new column ‘Coffee Type Name’ and generate full name using IF function =IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))
   ![](/assets/images/getCoffeeName.png)
     
6. Create new column ‘Roast Type Name’ and generate full name using IF function =IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))
    ![](/assets/images/getRoastType.png)
     
7.  Add Loyalty Card column data into ‘Orders Sheet’ using formula =XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,"",0)
   ![](/assets/images/getLoyaltyCard.png)
     
8.	Format data column to ‘dd-mmm-yyyy’.
    ![](/assets/images/formatDate.png)
     
9.	Format Size column to ‘0.0 kg’ and add currency signs to 'Unit Price' and 'Sales' column.
    ![](/assets/images/formatSize.png)

10. Covert data from 'orders' sheet to a table and name it 'Orders'. Create pivot table from 'Orders' table.
   ![](/assets/images/pivotSales.png)
   ![](/assets/images/pivotTopFive.png)
   ![](/assets/images/pivotCountry.png)
   	
11. Insert line chart, timeline and slicers and format all visual components.
    ![](/assets/images/formatslicer.png)
     
12. Move all charts and slicers to new sheet and connect all slicers to remaining charts.
    ![](/assets/images/slicerConnections.png)

13. Arrange the visual components positioning to create final look.
    ![](/assets/images/FinalDashboard.png)
