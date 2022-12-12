# Steve's Stock Analysis
## Steve's first clients, his parents, have chosen to use his new financing degree for themselves. Steve's parents want to invest their money into green energy with DAQ0 New Energy Corporation. Steve wants to do a comparison of a few different green energy options in order to create some diversificaiton for his parent's portfolio. Calculating total daily volume with VBA, Steve is able to visualize the activity of the stocks traded. His findings report that DAQ0 may not be the best option for Steve's parents. Steve continues on to calculate multiple options. 
### The results show that in 2018, ENPH and RUN performed the best on the market overall with an 81.9% (ENPH) return and 84% (RUN) return 2017 with JKS and DQ taking the biggest drop from 2017.
![2018](https://user-images.githubusercontent.com/117100491/207067237-90a72904-15d2-42fa-ad08-3f718551ad27.PNG)

### 2017 sported DQ at 199.4% (DQ) return, 184.5% (SEDG) return, and 129.5% return (ENPH). *ENPH has run high returns for both 2017 and 2018.*

![2017](https://user-images.githubusercontent.com/117100491/207066975-fa271076-3d9d-4a64-ae9b-3df61a6bc617.PNG)

*(Image difference due to using PC during work and Mac at home)*
 
 ### Using this code to format, we have a clear picture analysis of the stock findings.
 'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
## Summary
### Refactoring code resulted in breaking and reorganizing the coding script to include the 2017 results as well as the issues with replicating the code to make sure that it showed both 2017 and 2018 when the inputBox came up. Refactoring the code was easier with a comment template placed over the script so I could identify the order of the coding and make the appropriate commands to reinegrate the data from the 2017 and 2018 stock data sheets.  
