# stock-analysis
### Project Overview
---                                                          
Steve, graduated in fall of 2021, with a bachelor’s degree in finance. Following the 
virtual ceremony, his parents sat him down for a chat. Socially, his parents live a “green” lifestyle, and in their social spheres they actively promote alternative energy (ATM: solar, wind, geothermal, hydroelectric, tidal, biomass and hydrogen). Comfortable now, Mom and Dad conveyed spirited intentions on expanding their support, economically. 

An intriguing new venture; a challenge such as this one entails risk - requiring collaborative effort, and prospective investment to counter. Luckily, the projects timing fittingly coincided with the fruition of Steve’s 4-year-long academic pursuit - attaining financial prowess, and acumen. In a gesture of good-faith, Mom and Dad - as his inaugural clients – decided to employ Steve. 

His new clients had initially  decided to invest all their money into, Daqo New Energy Corp. DQ (NYSE) - a manufacturer of silicon wafers, designed for solar panels. Steve’s first assignment: to implement an actionable, value-driven investment strategy in support of his client’s investment interests in alternative energy(any energy source that does not use fossil fuels - coal, gasoline and natural gas). In turn, Steve pulled, and wrangled sets of DQ stock data compared to other companies - using Excel to organize the data, and perform analysis; however, Steve needed help, and commissioned this project to explore alternative energy stock performance by analyzing financial data using VBA (Visual Basic for Applications).

---

### Deconstructing the Objective 

One way to perform this data analysis would be to go through all of Steve's stock data manually and use Excel formulas for calculations. But with Visual Basic for Applications,"VBA," we can write code that will automate these analyses for us. Often used in the finance industry, VBA provides essentially infinite extensibility to Excel. Using code to automate tasks decreases the chance of errors and reduces the time needed to run analyses, especially if they need to be done repeatedly. 

Code from the original script: 

    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        'Loop over the data
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            'totalVolume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'startingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'endingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        'Output results
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i

Refactored code:

    Dim tickerIndex As Integer

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Initialize ticker volumes to zero
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next tickerIndex
    're-initialize tickerIndex to zero before looping over all rows
    tickerIndex = 0
        
    '2b) loop over all the rows
    For i = 2 To RowCount
         
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            'starting price for the current ticker
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            'ending price for the current ticker
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3d) Increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

### Stock Performance Comparison Between 2017 and 2018
--- 
- Analysis & execution times for both years with the original VBA script:
<img width="574" alt="Screen Shot 2021-08-12 at 3 01 24 PM" src="https://user-images.githubusercontent.com/82069038/129255690-27c1f86e-c4ff-41f3-b87d-473ce7a9c8a9.png">
<img width="569" alt="Screen Shot 2021-08-12 at 3 01 46 PM" src="https://user-images.githubusercontent.com/82069038/129255691-cbbb1c87-7d47-40fb-93d8-78606d9d0c6e.png">
- Analysis & execution times for both years with the refactored VBA script:
<img width="566" alt="Screen Shot 2021-08-12 at 3 06 37 PM" src="https://user-images.githubusercontent.com/82069038/129256313-5597b74c-153b-4137-846c-7501713fa445.png">
<img width="572" alt="Screen Shot 2021-08-12 at 3 07 12 PM" src="https://user-images.githubusercontent.com/82069038/129256318-4d981ac3-6f3f-40fe-8e4b-fccac314fa26.png">

### Summary
--- 
Refactored code pros: code runs faster, cleaner, and more efficient.
Refactored code cons: no additional functionality, increased development time, and continuity TAT for further development.
Takeaway: Recoding the refactored script allows us to analyze any set of stocks.
