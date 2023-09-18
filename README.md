# VBA-challenge
# TITLE    STOCK MARKET DATA ANALYSIS     
# DESCRIPTION 
This is a generated stock market data from the year 2018 to 2020. Need to create a script that will loop through each year of stock data and grab 1) the ticker symbol; 2) yearly change from the opening price to the closing price; 3) percent change from the beginning of a given year to the closing price at the end of that year; 4) the total stock volume of the stock.

Additionally, the solution will include the “Greatest % increase”, “Greatest % decrease” and “Greatest total volume ”.

# CHALLENGE

---Need to make the adjustment to the script that will make the macro run on every worksheet just by running it once.
---To find out the same Ticker by using the 
       If Ticker <> ws.Cells(i + 1, 1).Value 
                
---The For Loop for repeating actions. And the trick part is that once the end value of the loop is reached, it should change another cell’s value for the next new loop.
--- During the looping part, need to define the first” open price”  for the stock’s  “yearly change .“
     OpenPrice = ws.Cells(2, 3).Value
---Also need to define the “total volume=0” for the loop that calculate the  every  stock’s “total volume”. 
        Total_Volume = 0
----Out of the loop, the “open price” would be the row of “i+1”, and the” Total volume“ would be the value of cells(“i+1,3).          
                OpenPrice = ws.Cells(i + 1, 3).Value
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
