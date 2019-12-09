# Excel VBA

## Project: Stock Market Data

This project has the goal to put in practice my VBA skills. For this reason, I gathered a dataset from a stock market. And from this point, I developed three macros to iterate within all worksheets. Moreover, the results you can check below (Result header).

## Result
### Macro Calculate Vol_Easy 

    - It will loop through one year of stock data for each run and return the total volume each stock (Column J) had over that year
    - Display the ticker symbol (Column I) to coincide with the total stock volume

![calculate](/Screenshots/calculate.png)

### VBA Calculate Moderate
    - The ticker symbol (Column I)
    - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year (Column J)
    - The percent change from opening price at the beginning of a given year to the closing price at the end of that year (Column K)
    - The total stock volume of the stock (Column L)
    
![calculate_Moderate](/Screenshots/calculate_Moderate.png)

### VBA Calculate Hard
    - This includes everything from the macro calculate_Moderate
    - Also, it returns the stock with the "Greatest % increase" (Column N), "Greatest % Decrease" (Column O) and "Greatest total volume" (Column P).
    
![calculate_Hard](/Screenshots/calculate_Hard.png)


## VBA Code Snippet

![codesnippet](/Screenshots/codesnippet.png)


