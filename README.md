# VBA-challenge
## Module 2 Challenge,coding with VBA in Excel

### First Code 
Create a script that loops through all the stocks for one year and outputs the following information:
•	The ticker symbol
•	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
•	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
### * [1-Algoritme ]
 ##### Retrieval of Data 
•	The script loops through one year of stock data and reads/ stores all of the following values from each row:
  ```
  ' Loop through all ticker purchases
    For i = 2 To lastrow
```
 ###### o	ticker symbol 
 ###### o	volume of stock 
 ###### o	open price 
 ###### o close price 
 ##### Column Creation 
 •	On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
 ###### o	ticker symbol 
 ###### o	total stock volume 
 ###### o	yearly change ($) 	
 ###### o percent change 

```          
  ' Check if we are still within the same ticker, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

  ' Set the new value of ticker and it s close_value
      Ticker = Cells(i, 1).Value
      close_value = Cells(i, 6).Value
      
  ' Add  the Total volume
      Total_volume = Total_volume + Cells(i, 7).Value
  ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
      Yearly_change = close_value - open_value
  ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
     percentage_change = Yearly_change / open_value
```

```
  ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker
  ' Print the yearly change in the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_change
  ' Print the percentage change in the Summary Table
      Range("K" & Summary_Table_Row).Value = percentage_change
  ' Print the Total_volume in the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_volume
```

 ##### Conditional Formatting 
•	Conditional formatting is applied correctly and appropriately to the yearly change column 
```            
          If Yearly_change < 0 Then
                 Range("J" & Summary_Table_Row).Interior.ColorIndex = 3  ' Red
                 Else
                 Range("J" & Summary_Table_Row).Interior.ColorIndex = 4  ' Green

          End If
``` 
      
•	Conditional formatting is applied correctly and appropriately to the percent change column
 ```
If percentage_change < 0 Then
                 Range("K" & Summary_Table_Row).Interior.ColorIndex = 3  ' Red
                 Else
                 Range("K" & Summary_Table_Row).Interior.ColorIndex = 4  ' Green

 End If
```
#### 2-Code "Module1
![module1]([module1 - Copy.pdf](https://github.com/fahr-khadija/VBA-challenge/blob/main/module1%20-%20Copy.pdf))
#### 3-Execution Module1
•	The total stock volume of the stock. The result should match the following image:
-![image](https://github.com/fahr-khadija/VBA-challenge/blob/main/year2018.jpg))
-![image](https://github.com/fahr-khadija/VBA-challenge/blob/main/year2019.jpg))
-![image](https://github.com/fahr-khadija/VBA-challenge/blob/main/year2020.jpg))
### Second Code 
•	Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
#### 1-Algoritme 
##### Calculated Values 
•	All three of the following values are calculated correctly and displayed in the output:
 ##### o	Greatest % Increase 
```
For j = 2 To lastrow
     '===================================Greatest % Increase
       If Cells(j + 1, 11).Value > Max Then
             ' Print the ticker and Max  in the Summary Table
             Max = Range("K" & j + 1).Value
             Tricker1 = Range("I" & j + 1).Value
             Range("P" & 2).Value = Max
             Range("O" & 2).Value = Tricker1
        Else
          Range("P" & 2).Value = Max
          Range("O" & 2).Value = Tricker1
     End If
```
##### o	Greatest % Decrease 
```
'===================================Greatest % Decrease
      If Cells(j + 1, 11).Value < Min Then
             ' Print the ticker and Min  in the Summary Table
              Min = Range("K" & j + 1).Value
              Tricker2 = Range("I" & j + 1).Value
              Range("P" & 3).Value = Min
              Range("O" & 3).Value = Tricker2
                        
          Else
          Range("P" & 3).Value = Min
          Range("O" & 3).Value = Tricker2
     End If
```
##### o	Greatest Total Volume 
```
'===================================Greatest Total Volume 
 If Cells(j + 1, 12).Value > Total_value Then
             ' Print the ticker and Total_value in the Summary Table
             Total_value = Range("L" & j + 1).Value
             Tricker3 = Range("I" & j + 1).Value
      Range("P" & 4).Value = Total_value
      Range("O" & 4).Value = Tricker3
                        
          Else
          Range("P" & 4).Value = Total_value
          Range("O" & 4).Value = Tricker3
     End If
```

##### Looping Across Worksheet 
•	The VBA script can run on all sheets successfully.
#### 2-Code "Module2
![module2]([module2 - Copy.pdf](https://github.com/fahr-khadija/VBA-challenge/blob/main/module2%20-%20Copy.pdf))
#### 3-Execution Module2
  
  



  
