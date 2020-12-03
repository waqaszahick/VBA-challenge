# VBA-challenge
Homework2-VBA

This is my second homework of ‘Excel VBA’ (Visual Basic for Applications) challenge. Two ‘Microsoft Excel’ files: ‘alphabetical_test.xlsx’ and ‘multiple_year_stock_data.xlsx’ were given to me to work on. One file (alphabetical_test.xlsx) is way smaller than the other (multiple_year_stock_data.xlsx). Each file has data with multiple worksheets. The aim of this challenge was to develop a generalized (VBA) script that would run on any and / or all the worksheets of both the ‘Microsoft Excel’ files just with a click of button.
I used ‘alphabetical_test.xlsx’ file to develop the generic script because the data and the file size of this file was heaps smaller (almost 14MB) as compared to the other, almost 96MB file i.e. multiple_year_stock_data.xlsx. It allowed me to test and debug the script in less than 3-5 minutes on ‘alphabetical_test.xlsx’.
I used ‘for’ and ‘if’ loops in my script, that ran through all the available data in order to calculate the following:
* The ‘Ticker’ symbols
* Yearly change i.e. from opening price at the beginning of year to the closing price at the end of the year
* Percentage change i.e. opening price at the beginning of year to the closing price at the end of the year
* Total volume of the stock
* Greatest % increase ‘Ticket’ symbol
* Greatest % decrease ‘Ticket’ symbol
* Greatest total volume ‘Ticket’ symbol
* Greatest % increase ‘Value’
* Greatest % decrease ‘Value’
* Greatest total volume ‘Value’
Apart from the above-mentioned tasks performed, I applied conditional formatting that highlighted the negative values in ‘red’ colour and the positive values in ‘green’ colour.
Also, I uploaded three screenshots of my results of three respective worksheets / years on the ‘Multi Year Stock Data’ file i.e. ‘multiple_year_stock_data.xlsx’, along with a separate ‘VBA Script’ file i.e. ‘VBA-challenge.vbs’.
