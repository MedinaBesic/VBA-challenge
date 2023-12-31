Sub stocks()

'Set worksheet variable
Dim ws As Worksheet

'Loop through each row
For Each ws In Worksheets

'Sort the entire ws in alphabetical order by ticker
'I referenced this youtube video to be able to do this: https://www.youtube.com/watch?v=YzAkrx3nvGM
'Chat gpt helped me with an error I was getting (I forgot ws): https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08
ws.Range("A1").CurrentRegion.Sort key1:=ws.Range("A1"), order1:=xlAscending, Header:=xlYes

'Create column headers for Ticker, Yearly Change, Percent Change, and Total Stock Volume for each worksheet
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Create columns for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Coming back to add Greatest % Increase, Greatest % Decrease, Greatest Total Volume and a new Ticker and Value column headers in each ws
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Set an initial variable to hold the ticker
Dim ticker As String

'Set an initial variable to represent total volume
Dim tickervolume As Double
tickervolume = 0

'Set an initial variable to keep track of the location of the ticker in the summary table
'Why is this 2? In class we started this at 2.
Dim summary_ticker_row As Double
summary_ticker_row = 2

'Create additional variables to help clean up the code such as open price, close price, and yearly change. Replace the existing code with variables. This will make percent change look neater.

'Set a variable for open price
Dim open_price As Double

'Set a variable for close price
Dim close_price As Double

'Set variable for yearly change
Dim yearly_change As Double

'Set a variable for percent change
Dim percent_change As Double

'Determine the last row for each worksheet
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through ticker
For i = 2 To lastrow

'Assign a value to ticker, this represents the ticker symbol, this is given in column 1 through rows i
ticker = ws.Cells(i, 1).Value

'Check to see if we are still in the same ticker symbol, if not
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Add the ticker volumes together from the volumes listed in column 7 for the same ticker
'The next line of code is from: https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Moderate_DS.vbs
tickervolume = tickervolume + ws.Cells(i, 7).Value

'Place the ticker in column I
'This was also shown in the in-class
ws.Range("I" & summary_ticker_row).Value = ticker

'Assign a value to the close_price variable
'Office hours 6/22: We can clean up code by putting parts of a function in a variable: Example: close_price is used instead of ws.Cells(i, 6).Value when calculating yearly change/ percent change
close_price = ws.Cells(i, 6).Value

'Assign a value to the open_price variable
''Office hours 6/22: We can clean up code by putting parts of a function in a variable: Example: open_price is used instead of ws.Cells(i, 3).Value when calculating yearly change/ percent change
open_price = ws.Cells(i, 3).Value

'Assign a value to yearly_change
'The following line of code was provided in office hours on 6/22
yearly_change = (close_price) - (open_price)

'Calculate the price change in column J
ws.Range("J" & summary_ticker_row).Value = yearly_change

'Add conditional formatting to column J
'Initially, my entire row J was red
'Chat cpt helped fix pin-point where the issue was: https://chat.openai.com/share/8c39a28f-f753-4363-9489-5683afc35f76
If yearly_change > 0 Then

ws.Range("J" & summary_ticker_row).Interior.ColorIndex = 4

Else

ws.Range("J" & summary_ticker_row).Interior.ColorIndex = 3

End If

'Calculate the total volume in column L
ws.Range("L" & summary_ticker_row).Value = tickervolume

'Create an if/then statement to check the divisibility of open_price when calculating the percent change
'The next two lines of code are from: https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Moderate_DS.vbs
If open_price = 0 Then

percent_change = 0

Else

'The following line of code was provided in office hours on 6/22
percent_change = ((yearly_change) / (open_price)) * 100

End If

'Calculate the yearly change in column K
ws.Range("K" & summary_ticker_row).Value = percent_change

'Reset ticker counter
summary_ticker_row = summary_ticker_row + 1

'Reset total volume
tickervolume = 0

'Reset the open_price
'Chat gpt helped me with an error I was getting (I forgot ws): https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08
open_price = ws.Cells(i + 1, 3)

Else

'Add the tickervolume
'Chat gpt helped me with an error I was getting (I forgot ws): https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08
tickervolume = tickervolume + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub