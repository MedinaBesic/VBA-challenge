# VBA-challenge
Module 2 Challenge 
Resources used for this module:

**Line of code:**
ws.Range("A1").CurrentRegion.Sort key1:=ws.Range("A1"), order1:=xlAscending, Header:=xlYes
**Resource(s):**
https://www.youtube.com/watch?v=YzAkrx3nvGM
https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08

**Line of code:**
'Check to see if we are still in the same ticker symbol, if not
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Add the ticker volumes together from the volumes listed in column 7
tickervolume = tickervolume + ws.Cells(i, 7).Value

'Place the ticker in column I
ws.Range("I" & summary_ticker_row).Value = ticker

'Assign a value to the close_price variable
close_price = ws.Cells(i, 6).Value

'Assign a value to the open_price variable
open_price = ws.Cells(i, 3).Value

'Assign a value to yearly_change
yearly_change = (close_price) - (open_price)

'Calculate the price change in column J
ws.Range("J" & summary_ticker_row).Value = yearly_change
**Resource(s):**
https://github.com/DataTell/VBA-Challenge/blob/master/VBAStocks/VBAStocksScript.bas
https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Moderate_DS.vbs

**Line of code:**
open_price = ws.Cells(i, 3).Value
**Resoure(s):**
https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08

**Line of code:**
tickervolume = tickervolume + ws.Cells(i, 7).Value
**Resoure(s):**
https://chat.openai.com/share/73c834d6-babf-49ae-986b-dd47ef30fd08

**Line of code:**
If yearly_change > 0 Then
**Resource(s):**
https://chat.openai.com/share/8c39a28f-f753-4363-9489-5683afc35f76
