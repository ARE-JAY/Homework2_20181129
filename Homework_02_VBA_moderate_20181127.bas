Sub ticker_volume()
	
	Dim ticker As String
	Dim trading_volume As Double
	Dim ticker_count As Integer
	Dim r As Long
	Dim c As Long
	Dim ws As Worksheet
	Dim ws1 As Worksheet
	
	
	
	Set ws1 = ActiveSheet
	
	For Each ws In ThisWorkbook.Worksheets
		ws.Activate
		ticker_count = 0
		ticker = ""
		r = 1
		c = 1
		trading_volume = 0
		year0 = 21990101
		maxDate = 0
		
'Create column headers "Ticker" and "Trading Volume"
		ws.Cells(1, 9).Value = "Ticker"
		ws.Cells(1, 12).Value = "Trading Volume"
		
'While the values and column A are not empty/blank
		While IsEmpty(ws.Cells(r + 1, 1).Value) = False
			
'Enter each unique ticker symbol in column I and enter the aggregated trading volume column J
			If ws.Cells(r + 1, 1).Value = ticker Then
				trading_volume = trading_volume + ws.Cells(r + 1, 7).Value
				r = r + 1
			Else
				year0 = ws.Cells(r + 1, 3).Value
				trading_volume = 0
				ticker_count = ticker_count + 1
				ticker = ws.Cells(r + 1, 1).Value
				ws.Cells(ticker_count + 1, 9).Value = ticker
				trading_volume = trading_volume + ws.Cells(r + 1, 7).Value
				ws.Cells(ticker_count + 1, 12).Value = trading_volume
				ws.Cells(ticker_count + 1, 12).NumberFormat = "#,##0"
				r = r + 1
				
			End If
			ws.Cells(ticker_count + 1, 12).Value = trading_volume
			ws.Cells(ticker_count + 1, 12).NumberFormat = "#,##0"
		Wend
		Next ws
		ws1.Activate
		
		
	End Sub
	
	
	Sub yearly_change()
		
		Dim ticker As String
		Dim ticker_count As Integer
		Dim ws As Worksheet
		Dim ws1 As Worksheet
		Dim minDate As Long 
		Dim maxDate As Long
		Dim minPrice As Currency
		Dim maxPrice As Currency
		
		
		
'Set worksheet variable to the active sheet		
		Set ws1 = ActiveSheet

'Loop thu each wrksheet using a ForLoop()		
		For Each ws In ThisWorkbook.Worksheets
			ws.Activate
			ticker_count = 1
			ticker = ""
			r = 1
			c = 1
			minDate = 21990101
			maxDate = 0
			minPrice = 0
			maxPrice = 0
			
			
'Create column headers "Yearly Change" and "Percent Change"
			ws.Cells(1, 10).Value = "Yearly Change"
			ws.Cells(1, 11).Value = "Percent Change"
'Seed first ticker value
			ticker = ws.Cells(r + 1, 1).Value
			
'While the values and column A are not empty/blank
			While IsEmpty(ws.Cells(r + 1, 1).Value) = False
				
'Loop thru each row until you to capture price on the minDate and maxDate 
				If ws.Cells(r + 1, 1).Value <> ticker Then
					If minPrice <> 0 Then
						ws.Cells(ticker_count + 1, 11).Value = (maxPrice - minPrice) / minPrice
						Else: ws.Cells(ticker_count + 1, 11).Value = 0
					End If
					ws.Cells(ticker_count + 1, 11).NumberFormat = "0.00%"
					ticker_count = ticker_count + 1
					minDate = 21990101
					maxDate = 0
					minPrice = 0#
					maxPrice = 0#
					ticker = ws.Cells(r + 1, 1).Value
					
					r = r + 1
					
'Conditional statments used to identify min and max dates   					
				Else
					If maxDate < ws.Cells(r + 1, 2).Value Then
						maxDate = ws.Cells(r + 1, 2).Value
						maxPrice = ws.Cells(r + 1, 6).Value
					End If
					If minDate > ws.Cells(r + 1, 2).Value Then
						minDate = ws.Cells(r + 1, 2).Value
						minPrice = ws.Cells(r + 1, 6).Value
					End If
					
					r = r + 1
					
				End If
				ws.Cells(ticker_count + 1, 10).Value = (maxPrice - minPrice)
				If ws.Cells(ticker_count + 1, 10).Value >= 0 Then
					ws.Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
					Else: ws.Cells(ticker_count + 1, 10).Interior.ColorIndex = 3
				End If
				If minPrice <> 0 Then
					ws.Cells(ticker_count + 1, 11).Value = (maxPrice - minPrice) / minPrice
					Else: ws.Cells(ticker_count + 1, 11).Value = 0
				End If
				ws.Cells(ticker_count + 1, 11).NumberFormat = "0.00%"
				
			Wend
			Next ws
			ws1.Activate
			
			
		End Sub
		
		
		
		
		
		
		
		
