Sub Stock_Data_Analysis():
''Christ_is_King ICXC''
For Each ws In Worksheets

    ''make the selected worksheet active
    Worksheets(ws.Name).Activate
    
    
    ''Set Title for Macro Calculations
    Range("J1").Value = "Ticker Symbol"
    Range("K1").Value = "Percentage Change (%)"
    Range("L1").Value = "Yearly Change ($)"
    Range("M1").Value = "Total Stock Volume"
    
    ''Set an initial variable for holding Ticker Name
    Dim Ticker_Name As String
    
    ''Set an initial variable for holding the total per Ticker
    
    Dim Ticker_Total_Volume As Double
            Ticker_Total_Volume = 0
    
    ''Set an initial variable for holding the % Change per Ticker
      Dim Percentage_Change As Double
            Percentage_Change = 0
         
    ''Set an initial variable for holding the Yearly Change per Ticker
        Dim Yearly_Change As Variant
            Yearly_Change = 0

    ''Set an initial variable for Open&Close Prices
        Dim First_Open_Price As Double
            First_Open_Price = 0
            Dim Last_Close_Price As Double
            Last_Close_Price = 0
    
    ''Find the last row of the set
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    
    ''Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
    
        Ticker_Name = Cells(2, 1).Value
        First_Open_Price = Cells(2, 3).Value
        Current_Stock = Cells(2, 1).Value
    
    ''Begin Loop for Subroutine
    For i = 2 To LastRow 'Range("G1").End(xlDown).Row
    
    ''Get the first open price value for the loop
    If Cells(i, 1).Value <> Current_Stock Then
   
 '---------------------------------------------------------------
        '''Print Column Headers for analysis
 
    ''Print the Ticker_Name in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Name


    ''Print the Percentage Change_Name in the Summary Table
      Range("K" & Summary_Table_Row).Value = Percentage_Change
    
    ''Print the Yearly Change Name in the Summary Table
      Range("L" & Summary_Table_Row).Value = Yearly_Change
      
      
    ''Print the Volume Amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = Ticker_Total_Volume
     
'---------------------------------------------------------------

       Current_Stock = Cells(i, 1).Value
    First_Open_Price = Cells(i, 3).Value
    
    ''Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      
        '''Reset all numerical variables
    ''Reset the Ticker Total Volume
        Ticker_Total_Volume = 0
      
    ''Reset the Yearly Change
        Yearly_Change = 0
        
    ''Reset the Percentage Change
      Percentage_Change = 0
    
    
    ''Check if we are still have the same Ticker name, if it is not...

    ElseIf Cells(i, 1).Value = Current_Stock Then

        ''Set the Ticker name
        Ticker_Name = Cells(i, 1).Value
      
        ''Set the Last_Close_Price
        Last_Close_Price = Cells(i, 6).Value
        
        ''Add to the Yearly Change
        Yearly_Change = Last_Close_Price - First_Open_Price
        

        ''Add to the Ticker Total Volume
        Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
              
        ''Set the Percentage Change and avoid divide by zero error
        If First_Open_Price <> 0 Then
        Percentage_Change = ((Last_Close_Price - First_Open_Price) / First_Open_Price)
        Else:
        Percentage_Change = 0
        End If
             
        ''If the cell immediately following a row is the same Ticker
    Else
        ''Add to the Ticker Total
      Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
    
    End If
       Next i
    
        ''Set up the Conditional Formatting
            
    Range("L2:L4000").Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = 1
            End With
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbGreen
            End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
         Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
        .Color = 1
            End With
            
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 500
            End With
    Selection.FormatConditions(1).StopIfTrue = False


    ''Change Header Style
        Range("J1:M1").Style = "Check Cell"
        
    ''Change Number style format of Yearly Change and Percentage Change
        Range("K2:K4000").NumberFormat = "0.00%"
        Range("L2:L4000").NumberFormat = "0.00"
        
    
    Next ws
    
    MsgBox ("Thanks for Grading my submission :)")
  
End Sub
