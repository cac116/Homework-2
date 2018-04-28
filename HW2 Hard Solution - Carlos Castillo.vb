Sub FindMax()
    Dim PercentRnG, VolumeRnG As Range
    Dim Max_Percent_Change, Min_Percent_Change, Max_Volume As Double
    Dim MaxPercentRow, MinPercentRow, MaxVolumeRow  As Long
    
    For worksheet_index = 1 To 1 'worksheet_count

        Set PercentRnG = Worksheets(worksheet_index).Range("Q1", Range("Q" & Rows.Count).End(xlUp))
        Set VolumeRnG = Worksheets(worksheet_index).Range("K1", Range("K" & Rows.Count).End(xlUp))

        Max_Percent_Change = Worksheets(worksheet_index).Application.WorksheetFunction.Max(PercentRnG)
        Min_Percent_Change = Worksheets(worksheet_index).Application.WorksheetFunction.Min(PercentRnG)
        Max_Volume = Worksheets(worksheet_index).Application.WorksheetFunction.Max(VolumeRnG)


        MaxPercentRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Max_Percent_Change, PercentRnG, 0)
        MinPercentRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Min_Percent_Change, PercentRnG, 0)
        MaxVolumeRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Max_Volume, VolumeRnG, 0)


        Worksheets(worksheet_index).Range("T1").Value = "<ticker>"
        Worksheets(worksheet_index).Range("U1").Value = "<value>"

        Worksheets(worksheet_index).Range("S2").Value = "<greatest % increase>"
        Worksheets(worksheet_index).Range("T2").Value = Worksheets(worksheet_index).Range("M" & MaxPercentRow).Value 
        Worksheets(worksheet_index).Range("U2").Value = Worksheets(worksheet_index).Range("Q" & MaxPercentRow).Value
        Worksheets(worksheet_index).Range("U2").Style = "Percent"
        Worksheets(worksheet_index).Range("U2").NumberFormat = "0.00%"

        Worksheets(worksheet_index).Range("S3").Value = "<greatest % decrease>"
        Worksheets(worksheet_index).Range("T3").Value = Worksheets(worksheet_index).Range("M" & MinPercentRow).Value 
        Worksheets(worksheet_index).Range("U3").Value = Worksheets(worksheet_index).Range("Q" & MinPercentRow).Value
        Worksheets(worksheet_index).Range("U3").Style = "Percent"
        Worksheets(worksheet_index).Range("U3").NumberFormat = "0.00%"
        
        Worksheets(worksheet_index).Range("S4").Value = "<greatest total value>"
        Worksheets(worksheet_index).Range("T4").Value = Worksheets(worksheet_index).Range("M" & MaxVolumeRow).Value 
        Worksheets(worksheet_index).Range("U4").Value = Worksheets(worksheet_index).Range("K" & MaxVolumeRow).Value



        ' MsgBox ("Max percent change is " & Max_Percent_Change)
        ' MsgBox ("Max percent change row is " & MaxPercentRow)

        ' MsgBox ("Min percent change is " & Min_Percent_Change)
        ' MsgBox ("Min percent change row is " & MinPercentRow)

        ' MsgBox ("Max volume is " & Max_Volume)
        ' MsgBox ("Min volume row is " & MaxVolumeRow)

    Next worksheet_index


End Sub