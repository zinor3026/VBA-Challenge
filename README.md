# VBA-Challenge
Module 2 Challenge
Code sources:
  1. Code provided by tutor(Simon Rennocks):
            start = 0
            year_change = ws.Cells(i, 6) - ws.Cells(start, 3)
            percent_change = (year_change / ws.Cells(start, 3))
            start = i +1
            Vol_index = WorksheetFunction.Match(WorksheetFunction.Max(Max_range), Max_range, 0)
             ws.Range("Q4").Value = ws.Cells(Vol_index + 1, 10)
  
