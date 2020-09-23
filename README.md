<div align="center">

## crystal report date range


</div>

### Description

Print a crystal report in a certain date range
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jimmy\_](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jimmy.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jimmy-crystal-report-date-range__1-14757/archive/master.zip)





### Source Code

```
report.ReportFileName = gvPath & "\cheques.rpt"
report.CopiesToPrinter = InputBox("How many copies would you like to print")
report.SelectionFormula = "{cheques.date} in Date (" & Format$(Startdatetextbox.Value, "yyyy,mm,dd") & ") to Date (" & Format$(enddatetextbox.Value, "yyyy,mm,dd") & ")"
report.ReportTitle = "Report between" & " " & Format$(Startdatetextbox.Value, "long date") & " " & "and" & " " & Format(enddatetextbox.Value, "long date")
report.Action = 1
```

