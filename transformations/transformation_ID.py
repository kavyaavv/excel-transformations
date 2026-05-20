#let
    #Source = Excel.Workbook(Parameter1, null, true),
    #Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Renamed Columns" = Table.RenameColumns(#"Promoted Headers",{{"Account", "Customer No"}, {"Name 1", "Customer NM"}, {"Amount in local currency", "AR Balance"}, {"Company Code", "Country"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"AR Balance", Int64.Type}, {"Closing Date", type date}, {"Customer No", type text}, {"Customer NM", type text}, {"Country", type text}, {"Net due date", type date}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type","ID11","Indonesia",Replacer.ReplaceText,{"Country"}),
    #"Added Custom" = Table.AddColumn(#"Replaced Value", "Days Left", each [Closing Date]-[Net due date]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"Days Left", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Not Due", each if [Days Left] > 1 then [AR Balance] else 0,Int64.Type),
    #"Added Custom1" = Table.AddColumn(#"Added Conditional Column", "AR Overdue", each [AR Balance]-[Not Due],Int64.Type),
    #"Calculated Start of Month" = Table.TransformColumns(#"Added Custom1",{{"Closing Date", Date.StartOfMonth, type date}}),
    #"Filtered Rows" = Table.SelectRows(#"Calculated Start of Month", each ([Country] = "Indonesia")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Country", "Customer No", "Customer NM", "AR Balance", "Closing Date", "AR Overdue"})
#in
    #"Removed Other Columns"
    
# Transformation Name: AR Overdue ID

