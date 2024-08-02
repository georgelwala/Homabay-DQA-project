let
    Source = Excel.Workbook(File.Contents("C:\Users\GEORGE OMONDI\Desktop\GitHub\Homabay-DQA-project\RCH DQA.xlsx"), null, true),
   #"Score for pbix_Sheet" = Source{[Item="Score for pbix",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(#"Score for pbix_Sheet",{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Changed Type", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Service", type text}, {"Indicators", type text}, {"Comparison", type text}, {"Oct-23", Int64.Type}, {"Nov-23", Int64.Type}, {"Dec-23", Int64.Type}}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Changed Type1", {"Service", "Indicators", "Comparison"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Date"}, {"Value", "Score"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns", "Date", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Date.1", "Date.2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Date.1", type text}, {"Date.2", type text}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type2", "Date", each Text.Combine({"01-",[Date.1],"-20",[Date.2]})),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom",{{"Date", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type3",{"Date.1", "Date.2"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Date", "Service", "Indicators", "Comparison", "Score"}),
    #"Added Conditional Column" = Table.AddColumn(#"Reordered Columns", "Facility", each if [Indicators] <> "" then "Rachuonyo DH" else null),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Conditional Column",{"Date", "Facility", "Service", "Indicators", "Comparison", "Score"})
in
    #"Reordered Columns1"