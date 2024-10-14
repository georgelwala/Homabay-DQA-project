let
  Source = Excel.Workbook(
    File.Contents(
      "C:\Users\GEORGE OMONDI\Desktop\GitHub\Homabay-DQA-project\Facility folder\Rachuonyo South\Rachuonyo South score.xlsx"
    ), 
    null, 
    true
  ), 
  nyatike_dqa_variance_Table = Source{[Item = "RACHUONYO_SOUTH", Kind = "Table"]}[Data], 
  #"Changed Type" = Table.TransformColumnTypes(
    nyatike_dqa_variance_Table, 
    {
      {"Service", type text}, 
      {"Indicators", type text}, 
      {"Comparison", type text}, 
      {"Jan-24", type number}, 
      {"Feb-24", type number}, 
      {"Mar-24", type number}
    }
  ), 
  #"Changed Type1" = Table.TransformColumnTypes(
    #"Changed Type", 
    {{"Jan-24", Percentage.Type}, {"Feb-24", Percentage.Type}, {"Mar-24", Percentage.Type}}
  ), 
  #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(
    #"Changed Type1", 
    {"Service", "Facility", "Indicators", "Comparison"}, 
    "Attribute", 
    "Value"
  ), 
  #"Renamed Columns" = Table.RenameColumns(
    #"Unpivoted Other Columns", 
    {{"Attribute", "Date"}, {"Value", "Variance"}}
  ), 
  #"Split Column by Delimiter" = Table.SplitColumn(
    #"Renamed Columns", 
    "Date", 
    Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), 
    {"Date.1", "Date.2"}
  ), 
  #"Changed Type2" = Table.TransformColumnTypes(
    #"Split Column by Delimiter", 
    {{"Date.1", type text}, {"Date.2", type text}}
  ), 
  #"Added Custom" = Table.AddColumn(
    #"Changed Type2", 
    "Date", 
    each Text.Combine({"01-", [Date.1], "-20", [Date.2]})
  ), 
  #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom", {{"Date", type date}}), 
  #"Removed Columns" = Table.RemoveColumns(#"Changed Type3", {"Date.1", "Date.2"}), 
  #"Reordered Columns" = Table.ReorderColumns(
    #"Removed Columns", 
    {"Date", "Service", "Facility", "Indicators", "Comparison", "Variance"}
  ), 
  #"Added Conditional Column1" = Table.AddColumn(
    #"Reordered Columns", 
    "Subcounty", 
    each if [Indicators] <> "" then "RACHUONYO SOUTH" else null
  ), 
  #"Reordered Columns1" = Table.ReorderColumns(
    #"Added Conditional Column1", 
    {"Date", "Subcounty", "Facility", "Service", "Indicators", "Comparison", "Variance"}
  ), 
  #"Added Custom1" = Table.AddColumn(
    #"Reordered Columns1", 
    "Score (Numerator)", 
    each if [Variance] >= - 0.05 and [Variance] <= 0.05 then 1 else 0
  ), 
  #"Added Conditional Column2" = Table.AddColumn(
    #"Added Custom1", 
    "Denom", 
    each if [#"Score (Numerator)"] <> null then 1 else null
  ),
    #"Renamed Columns1" = Table.RenameColumns(#"Added Conditional Column2",{{"Score (Numerator)", "Score"}}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Renamed Columns1",{{"Score", Int64.Type}, {"Denom", Int64.Type}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type4","PARTOGRAPHS","Partograph Accuracy",Replacer.ReplaceText,{"Comparison"})
in
  #"Replaced Value"